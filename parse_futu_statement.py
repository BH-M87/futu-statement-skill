#!/usr/bin/env python3
"""Parse Futu (富途證券) HK monthly e-statements (證券月結單 PDF) into tax-ready CSVs.

Input is ONLY the official Futu monthly statement PDFs — no Futu OpenD / API data is
used. Get the statements from the Futu app: 我的 → 賬戶詳情 → 電子結單 / 月結單.

Why this exists: Futu's HK monthly statements have NO "realized P&L" section and NO
section literally named "dividend". Dividends hide inside 公司行動 (corporate actions),
interest hides inside 資金進出, and realized P&L must be reconstructed from the trades.
This tool extracts those facts and computes per-instrument realized P&L for tax /
accounting (个税 境外所得 / CRS / P&L reconciliation).

Usage:
    python3 parse_futu_statement.py STATEMENT.pdf [MORE.pdf ...] -o OUTDIR
    python3 parse_futu_statement.py /path/to/folder -o OUTDIR          # all *.pdf in folder
    python3 parse_futu_statement.py /folder -o OUTDIR --rate 0.90322   # add an RMB column

Outputs (CSV, utf-8-sig so Excel opens them without mojibake; YEAR is auto-detected):
    futu_<YEAR>_成交明细.csv          per-fill stock/option trades; 变动金额 already net of fees
    futu_<YEAR>_股息利息现金流.csv      dividends(公司行动)/interest/deposits/withdrawals/fees
    futu_<YEAR>_期权行权到期.csv        option expiry(EXP) / assignment(ASS) events
    futu_<YEAR>_已实现盈亏_按标的.csv    realized P&L per instrument (average-cost)
    futu_<YEAR>_账户净值.csv           per-statement opening/closing net asset value (cross-check)

Requires: pdftotext (poppler).  Install: brew install poppler  /  apt install poppler-utils
No personal data is embedded; everything is read from the PDFs you pass in.
"""
from __future__ import annotations
import argparse, csv, glob, os, re, subprocess, sys
from collections import Counter

# Only the explicit open/close directions appear in Futu statements.
DIRECTIONS = ("買入開倉", "賣出平倉", "賣出開倉", "買入平倉")
SKIP_FRAG = ("佣金", "平台使用費", "交易系統使用費", "交收費", "印花稅", "交易費",
             "證監會徵費", "財匯局徵費", "小計", "合計")
# tax-relevant cashflow types (exclude internal money-market fund parking 基金申購/贖回)
TAX_CASH_TYPES = ("公司行動", "出入金", "證券月度利息扣除", "月度利息扣除",
                  "融券利息", "首次稅局登記費", "稅局登記費", "月度利息")
FILL_RE = re.compile(
    r"(SEHK|NASDAQ|NYSE|ARCA|AMEX|US)\s+(HKD|USD|CNH|JPY|SGD)\s+"
    r"(\d{4}/\d{2}/\d{2})\s+(\d{4}/\d{2}/\d{2})\s+"
    r"([\d,]+)\s+([\d.]+)\s+([\d,]+\.\d{2})\s+([+-]?[\d,]+\.\d{2})")
CASHFLOW_RE = re.compile(
    r"\s*(\d{4}/\d{2}/\d{2})\s+(增加|減少)\s+(\S+?)\s+(HKD|USD|CNH|JPY|SGD)\s+"
    r"([+-]?[\d,]+\.\d{2})\s*(.*)$")
OPT_EVENT_RE = re.compile(r"Opt (EXP|ASS)-[A-Z]+-([A-Z0-9]+)-\d{8}")
STOCK_CODE_RE = re.compile(r"([0-9]{4,5}|[A-Z]{2,4}\d{6}[CP]\d+)\(([^)]*)\)?")
NUMERIC = re.compile(r"^(HKD|USD|CNH|JPY|SGD|[\d,]+\.?\d*)$")
OPT_CODE_RE = re.compile(r"^([A-Z]{2,4})(\d{6})([CP])(\d+)$")
# HK option root -> underlying name; unknown roots fall back to the root code itself.
UNDERLYING = {"MIU": "小米", "TCH": "腾讯", "BAB": "阿里巴巴", "JDC": "京东", "MET": "美团"}


def _f(s): return float(s.replace(",", ""))


def nice_name(code, fallback=""):
    """Decode an option code to a readable contract name; pass through stocks."""
    m = OPT_CODE_RE.match(code or "")
    if not m:
        return fallback
    root, date, cp, strike = m.groups()
    return f"{UNDERLYING.get(root, root)} {date} {int(strike) / 1000:.2f} {'沽' if cp == 'P' else '购'}"


def pdf_to_text(path):
    out = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True)
    if out.returncode != 0:
        raise RuntimeError(f"pdftotext failed on {path}: {out.stderr}")
    return out.stdout


def parse_one(text):
    """Return (trades, cashflows, opt_events, opening_nav, closing_nav) for one statement."""
    trades, cashflows, opt_raw = [], [], []
    asset_opt = {}                              # code -> {name, contracts} from 資產進出
    opening_nav = closing_nav = None
    sec = None
    cur_dir = cur_code = cur_name = ""
    building = False

    for raw in text.splitlines():
        ln, s = raw, raw.strip()

        if "交易-股票和股票期權" in ln:
            sec = "trades"; cur_dir = cur_code = cur_name = ""; building = False; continue
        if "資金進出" in ln:
            sec = "cash"; continue
        if "資產進出" in ln:
            sec = "asset"; continue
        if any(h in ln for h in ("交易-基金", "融資總覽", "重要提示", "期末概覽")):
            if sec == "trades": sec = None

        if "資產淨值" in ln and opening_nav is None:
            nums = re.findall(r"[-\d,]+\.\d{2}", ln)
            if len(nums) >= 2:
                opening_nav, closing_nav = _f(nums[0]), _f(nums[1])

        # ----- trades -----
        if sec == "trades":
            hdr = s.lstrip("*").strip()           # '*' = forced-liquidation marker
            is_header = False
            for d in DIRECTIONS:
                if hdr.startswith(d):
                    cur_dir = d
                    rest = hdr[len(d):].split()
                    first = rest[0] if rest else ""
                    cur_code = "" if (not first or NUMERIC.match(first)) else first
                    cur_name = ""
                    building = bool(cur_code)      # option header carries a code fragment
                    is_header = True
                    break
            if is_header:
                continue

            m = FILL_RE.search(ln)
            if m:
                mkt, ccy, td, sd, q, p, g, net = m.groups()
                before = ln[:ln.find(mkt)]
                pm = STOCK_CODE_RE.search(before)
                if pm and not building:           # stock fill: code printed on the line
                    cur_code, cur_name = pm.group(1), pm.group(2).strip()
                elif "(" in cur_code:
                    cur_code = cur_code.split("(")[0]
                gross, netv = _f(g), _f(net)
                fee = round(abs(abs(netv) - gross), 2)
                trades.append([td, sd, cur_code, cur_name, cur_dir, ccy,
                               int(_f(q)), float(p), gross, fee, netv])
                building = False
            elif (building and "SEHK" not in ln and s
                  and not any(x in s for x in SKIP_FRAG)):
                if "(" in s:
                    pre, post = s.split("(", 1)
                    cur_code += pre.strip(); cur_name = post.rstrip(")").strip(); building = False
                else:
                    cur_code += s.strip()

        # ----- 資產進出: option name + contract count (event date-suffix wraps, so grab here) -----
        if sec == "asset":
            am = re.search(r"([A-Z]{2,4}\d{6}[CP]\d+)\(([^)]*)", ln)
            if am:
                cm = re.search(r"\+(\d+)\b", ln)
                asset_opt[am.group(1)] = {"name": am.group(2).strip(),
                                          "contracts": cm.group(1) if cm else ""}

        # ----- cash flow -----
        if sec == "cash":
            m = CASHFLOW_RE.match(ln)
            if m:
                date, _d, typ, ccy, amt, remark = m.groups()
                cashflows.append([date, typ, ccy, _f(amt), remark.strip()])

        # ----- option expiry / assignment (appears in both 資金進出 and 資產進出) -----
        em = OPT_EVENT_RE.search(ln)
        if em:
            dm = re.match(r"\s*(\d{4}/\d{2}/\d{2})", ln)
            nm = re.search(r"\(([^)]*)\)?", ln[ln.find(em.group(2)):]) if "(" in ln else None
            qm = re.search(r"\+(\d+)", ln)
            opt_raw.append({"date": dm.group(1) if dm else "", "code": em.group(2),
                            "kind": em.group(1), "name": nm.group(1).strip() if nm else "",
                            "contracts": qm.group(1) if qm else ""})

    # merge duplicate option events by (code, kind), keeping best date/name/contracts
    merged = {}
    for e in opt_raw:
        k = (e["code"], e["kind"])
        m0 = merged.setdefault(k, {"date": "", "code": e["code"], "kind": e["kind"],
                                   "name": "", "contracts": ""})
        for f in ("date", "name", "contracts"):
            if not m0[f] and e[f]:
                m0[f] = e[f]
    for ev in merged.values():                 # enrich from 資產進出 capture
        a = asset_opt.get(ev["code"])
        if a:
            ev["name"] = ev["name"] or a["name"]
            ev["contracts"] = ev["contracts"] or a["contracts"]
    return trades, cashflows, list(merged.values()), opening_nav, closing_nav


def realized_by_ticker(trades, opt_events):
    """Average-cost realized P&L per instrument, PDF-only.

    Uses Futu's explicit 開倉/平倉 labels. Per unit, realized = (opening signed cashflow
    + closing signed cashflow). Options that expired/were assigned (in opt_events) but
    have no buy/sell-to-close are closed synthetically at 0 (premium fully realized).
    Positions still genuinely held at year-end (e.g. assignment shares not yet sold) keep
    realized = 0; their cost stays unrealized and is NOT taxable for the year.
    """
    book = {}
    ev_codes = {e["code"] for e in opt_events}
    for t in sorted(trades, key=lambda r: r[0]):          # chronological by trade_date
        code, name, d, q, net = t[2], t[3], t[4], float(t[6]), float(t[10])
        b = book.setdefault(code, {"name": "", "pos": 0.0, "avg": 0.0,
                                   "realized": 0.0, "nfills": 0, "sum_net": 0.0})
        if name and not b["name"]:
            b["name"] = name
        b["nfills"] += 1; b["sum_net"] += net
        if "開倉" in d:                                    # OPEN
            prev = abs(b["pos"]); tot = b["avg"] * prev + net
            b["pos"] += q if "買入" in d else -q
            b["avg"] = tot / abs(b["pos"]) if abs(b["pos"]) > 1e-9 else 0.0
        else:                                             # CLOSE (平倉)
            b["realized"] += (net / q + b["avg"]) * q
            b["pos"] += q if "買入" in d else -q
            if abs(b["pos"]) < 1e-9:
                b["avg"] = 0.0
    # synthetic close for expired / assigned options still open
    for code, b in book.items():
        if abs(b["pos"]) > 1e-9 and code in ev_codes:
            b["realized"] += b["avg"] * abs(b["pos"]); b["pos"] = 0.0
    return book


def detect_year(trades):
    yrs = Counter(t[0][:4] for t in trades if t[0])
    return yrs.most_common(1)[0][0] if yrs else "YEAR"


def main(argv=None):
    ap = argparse.ArgumentParser(description="Parse Futu HK monthly statement PDFs into tax CSVs")
    ap.add_argument("inputs", nargs="+", help="PDF file(s) or a folder containing *.pdf")
    ap.add_argument("-o", "--outdir", default="futu_parsed", help="output directory")
    ap.add_argument("--rate", type=float, default=None,
                    help="HKD->RMB rate (e.g. year-end 中间价 0.90322); adds an RMB column")
    args = ap.parse_args(argv)

    pdfs = []
    for p in args.inputs:
        pdfs += sorted(glob.glob(os.path.join(p, "*.pdf"))) if os.path.isdir(p) else [p]
    if not pdfs:
        print("no PDFs found", file=sys.stderr); return 2
    os.makedirs(args.outdir, exist_ok=True)

    trades, cashflows, opt_events, summ = [], [], [], []
    names_map = {}                              # code -> full name, harvested from PDFs
    full_name_re = re.compile(r"([0-9]{4,5}|[A-Z]{2,4}\d{6}[CP]\d+)\(([^)\n]{1,24})\)")
    for pdf in pdfs:
        text = pdf_to_text(pdf)
        for nm in full_name_re.finditer(text):
            names_map.setdefault(nm.group(1), nm.group(2).strip())
        tr, cf, ev, onav, cnav = parse_one(text)
        trades += tr; cashflows += cf; opt_events += ev
        summ.append([os.path.basename(pdf), onav, cnav])

    def disp_name(code, fallback=""):           # option -> decode; stock -> harvested name
        return nice_name(code, names_map.get(code, fallback))
    # merge option events across statements by (code, kind)
    ev_merged = {}
    for e in opt_events:
        ev_merged.setdefault((e["code"], e["kind"]), e)
    opt_events = list(ev_merged.values())

    year = detect_year(trades)
    rate = args.rate

    def out(name): return os.path.join(args.outdir, name)

    def write(path, header, rows):
        with open(path, "w", newline="", encoding="utf-8-sig") as fp:
            w = csv.writer(fp); w.writerow(header); [w.writerow(r) for r in rows]

    # 1) 成交明细
    tr_sorted = [[r[0], r[1], r[2], disp_name(r[2], r[3])] + r[4:] for r in sorted(trades)]
    write(out(f"futu_{year}_成交明细.csv"),
          ["成交日期", "交收日期", "代码", "名称", "买卖方向", "货币",
           "数量", "价格", "成交金额", "手续费", "变动金额(净额)"],
          tr_sorted + [["合计", "", "", "", "", "", sum(r[6] for r in trades), "", "",
                        round(sum(r[9] for r in trades), 2),
                        round(sum(r[10] for r in trades), 2)]])

    # 2) 股息利息现金流 (tax-relevant types only)
    cf_tax = sorted([c for c in cashflows if c[1] in TAX_CASH_TYPES])
    write(out(f"futu_{year}_股息利息现金流.csv"),
          ["日期", "类型", "货币", "金额", "备注"], cf_tax)

    # 3) 期权行权到期
    write(out(f"futu_{year}_期权行权到期.csv"),
          ["日期", "类型", "代码", "名称", "合约数", "说明"],
          [[e["date"], "到期(EXP)" if e["kind"] == "EXP" else "行权(ASS)",
            e["code"], disp_name(e["code"], e["name"]), e["contracts"],
            "到期作废,保费为已实现收益" if e["kind"] == "EXP"
            else "被行权->对应正股买入见成交明细"]
           for e in sorted(opt_events, key=lambda e: e["date"])])

    # 4) 已实现盈亏 按标的
    book = realized_by_ticker(trades, opt_events)
    rows, total = [], 0.0
    for code in sorted(book, key=lambda c: book[c]["realized"]):
        b = book[code]; held = abs(b["pos"]) > 1e-9
        realized = 0.0 if held else round(b["realized"], 2)
        total += realized
        note = "年末仍持有,未实现部分不计入(本年已实现=0)" if held else ""
        row = [code, disp_name(code, b["name"]), b["nfills"], round(b["sum_net"], 2), realized]
        if rate is not None:
            row.append(round(realized * rate, 2))
        row.append(note)
        rows.append(row)
    header4 = ["代码", "名称", "成交笔数", "成交净额合计(HKD)", "已实现盈亏(HKD)"]
    if rate is not None:
        header4.append(f"已实现盈亏(RMB,×{rate})")
    header4.append("备注")
    tot_row = ["合计", "", "", "", round(total, 2)]
    if rate is not None:
        tot_row.append(round(total * rate, 2))
    tot_row.append("已实现合计;年末持有标的不计入")
    write(out(f"futu_{year}_已实现盈亏_按标的.csv"), header4, rows + [tot_row])

    # 5) 账户净值 (NAV cross-check)
    write(out(f"futu_{year}_账户净值.csv"),
          ["结单", "期初净值(HKD)", "期末净值(HKD)"], summ)

    div = sum(c[3] for c in cashflows if c[1] == "公司行動" and c[3] > 0)
    print(f"parsed {len(pdfs)} statement(s), year={year} -> {args.outdir}/")
    print(f"  成交明细:        {len(trades)} fills, Σ变动金额={sum(r[10] for r in trades):,.2f}")
    print(f"  股息利息现金流:  {len(cf_tax)} rows, 股息(公司行动+)={div:,.2f}")
    print(f"  期权行权到期:    {len(opt_events)} events")
    print(f"  已实现盈亏:      Σ={total:,.2f} HKD" + (f" = RMB {total*rate:,.2f}" if rate else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
