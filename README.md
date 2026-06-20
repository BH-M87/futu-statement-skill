# futu-statement-skill

*English | [中文](README.zh-CN.md)*

A Claude/agent **skill** + standalone Python tool that parses **Futu (富途證券) Hong Kong
statements** — the **annual statement (`年度账单` .xlsx)** or the **monthly e-statements
(`證券月結單` .pdf)** — into tax-ready CSVs, for **individual income tax on foreign income
(个税 境外所得)**, CRS reporting, and broker P&L reconciliation.

It reads **official statement files only** — it does **not** use the Futu OpenD / API (which
exposes no realized P&L, no fees, and unlabeled dividends).

## Why

Futu's HK statements are authoritative for tax, but:

- there is **no realized-P&L section** — it must be reconstructed from the trade list;
- there is **no "dividend" section** — cash dividends hide inside **`公司行動` (corporate actions)**;
- interest, deposits and withdrawals are mixed together inside **`資金進出`**.

This repo captures the method (see [`SKILL.md`](SKILL.md)) and ships a parser that does it.

## Input: two formats, auto-detected (xlsx preferred)

| Input | Where to get it | Notes |
|---|---|---|
| **`年度账单` .xlsx** (preferred) | Futu app → 我的 → 賬戶詳情 → 年度账单 | Cleaner; has **期初/期末 holdings**, so cross-year positions realize correctly |
| **`月結單` .pdf** (one or many) | Futu app → 我的 → 賬戶詳情 → 電子結單/月結單 | Needs `pdftotext`; no opening positions |

Point the tool at a file or a folder. **If a folder has both, the xlsx wins** — it's more
accurate: the monthly PDFs only cover the calendar year, so a position opened last December
and closed in January is invisible, whereas the annual xlsx seeds it from `期初` (opening)
holdings. (In one real case the two methods differed by exactly one year-boundary option's premium.)

## Install

```bash
pip install openpyxl          # for the annual .xlsx
brew install poppler          # for the monthly .pdf (or: apt install poppler-utils) → pdftotext
```

Install only what your input needs. Tested on Python 3.10+.

## Usage

```bash
python3 parse_futu_statement.py 2025_年度账单.xlsx -o out/ --rate 0.90322   # annual xlsx
python3 parse_futu_statement.py /folder-of-pdfs   -o out/ --rate 0.90322   # monthly PDFs
python3 parse_futu_statement.py /folder           -o out/ --rate 0.90322   # both -> xlsx wins
```

`--rate` is optional — the HKD→RMB year-end mid-rate (中间价); when given, an RMB column is
added. The year is auto-detected.

### Outputs (UTF-8-BOM, Excel-friendly)

| File | Contents |
|---|---|
| `futu_<YEAR>_成交明细.csv` | stock/option trades; `变动金额` is already net of fees |
| `futu_<YEAR>_股息利息现金流.csv` | dividends (`公司行动`), interest, deposits/withdrawals, fees |
| `futu_<YEAR>_期权行权到期.csv` | option expiry (EXP) / assignment (ASS) events |
| `futu_<YEAR>_已实现盈亏_按标的.csv` | realized P&L per instrument (average-cost) |
| `futu_<YEAR>_账户净值.csv` | opening / closing balances (cross-check) |
| `futu_<YEAR>_税务汇总.csv` | tax summary — capital gains / dividends / interest + tax due (needs `--rate`) |

The script prints `Σ变动金额`, total dividends, and realized P&L so you can sanity-check.

## How realized P&L is computed

Average-cost per instrument, using Futu's explicit `開倉/平倉` (open/close) labels (buy vs
sell from the sign of `变动金额`, so it also handles `强平` forced-liquidation rows). Options
that expired or were assigned are closed at zero (premium fully realized); positions still
held at year-end keep realized = 0 (their gain is unrealized and not taxable for the year).
With the annual xlsx, `期初` holdings seed carried-in positions so cross-year items are
captured. Cross-check against `ΔNAV − net deposits`. Details in [`SKILL.md`](SKILL.md).

## Use as a Claude Code skill

Drop this folder into your skills directory (e.g. `~/.claude/skills/parsing-futu-statements/`)
or install via your plugin manager. Claude loads [`SKILL.md`](SKILL.md) when you ask it to
parse Futu statements or compute Futu P&L / dividends for tax.

## Privacy

This repo contains **no personal data**. The parser reads whatever file you point it at and
writes CSVs locally. `.gitignore` blocks `*.pdf`, `*.xlsx` and `*.csv` so you can't accidentally
commit statements or results.

## Reliability

Validated against a full year via both inputs: the monthly PDFs and the annual xlsx reconcile
exactly on every trade, dividend, interest and per-instrument figure — and the xlsx
additionally catches a cross-year option the monthly-only view misses. Trade numbers and
codes are exact; option names are decoded from the contract code, stock names harvested from
the holdings data.

## License

MIT
