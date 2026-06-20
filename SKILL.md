---
name: parsing-futu-statements
description: Use when extracting trades, dividends, interest, fees, or realized P&L from Futu (富途) HK monthly e-statement PDFs (證券月結單) — for individual income tax on foreign income (个税 境外所得), CRS reporting, accounting, or broker P&L reconciliation. Covers where dividends and realized P&L actually live (they are not where you expect) and how to reconstruct figures the statement does not compute. PDF-only; does not use the Futu OpenD/API.
---

# Parsing Futu Statements

## Overview

Futu (富途證券) HK monthly e-statements (`保證金綜合帳戶 證券月結單`) are the authoritative
record for tax/accounting, but two things people look for **are not there**:

- There is **no "已實現盈虧" (realized P&L) section** — reconstruct it from the trade list.
- There is **no section named "派息/股息" (dividend)** — cash dividends are buried inside
  **`公司行動` (corporate actions)** in the `資金進出` section.

**Input is PDF-only.** Do not use the Futu OpenD trading API: it exposes no realized P&L,
no fees, no labeled dividends (cash-flow types come back as "其他"), and only single-day
cash-flow queries. The statement PDFs are the source of truth. Get them from the Futu app:
`我的 → 賬戶詳情 → 電子結單 / 月結單`.

`parse_futu_statement.py` (this skill) turns a folder of statement PDFs into tax-ready CSVs.

## When to Use

- Computing 个税 境外所得: realized capital gains + dividends, kept separate
- CRS / annual account reconciliation
- "How much did I make/lose on Futu last year, and how much was dividends?"

Not for: real-time positions/quotes; US-broker statements (different layout).

## Quick Start

```bash
brew install poppler                       # or: apt install poppler-utils  (gives pdftotext)
python3 parse_futu_statement.py /folder-of-pdfs -o out/ --rate 0.90322
```

`--rate` (optional) is the HKD→RMB year-end 中间价; it adds an RMB column. Outputs
(`YEAR` auto-detected; utf-8-sig for Excel):

| File | Contents |
|---|---|
| `futu_<YEAR>_成交明细.csv` | per-fill trades; `变动金额` already net of fees |
| `futu_<YEAR>_股息利息现金流.csv` | dividends (公司行動), interest, deposits/withdrawals, fees |
| `futu_<YEAR>_期权行权到期.csv` | option expiry (EXP) / assignment (ASS) events |
| `futu_<YEAR>_已实现盈亏_按标的.csv` | realized P&L per instrument (average-cost) |
| `futu_<YEAR>_账户净值.csv` | per-statement opening/closing NAV (cross-check) |

The script prints `Σ变动金额`, dividends, and realized total so you can sanity-check.

## Statement Structure

| Section | Holds | Use for |
|---|---|---|
| `資產組合摘要` | 期初/期末 資產淨值 (opening/closing NAV) | NAV-change cross-check |
| `交易-股票和股票期權` | per-fill stock & option trades | realized P&L, fees |
| `資金進出` | dividends (`公司行動`), interest, deposits/withdrawals | dividend & interest income |
| `資產進出` | option expiry/assignment, share transfers | option event qty + name |
| `期末概覽` | year-end holdings (qty, market value, full names) | what is still open (unrealized) |
| `融資總覽` | daily margin balance & rate | interest derivation |

## Critical Gotchas (each cost real debugging)

1. **Dividends live in `公司行動`, not a dividend section.** Remark codes: `F/D` = final
   dividend, `I/D` = interim. A row like `<date> F/D-HKD<rate>/SH <SEHK NNNN NAME> <qty> shares`
   is a cash dividend of `rate × qty`. Separate `Scrip Charge` / `Handling Charge` rows reduce the net.
2. **`變動金額` (net amount) is already net of fees.** The fee line (`佣金 / 平台使用費 /
   交易系統使用費 / 印花稅 / 交收費 / 證監會徵費 / 財匯局徵費 → 小計`) is informational.
   Per-fill fee = `|net| − gross`.
3. **Direction encodes open/close & long/short:** `買入開倉` (buy-to-open), `賣出平倉`
   (sell-to-close), `賣出開倉` (sell-to-open / short), `買入平倉` (buy-to-close).
4. **A leading `*` marks a forced-liquidation trade** (`*買入平倉 …`). Strip it before
   matching the direction, or the row inherits the previous trade's instrument code (this
   silently mis-attributed option fills to a stock in early versions).
5. **Option assignment is split:** an `Opt ASS` event in `資金進出/資產進出` (0 cash) **plus**
   a normal stock **BUY at the strike** in `交易` (this carries the cash). Option **expiry**
   (`Opt EXP`) has no closing trade — the `賣出開倉` premium is the realized gain.
6. **Interest is a cost, not income.** `月度利息扣除` / `證券月度利息扣除` (margin) and `融券利息`
   (stock-borrow for shorts) appear in `資金進出` as negatives.
7. **External vs internal cash:** `出入金` = real deposits/withdrawals. `基金申購/贖回` =
   money parked in/out of a money-market fund (internal — NOT a deposit; exclude).
8. **`pdftotext -layout` is the only reliable extractor.** pdfplumber mis-handles the CJK
   trade tables. Trade numbers are always exact and stock codes are exact. Option codes can
   column-wrap, so option **names are decoded from the code** (format `<ROOT><yymmdd><C|P><strike>`,
   e.g. a `…P…` put → `<underlying> <date> <strike> 沽`) and stock **names are harvested** from
   the holdings sections.

## Reconstructing Realized P&L (the statement won't)

Average-cost per instrument, using the explicit `開倉/平倉` labels. Per unit,
`realized = opening signed cashflow + closing signed cashflow`. Special cases the script
handles:

- Options that **expired / were assigned** (in `期权行权到期`) but have no buy/sell-to-close
  are closed synthetically at 0 → the premium is fully realized.
- Positions **still held at year-end** (e.g. assignment shares not yet sold) keep
  realized = 0; their cost stays unrealized and is **not taxable for the year**.
- **Cross-check** with NAV: `ΔNAV (期末−期初 資產淨值) − net 出入金` ≈
  realized + unrealized + dividends + interest. Agreement within ~1–2% validates the number.

## Common Mistakes

| Mistake | Reality |
|---|---|
| "No dividend section → no dividends" | They're in `公司行動`. Always scan it. |
| Using ΔNAV as the taxable gain | ΔNAV includes unrealized + deposits. Isolate realized. |
| Counting `基金申購/贖回` as deposits | Internal money-fund parking, not external cash. |
| Treating interest as income | `月度利息扣除` is a financing cost you paid. |
| Re-summing fees onto net | `變動金額` already deducted them. |
| Ignoring the `*` forced-liquidation flag | It breaks direction parsing and mis-codes the fill. |

## Tax note (个税 境外所得)

Convert HKD to RMB at the year-end 人民币汇率中间价 (汇算清缴 口径; pass via `--rate`).
Capital gains (财产转让所得) and dividends (利息股息红利所得, flat 20%) are taxed separately —
keep them in separate files (this skill does). Same-market securities gains/losses generally
net within 财产转让所得. Confirm netting scope, foreign-tax credit, and FX 口径 with a tax advisor.
