# futu-statement-skill

*English | [中文](README.zh-CN.md)*

A Claude/agent **skill** + standalone Python tool that parses **Futu (富途證券) Hong Kong
monthly e-statements** (`證券月結單` PDFs) into tax-ready CSVs — built for **individual income
tax on foreign income (个税 境外所得)**, CRS reporting, and broker P&L reconciliation.

**Input is PDF-only.** It does **not** use the Futu OpenD / API (which exposes no realized
P&L, no fees, and unlabeled dividends). The statement PDFs are the source of truth.

## Why

Futu's HK monthly statements are authoritative for tax, but:

- there is **no realized-P&L section** — it must be reconstructed from the trade list;
- there is **no "dividend" section** — cash dividends hide inside **`公司行動` (corporate actions)**;
- interest, deposits and withdrawals are mixed together inside **`資金進出`**.

This repo captures the method (see [`SKILL.md`](SKILL.md)) and ships a parser that does it.

## Install

```bash
brew install poppler          # macOS  (or: apt install poppler-utils)  → provides pdftotext
```

No Python packages required (standard library only). Tested on Python 3.10+.

## Usage

```bash
# point it at one statement, or a whole folder of them
python3 parse_futu_statement.py /path/to/statements/ -o out/ --rate 0.90322
```

`--rate` is optional — the HKD→RMB year-end mid-rate (中间价); when given, an RMB column is
added. The year is auto-detected from the statements.

### Outputs (UTF-8-BOM, Excel-friendly)

| File | Contents |
|---|---|
| `futu_<YEAR>_成交明细.csv` | per-fill stock/option trades; `变动金额` is already net of fees |
| `futu_<YEAR>_股息利息现金流.csv` | dividends (`公司行动`), interest, deposits/withdrawals, fees |
| `futu_<YEAR>_期权行权到期.csv` | option expiry (EXP) / assignment (ASS) events |
| `futu_<YEAR>_已实现盈亏_按标的.csv` | realized P&L per instrument (average-cost) |
| `futu_<YEAR>_账户净值.csv` | per-statement opening / closing net asset value (cross-check) |

The script prints `Σ变动金额`, total dividends, and realized P&L so you can sanity-check.

## How realized P&L is computed

Average-cost per instrument, using Futu's explicit `開倉/平倉` (open/close) labels. Options
that expired or were assigned are closed at zero (premium fully realized); positions still
held at year-end keep realized = 0 (their gain is unrealized and not taxable for the year).
Cross-check against `ΔNAV − net deposits`. Details in [`SKILL.md`](SKILL.md).

## Use as a Claude Code skill

Drop this folder into your skills directory (e.g. `~/.claude/skills/parsing-futu-statements/`)
or install via your plugin manager. Claude loads [`SKILL.md`](SKILL.md) when you ask it to
parse Futu statements or compute Futu P&L / dividends for tax.

## Privacy

This repo contains **no personal data**. The parser reads whatever PDFs you point it at and
writes CSVs locally. `.gitignore` blocks `*.pdf` and `*.csv` so you can't accidentally commit
statements or results.

## Reliability

Validated against a full year (12 monthly statements): every trade number, all instrument
codes, and all per-instrument realized P&L figures reconcile exactly. Trade numbers and stock
codes are exact; option names are decoded from the contract code, stock names harvested from
the holdings sections.

## License

MIT
