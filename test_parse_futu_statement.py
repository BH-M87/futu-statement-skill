import csv
import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import parse_futu_statement as F


def trade_row(date, code, currency, net, qty=1, opening=True):
    # [成交日期, 交收日期, 代码, 名称, 方向, 货币, 数量, 价格, 成交金额, 手续费, 变动金额]
    if net < 0:                                  # cash out -> buy
        direction = "買入開倉" if opening else "買入平倉"
    else:                                        # cash in -> sell
        direction = "賣出平倉" if opening is False else "賣出開倉"
    return [date, "", code, "", direction, currency, qty, 0.0, abs(net), 0.0, net]


class CurrencyHandlingTest(unittest.TestCase):
    def test_realized_pnl_keeps_currency_books_separate(self):
        trades = [
            trade_row("2025/01/01", "ABC", "HKD", -100, opening=True),    # 買入開倉
            trade_row("2025/01/02", "ABC", "HKD", 110, opening=False),    # 賣出平倉
            trade_row("2025/01/03", "ABC", "USD", -10, opening=True),
            trade_row("2025/01/04", "ABC", "USD", 13, opening=False),
        ]

        book = F.realized_by_ticker(trades, [])

        self.assertEqual(round(book[("HKD", "ABC")]["realized"], 2), 10.0)
        self.assertEqual(round(book[("USD", "ABC")]["realized"], 2), 3.0)

    def _run_main(self, extra_args):
        trades = [
            trade_row("2025/01/01", "HK001", "HKD", -100, opening=True),
            trade_row("2025/01/03", "HK001", "HKD", 110, opening=False),
            trade_row("2025/02/01", "US001", "USD", -10, opening=True),
            trade_row("2025/02/04", "US001", "USD", 13, opening=False),
        ]
        cashflows = [
            ["2025/03/01", "公司行动", "HKD", 8.0, "dividend"],
            ["2025/03/02", "公司行动", "USD", 2.0, "dividend"],
        ]
        nav = (["时期类型", "类别", "货币", "金额(原币种)"], [])
        loaded = (trades, cashflows, [], {}, nav, {})

        outdir = tempfile.mkdtemp()
        with patch.object(F, "load_pdfs", return_value=loaded):
            with contextlib.redirect_stdout(io.StringIO()):
                F.main(["x.pdf", "-o", outdir, *extra_args])
        return Path(outdir)

    def test_realized_and_tax_group_by_currency_with_fx_rate(self):
        outdir = self._run_main(["--fx-rate", "HKD=0.9", "--fx-rate", "USD=7.1"])

        with (outdir / "futu_2025_已实现盈亏_按标的.csv").open(encoding="utf-8-sig", newline="") as fp:
            realized = {r["货币"]: float(r["已实现盈亏(原币)"]) for r in csv.DictReader(fp) if r["代码"] == "合计"}
        self.assertEqual(realized, {"HKD": 10.0, "USD": 3.0})

        with (outdir / "futu_2025_税务汇总.csv").open(encoding="utf-8-sig", newline="") as fp:
            tax_rows = list(csv.DictReader(fp))
        capital = [r for r in tax_rows if r["所得项目"] == "财产转让所得·已实现(本账户股票/期权)"]
        self.assertEqual(
            {(r["货币"], float(r["金额(原币)"]), float(r["金额(RMB)"])) for r in capital},
            {("HKD", 10.0, 9.0), ("USD", 3.0, 21.3)},
        )
        dividends = [r for r in tax_rows if r["所得项目"] == "利息股息红利所得·现金分红(毛额)"]
        self.assertEqual(
            {(r["货币"], float(r["金额(原币)"]), float(r["金额(RMB)"])) for r in dividends},
            {("HKD", 8.0, 7.2), ("USD", 2.0, 14.2)},
        )

    def test_partial_close_realized_is_reported_not_zeroed(self):
        # 开仓100股(成本-1000),年中平仓30股(+330,锁定盈利30),年末仍持有70股
        trades = [
            trade_row("2025/01/02", "HK700", "HKD", -1000, qty=100, opening=True),
            trade_row("2025/06/01", "HK700", "HKD", 330, qty=30, opening=False),
        ]
        nav = (["时期类型", "类别", "货币", "金额(原币种)"], [])
        loaded = (trades, [], [], {}, nav, {})

        outdir = tempfile.mkdtemp()
        with patch.object(F, "load_pdfs", return_value=loaded):
            with contextlib.redirect_stdout(io.StringIO()):
                F.main(["x.pdf", "-o", outdir, "--fx-rate", "HKD=0.9"])

        with (Path(outdir) / "futu_2025_已实现盈亏_按标的.csv").open(encoding="utf-8-sig", newline="") as fp:
            rows = list(csv.DictReader(fp))
        detail = next(r for r in rows if r["代码"] == "HK700")
        self.assertEqual(float(detail["已实现盈亏(原币)"]), 30.0)   # 已平仓部分,不再被清零
        self.assertIn("年末仍有持仓", detail["备注"])
        total = next(r for r in rows if r["代码"] == "合计")
        self.assertEqual(float(total["已实现盈亏(原币)"]), 30.0)

    def test_uses_2025_year_end_default_fx_rates_when_none_passed(self):
        outdir = self._run_main([])

        with (outdir / "futu_2025_税务汇总.csv").open(encoding="utf-8-sig", newline="") as fp:
            tax_rows = list(csv.DictReader(fp))
        capital = [r for r in tax_rows if r["所得项目"] == "财产转让所得·已实现(本账户股票/期权)"]
        self.assertEqual(
            {(r["货币"], float(r["金额(原币)"]), float(r["金额(RMB)"])) for r in capital},
            {("HKD", 10.0, 9.03), ("USD", 3.0, 21.09)},
        )


if __name__ == "__main__":
    unittest.main()
