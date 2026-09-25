"""Pruebas de ModelStyler.build_ratios: ventana UDM en hojas trimestrales.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import xlsx_builder as B  # noqa: E402

CANONS = ["rev", "cogs", "gross", "ebit", "ebt", "ni", "interest", "tax", "ta", "equity",
          "cash", "ar", "inv", "ap", "ca", "cl", "debt", "re", "cfo", "da"]
ROW = {k: n for n, k in enumerate(CANONS, start=10)}


def _build(sheet: str, quarterly: bool, **kw: object) -> dict[str, list[object]]:
    styler = B.ModelStyler()
    ws = styler.new_sheet(sheet)
    if quarterly:
        styler.quarter_header(ws, 3, 3, [f"{q}Q{y}{'A' if y < 2025 else 'E'}"
                                         for y in (2024, 2025) for q in (1, 2, 3, 4)])
    else:
        styler.period_header(ws, 3, 3, B.PeriodHeader(2020, 2027, 2023))
    ref = {k: f"{sheet}!{{c}}{n}" for k, n in ROW.items()}
    end, skipped = styler.build_ratios(ws, 40, 3, 8, ref, wacc_ref="WACC", **kw)
    assert not skipped, skipped
    return {str(ws.cell(row=r, column=2).value): [ws.cell(row=r, column=c).value
                                                  for c in range(3, 11)]
            for r in range(40, end)
            if ws.cell(row=r, column=2).value and ws.cell(row=r, column=1).value != "x"}


def _row(rows: dict[str, list[object]], prefix: str) -> list[object]:
    return next(v for k, v in rows.items() if k.startswith(prefix))


class BuildRatiosTests(unittest.TestCase):

    def test_quarterly_sheet_uses_ltm_flows_and_four_quarter_averages(self) -> None:
        rows = _build("Operating", quarterly=True)
        debt_ebitda = _row(rows, "Deuda / EBITDA")
        ebit = ",".join(f"Operating!{c}{ROW['ebit']}" for c in "CDEF")
        da = ",".join(f"Operating!{c}{ROW['da']}" for c in "CDEF")
        self.assertEqual(debt_ebitda[:3], [None, None, None])
        self.assertIn(f"SUM({ebit})", debt_ebitda[3])
        self.assertIn(f"SUM({da})", debt_ebitda[3])
        roe = _row(rows, "ROE DuPont 3")[3]
        self.assertIn("AVERAGE(" + ",".join(f"Operating!{c}{ROW['equity']}" for c in "CDEF")
                      + ")", roe)
        self.assertTrue(all(k.endswith("[UDM]") for k in rows))

    def test_ltm_days_ratios_use_days_year(self) -> None:
        self.assertIn("*DAYS_YEAR", _row(_build("Operating", quarterly=True), "DSO")[3])

    def test_days_quarter_with_ltm_window_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            _build("Operating", quarterly=True, days_ref="DAYS_QUARTER")

    def test_explicit_single_quarter_window_keeps_legacy_behavior(self) -> None:
        rows = _build("Operating", quarterly=True, window=1, days_ref="DAYS_QUARTER")
        self.assertIn("*DAYS_QUARTER", _row(rows, "DSO")[1])
        self.assertEqual(_row(rows, "Margen neto")[1],
                         f'=IF(Operating!D{ROW["rev"]}=0,"",'
                         f'Operating!D{ROW["ni"]}/Operating!D{ROW["rev"]})')

    def test_annual_sheet_keeps_begin_end_average(self) -> None:
        rows = _build("Annual", quarterly=False)
        turnover = _row(rows, "Rotacion de activos")
        self.assertIsNone(turnover[0])
        self.assertIn(f"AVERAGE(Annual!C{ROW['ta']},Annual!D{ROW['ta']})", turnover[1])
        self.assertFalse(any(k.endswith("[UDM]") for k in rows))

    def test_ccc_guards_zero_revenue(self) -> None:
        ccc = _row(_build("Annual", quarterly=False), "CCC")[1]
        self.assertIn(f"Annual!D{ROW['rev']}=0", ccc)


if __name__ == "__main__":
    unittest.main()
