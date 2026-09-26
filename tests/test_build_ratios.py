"""Pruebas de ModelStyler.build_ratios: ventana UDM explicita en hojas trimestrales.

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


def _refs(canon: str, cols: str) -> str:
    return ",".join(f"Operating!{c}{ROW[canon]}" for c in cols)


class QuarterlyWindowTests(unittest.TestCase):

    def test_quarterly_sheet_without_window_is_rejected(self) -> None:
        # Adivinar por el header anualizaba dos veces si ref ya trae filas UDM.
        with self.assertRaises(ValueError) as ctx:
            _build("Operating", quarterly=True)
        self.assertIn("window=4", str(ctx.exception))
        self.assertIn("window=1", str(ctx.exception))

    def test_window_one_with_ltm_rows_reproduces_main_formulas(self) -> None:
        # Caso de la cobertura AAPL: ref apunta a filas UDM ya construidas.
        # Formulas identicas a las que genera main hoy (salvo CCC, que ahora
        # tambien se protege contra ventas en cero).
        rows = _build("Operating", quarterly=True, window=1)
        self.assertEqual(_row(rows, "Margen neto")[1],
                         '=IF(Operating!D10=0,"",Operating!D15/Operating!D10)')
        self.assertEqual(_row(rows, "Rotacion de activos")[1],
                         '=IF(AVERAGE(Operating!C18,Operating!D18)=0,"",'
                         'Operating!D10/AVERAGE(Operating!C18,Operating!D18))')
        self.assertEqual(_row(rows, "Deuda / EBITDA")[1],
                         '=IF((Operating!D13+Operating!D29)=0,"",'
                         'Operating!D26/(Operating!D13+Operating!D29))')
        self.assertEqual(_row(rows, "DSO")[1],
                         '=IF(Operating!D10=0,"",AVERAGE(Operating!C21,Operating!D21)'
                         '/Operating!D10*DAYS_YEAR)')
        self.assertFalse(any(k.endswith("[UDM]") for k in rows))

    def test_window_four_sums_four_flows_and_averages_five_closes(self) -> None:
        rows = _build("Operating", quarterly=True, window=4)
        debt_ebitda = _row(rows, "Deuda / EBITDA")
        self.assertEqual(debt_ebitda[:4], [None, None, None, None])
        self.assertIn(f"SUM({_refs('ebit', 'DEFG')})", debt_ebitda[4])
        self.assertIn(f"SUM({_refs('da', 'DEFG')})", debt_ebitda[4])
        # Saldos: 5 cierres t-4..t (apertura y cierre del periodo de 12 meses).
        self.assertIn(f"AVERAGE({_refs('equity', 'CDEFG')})", _row(rows, "ROE DuPont 3")[4])
        self.assertTrue(all(k.endswith("[UDM]") for k in rows))

    def test_ltm_days_ratios_use_days_year(self) -> None:
        rows = _build("Operating", quarterly=True, window=4)
        self.assertIn("*DAYS_YEAR", _row(rows, "DSO")[4])

    def test_days_quarter_with_ltm_window_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            _build("Operating", quarterly=True, window=4, days_ref="DAYS_QUARTER")

    def test_single_quarter_ratios_with_days_quarter_still_allowed(self) -> None:
        rows = _build("Operating", quarterly=True, window=1, days_ref="DAYS_QUARTER")
        self.assertIn("*DAYS_QUARTER", _row(rows, "DSO")[1])


class AnnualTests(unittest.TestCase):

    def test_annual_sheet_defaults_to_begin_end_average(self) -> None:
        rows = _build("Annual", quarterly=False)
        turnover = _row(rows, "Rotacion de activos")
        self.assertIsNone(turnover[0])
        self.assertIn("AVERAGE(Annual!C18,Annual!D18)", turnover[1])
        self.assertFalse(any(k.endswith("[UDM]") for k in rows))

    def test_ccc_guards_zero_revenue(self) -> None:
        ccc = _row(_build("Annual", quarterly=False), "CCC")[1]
        self.assertIn("Annual!D10=0", ccc)


if __name__ == "__main__":
    unittest.main()
