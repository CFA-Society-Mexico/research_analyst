"""Pruebas del audit de formato de tools/xlsx_builder.py (stdlib unittest).

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import xlsx_builder as B  # noqa: E402


class AuditTests(unittest.TestCase):

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def _audit(self, styler: B.ModelStyler, name: str = "t.xlsx") -> dict[str, B.Finding]:
        path = str(Path(self.tmp.name) / name)
        styler.save(path)
        return {f.check.split()[0]: f for f in B.audit_format(path)}

    def _sheet(self) -> tuple[B.ModelStyler, object]:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.label_col_width(ws)
        styler.period_header(ws, 3, 3, B.PeriodHeader(2020, 2027, 2023))
        return styler, ws

    def test_demo_is_green(self) -> None:
        path = str(Path(self.tmp.name) / "carpeta" / "nueva" / "demo.xlsx")
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(path)
            code = B._print_report(B.audit_format(path))
        self.assertEqual(code, 0)

    def test_cash_roll_without_calculated_values_is_pending_not_green(self) -> None:
        styler, ws = self._sheet()
        for r, lab in ((10, "Efectivo al inicio"), (11, "Cambio neto en efectivo"),
                       (12, "Efectivo al cierre")):
            ws.cell(row=r, column=2, value=lab)
        for i in range(8):
            c, p = chr(ord("C") + i), chr(ord("C") + i - 1)
            ws[f"{c}10"] = f"={p}12" if i else 100
            ws[f"{c}11"] = 10
            ws[f"{c}12"] = f"={c}10+{c}11+999"          # roll roto a proposito
        f19 = self._audit(styler)["F19"]
        self.assertFalse(f19.ok)
        self.assertTrue(f19.pending)

    def test_cycle_below_old_400_row_limit_is_detected(self) -> None:
        styler, ws = self._sheet()
        ws["D450"] = "=D451*0.05"
        ws["D451"] = "=100+D450"
        self.assertFalse(self._audit(styler)["F18"].ok)

    def test_sheet_beyond_scan_limit_fails_coverage_check(self) -> None:
        styler, ws = self._sheet()
        ws["B60"] = "fila fuera del limite"
        with mock.patch.object(B, "_MAX_SCAN_ROWS", 50):
            f20 = self._audit(styler)["F20"]
        self.assertFalse(f20.ok)
        self.assertIn("Operating", f20.detail)

    def test_ratio_labels_without_formulas_fail_f13(self) -> None:
        styler, ws = self._sheet()
        for i, lab in enumerate(B.REQUIRED_RATIO_LABELS):
            ws.cell(row=6 + i, column=2, value=lab)
        f13 = self._audit(styler)["F13"]
        self.assertFalse(f13.ok)
        self.assertIn("sin formulas", f13.detail)


class BuilderTests(unittest.TestCase):

    def test_define_constant_supports_two_letter_columns(self) -> None:
        styler = B.ModelStyler()
        styler.new_sheet("Operating")
        styler.define_constant("DAYS_YEAR", "Operating", "AA10")
        self.assertEqual(styler.wb.defined_names["DAYS_YEAR"].attr_text,
                         "'Operating'!$AA$10")

    def test_report_exit_codes(self) -> None:
        ok = B.Finding("F1", True, "")
        pending = B.Finding("F19", False, "", pending=True)
        fail = B.Finding("F2", False, "")
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(B._print_report([ok]), 0)
            self.assertEqual(B._print_report([ok, pending]), 3)
            self.assertEqual(B._print_report([ok, pending, fail]), 1)


if __name__ == "__main__":
    unittest.main()
