"""Pruebas de tools/model_check.py: checks S y C10 en los dos sentidos.

Cada check tiene un libro que debe PASARLO (el demo del builder) y uno
deliberadamente roto que debe FALLARLO (CONTRIBUTING 4.2). Stdlib unittest,
sin red y sin Excel.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook.workbook import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import model_check as M  # noqa: E402
import xlsx_builder as B  # noqa: E402

VERSIONED = "demo_2026-01-01_v1.xlsx"
# Celda de forecast de "Ventas" en el demo (fila 12, primer trimestre E = col K).
FORECAST_CELL = "L12"


class _DemoCase(unittest.TestCase):
    """Base: escribe el demo una vez por prueba y permite mutarlo."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        self.demo = self.dir / VERSIONED
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(str(self.demo))

    def _mutated(self, mutate: Callable[[Workbook], None], name: str = VERSIONED) -> str:
        wb = load_workbook(self.demo)
        mutate(wb)
        out = self.dir / "roto" / name
        out.parent.mkdir(exist_ok=True)
        wb.save(out)
        return str(out)

    def _structure(self, path: str, **kwargs: object) -> dict[str, B.Finding]:
        return {f.check.split()[0]: f for f in M.audit_structure(path, **kwargs)}  # type: ignore[arg-type]

    def assertFails(self, finding: B.Finding) -> None:
        self.assertFalse(finding.ok, finding.detail)
        self.assertFalse(finding.pending, finding.detail)


class DemoIsCleanTests(_DemoCase):

    def test_demo_passes_every_structural_check_but_the_recalc_one(self) -> None:
        found = self._structure(str(self.demo))
        self.assertEqual(sorted(found), sorted(f"S{i}" for i in range(1, 11)))
        for check_id, finding in found.items():
            if check_id == "S10":
                continue
            with self.subTest(check=check_id):
                self.assertTrue(finding.ok and not finding.pending, finding.detail)

    def test_book_saved_without_cached_values_is_pending_not_green(self) -> None:
        s10 = self._structure(str(self.demo))["S10"]
        self.assertTrue(s10.pending)
        self.assertFalse(s10.ok)


class StructuralFailureTests(_DemoCase):

    def test_s1_missing_contract_tab_fails(self) -> None:
        path = self._mutated(lambda wb: wb.remove(wb["Macro"]))
        self.assertFails(self._structure(path)["S1"])
        self.assertIn("Macro", self._structure(path)["S1"].detail)

    def test_s1_driver_map_schedule_without_block_fails(self) -> None:
        found = self._structure(str(self.demo), driver_schedules=["PPE", "iPhone"])
        self.assertFails(found["S1"])
        self.assertIn("iPhone", found["S1"].detail)

    def test_s1_driver_map_schedules_present_pass(self) -> None:
        found = self._structure(str(self.demo), driver_schedules=["PPE", "WC"])
        self.assertTrue(found["S1"].ok)

    def test_s2_link_to_another_workbook_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Annual"]["C12"] = "=[otro.xlsx]Hoja1!A1"
        self.assertFails(self._structure(self._mutated(mutate))["S2"])

    def test_s2_bracket_inside_text_is_not_a_link(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Summary"]["C8"] = '=IF(Operating!C12>0,"[ok]","[x]")'
        self.assertTrue(self._structure(self._mutated(mutate))["S2"].ok)

    def test_s3_volatile_function_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Summary"]["C8"] = "=TODAY()"
        finding = self._structure(self._mutated(mutate))["S3"]
        self.assertFails(finding)
        self.assertIn("TODAY", finding.detail)

    def test_s4_hardcoded_number_outside_assumptions_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Annual"]["C12"] = "=C6*1.05"
        finding = self._structure(self._mutated(mutate))["S4"]
        self.assertFails(finding)
        self.assertIn("Annual!C12", finding.detail)

    def test_s5_different_formula_inside_a_segment_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Operating"][FORECAST_CELL] = "=K12+K13"
        finding = self._structure(self._mutated(mutate))["S5"]
        self.assertFails(finding)
        self.assertIn("fila 12", finding.detail)

    def test_s6_input_outside_assumptions_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            cell = wb["Operating"]["M13"]
            cell.fill = PatternFill("solid", fgColor=B.Color.INPUT_FILL.value)
        finding = self._structure(self._mutated(mutate))["S6"]
        self.assertFails(finding)
        self.assertIn("Operating!M13", finding.detail)

    def test_s6_without_assumptions_section_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Operating"]["B5"] = "Supuestos"
        self.assertFails(self._structure(self._mutated(mutate))["S6"])

    def test_s7_manual_calculation_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb.calculation.calcMode = "manual"
        self.assertFails(self._structure(self._mutated(mutate))["S7"])

    def test_s8_unversioned_file_name_fails(self) -> None:
        path = self._mutated(lambda wb: None, name="demo_final_v2_OK.xlsx")
        self.assertFails(self._structure(path)["S8"])

    def test_s9_loose_schedule_or_statement_sheet_fails(self) -> None:
        for sheet in ("Sch_WC", "IS", "Val_DCF"):
            with self.subTest(sheet=sheet):
                path = self._mutated(lambda wb, s=sheet: wb.create_sheet(s) and None)
                finding = self._structure(path)["S9"]
                self.assertFails(finding)
                self.assertIn(sheet, finding.detail)


class HardcodedLiteralTests(unittest.TestCase):

    def test_structural_and_named_values_are_not_hardcodes(self) -> None:
        clean = (
            "=C12*(1+C9)",                       # 1 estructural
            "=IF(C12=0,0,C13/C12-1)",            # 0 y 1
            "=C24/C12*DAYS_YEAR",                # constante con nombre
            "=LOG10(C12)",                       # digitos en nombre de funcion
            "=SUM(Operating!C12:F12)",           # digitos en referencias
            "='Val 2026'!C5+RATE_2Y",            # digitos en hoja y en nombre
            '=IF(C12>0,"4Q2026","")',            # digitos dentro de texto
        )
        for formula in clean:
            with self.subTest(formula=formula):
                self.assertEqual(M.hardcoded_literals(formula), [])

    def test_embedded_assumptions_are_reported(self) -> None:
        dirty = {"=C12*1.05": ["1.05"], "=C24/C12*365": ["365"],
                 "=AVERAGE(C12:F12)/4": ["4"], "=C12*(1-0.21)": ["0.21"]}
        for formula, expected in dirty.items():
            with self.subTest(formula=formula):
                self.assertEqual(M.hardcoded_literals(formula), expected)


class RelativeFormTests(unittest.TestCase):

    QUARTERS = {"Operating": {3 + i: M.PeriodColumn(3 + i, True, 2023 + i // 4, i % 4 + 1)
                              for i in range(8)}}

    def _form(self, formula: str, column: int, year: int | None = None) -> str:
        return M.relative_form(formula, column, year, self.QUARTERS, frozenset({"TAX21"}))

    def test_same_structure_in_neighbour_columns_is_one_form(self) -> None:
        self.assertEqual(self._form("=D12*(1+D9)-C13", 4), self._form("=E12*(1+E9)-D13", 5))

    def test_shifted_reference_is_a_different_form(self) -> None:
        self.assertNotEqual(self._form("=D12-C13", 4), self._form("=E12-E13", 5))

    def test_absolute_column_is_a_constant_not_a_period(self) -> None:
        self.assertEqual(self._form("=D12*$C$8", 4), self._form("=E12*$C$8", 5))

    def test_annual_aggregate_of_four_quarters_is_one_form(self) -> None:
        fy2023 = self._form("=SUM(Operating!C12:F12)", 3, year=2023)
        fy2024 = self._form("=SUM(Operating!G12:J12)", 4, year=2024)
        self.assertEqual(fy2023, fy2024)

    def test_annual_aggregate_of_the_wrong_quarters_is_caught(self) -> None:
        fy2023 = self._form("=SUM(Operating!C12:F12)", 3, year=2023)
        shifted = self._form("=SUM(Operating!F12:I12)", 4, year=2024)
        self.assertNotEqual(fy2023, shifted)

    def test_defined_name_that_looks_like_a_cell_is_left_alone(self) -> None:
        self.assertEqual(self._form("=D12*TAX21", 4), self._form("=E12*TAX21", 5))


class UncachedVerdictTests(unittest.TestCase):

    def test_all_formulas_calculated_is_green(self) -> None:
        finding = M.uncached_verdict(120, [])
        self.assertTrue(finding.ok and not finding.pending)

    def test_some_formulas_without_value_after_recalc_fails(self) -> None:
        finding = M.uncached_verdict(120, ["Operating!BH79", "Operating!BH119"])
        self.assertFalse(finding.ok)
        self.assertFalse(finding.pending)
        self.assertIn("Operating!BH79", finding.detail)

    def test_no_formula_calculated_is_pending(self) -> None:
        self.assertTrue(M.uncached_verdict(2, ["A!B1", "A!B2"]).pending)


class CashRollFormulaTests(unittest.TestCase):

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)

    def _book(self, opening: Callable[[str, str], str], balance: Callable[[str], str]) -> str:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Model")
        styler.period_header(ws, 3, 3, B.PeriodHeader(2020, 2025, 2022))
        for row, text in ((6, "Efectivo y equivalentes"), (10, "Cambio neto en efectivo"),
                          (11, "Efectivo al inicio"), (12, "Efectivo al cierre")):
            ws.cell(row=row, column=2, value=text)
        letters = "CDEFGH"
        for i, col in enumerate(letters):
            prev = letters[i - 1] if i else col
            ws[f"{col}6"] = balance(col)
            ws[f"{col}10"] = 5
            ws[f"{col}11"] = 40 if i == 0 else opening(col, prev)
            ws[f"{col}12"] = f"={col}11+{col}10"
        path = str(self.dir / "roll.xlsx")
        styler.save(path)
        return path

    def test_well_chained_roll_passes(self) -> None:
        path = self._book(lambda col, prev: f"={prev}12", lambda col: f"={col}12")
        (finding,) = M.audit_cash_roll_formulas(path)
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("11 celdas", finding.detail)

    def test_opening_cash_reading_same_column_close_fails(self) -> None:
        path = self._book(lambda col, prev: f"={col}12-{col}10", lambda col: f"={col}12")
        (finding,) = M.audit_cash_roll_formulas(path)
        self.assertFalse(finding.ok)
        self.assertIn("columna previa", finding.detail)

    def test_balance_cash_not_tied_to_cash_flow_close_fails(self) -> None:
        path = self._book(lambda col, prev: f"={prev}12", lambda col: f"={col}11")
        (finding,) = M.audit_cash_roll_formulas(path)
        self.assertFalse(finding.ok)
        self.assertIn("caja del balance", finding.detail)


class DriverMapTests(unittest.TestCase):

    def test_schedule_blocks_are_read_from_the_bloque_column(self) -> None:
        text = ("| Segmento | Driver | Bloque (`Sch: <nombre>`) | Ancla |\n"
                "|---|---|---|---|\n"
                "| iPhone | unidades | `Sch: iPhone` | guidance |\n"
                "| Servicios | suscriptores | `Sch: Services` | reporte |\n"
                "| Capital de trabajo | dias | `Sch: WC` | historico |\n")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "driver-map.md"
            path.write_text(text, encoding="utf-8")
            self.assertEqual(M.schedules_from_driver_map(str(path)), ["Services", "WC", "iPhone"])


class CliTests(_DemoCase):

    def _run(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = M.main(["model_check.py", *args])
        return code, out.getvalue()

    def test_uncalculated_demo_exits_pending(self) -> None:
        code, out = self._run(str(self.demo))
        self.assertEqual(code, 3)
        self.assertIn("Veredicto: PENDIENTE", out)
        self.assertTrue(out.isascii())

    def test_broken_book_exits_failure(self) -> None:
        code, out = self._run(self._mutated(lambda wb: wb.create_sheet("Sch_WC") and None))
        self.assertEqual(code, 1)
        self.assertIn("[x] S9", out)

    def test_missing_file_is_reported_without_traceback(self) -> None:
        code, out = self._run(str(self.dir / "no-existe.xlsx"))
        self.assertEqual(code, 2)
        self.assertIn("[x] no existe el modelo", out)


if __name__ == "__main__":
    unittest.main()
