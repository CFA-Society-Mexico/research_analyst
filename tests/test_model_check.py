"""Pruebas de tools/model_check.py: checks S y C10 en los dos sentidos.

Cada check tiene un libro que debe PASARLO (el demo del builder) y uno
deliberadamente roto que debe FALLARLO (CONTRIBUTING 4.2). Stdlib unittest,
sin red y sin Excel.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import csv
import io
import sys
import tempfile
import unittest
from copy import copy
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import model_check as M  # noqa: E402
import xlsx_builder as B  # noqa: E402

VERSIONED = "demo_2026-01-01_v1.xlsx"
# Celda de forecast de "Ventas" en el demo (fila 12, primer trimestre E = col K).
FORECAST_CELL = "L12"
CAPTURE_COLUMNS = ("line_item", "period", "value", "source_doc", "source_ref", "tag")


def _unlinked_row(value: float) -> Callable[[Workbook], None]:
    """Mutacion: una fila sin vinculo con una cifra historica tecleada (1Q2023).
    Copia la fuente de una celda del demo para no disparar F2."""
    def mutate(wb: Workbook) -> None:
        ws = wb["Operating"]
        ws["B200"] = "Otros activos"
        ws["B200"].font = copy(ws["B12"].font)
        ws["C200"] = value
        ws["C200"].font = copy(ws["C12"].font)
    return mutate


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

    def _beside_demo(self, mutate: Callable[[Workbook], None] | None = None) -> str:
        """Guarda el demo (mutado) JUNTO al original: comparte su inputs/."""
        wb = load_workbook(self.demo)
        if mutate is not None:
            mutate(wb)
        out = self.dir / "caso_2026-01-02_v1.xlsx"
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


class CaptureLayerTests(unittest.TestCase):
    """Lectura de la capa de captura: vinculos del libro, CSV y umbral."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)

    def _csv(self, text: str) -> Path:
        path = self.dir / "canonical_quarterly.csv"
        path.write_text(text, encoding="utf-8")
        return path

    def test_links_are_read_in_builder_and_excel_reference_forms(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.link_capture(ws, 91, "bs_ap", scale=1_000_000)
        # Excel reescribe la referencia sin comillas al guardar.
        ws.defined_names.add(DefinedName("CAPTURE.bs_ta", attr_text="Operating!$27:$27",
                                         comment="scale=1000000"))
        ws.defined_names.add(DefinedName("OTRO_NOMBRE", attr_text="Operating!$5:$5"))
        self.assertEqual(M.capture_links(ws),
                         {91: M.CaptureLink("bs_ap", 1_000_000.0),
                          27: M.CaptureLink("bs_ta", 1_000_000.0)})

    def test_link_without_scale_is_a_capture_error(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        ws.defined_names.add(DefinedName("CAPTURE.bs_ap", attr_text="Operating!$91:$91"))
        with self.assertRaises(M.CaptureError) as caught:
            M.capture_links(ws)
        self.assertIn("CAPTURE.bs_ap", str(caught.exception))

    def test_capture_values_are_keyed_by_line_and_period(self) -> None:
        path = self._csv("line_item,period,value,source_doc,source_ref,tag\n"
                         "bs_ap,3Q2016,26318000000.0,10-Q,\"XBRL, filed 2016-07-27\",observado\n"
                         "is_ni,4Q2016,9014000000,10-K,FY menos 3Q,derivado\n"
                         "is_ns_total,1Q2027,95000000000,call,p.3,guidance\n")
        self.assertEqual(M.load_capture(path),
                         {("bs_ap", "3Q2016"): 26318000000.0,
                          ("is_ni", "4Q2016"): 9014000000.0})   # guidance no respalda

    def test_capture_without_a_required_column_is_a_capture_error(self) -> None:
        path = self._csv("line_item,period,value\nbs_ap,3Q2016,1\n")
        with self.assertRaises(M.CaptureError) as caught:
            M.load_capture(path)
        self.assertIn("faltan columnas tag", str(caught.exception))

    def test_non_numeric_capture_value_names_file_and_line(self) -> None:
        for bad in ("abc", "", "nan"):
            with self.subTest(value=bad):
                path = self._csv("line_item,period,value,tag\n"
                                 "bs_ap,3Q2016,1,observado\n"
                                 f"bs_ar,3Q2016,{bad},observado\n")
                with self.assertRaises(M.CaptureError) as caught:
                    M.load_capture(path)
                self.assertIn("canonical_quarterly.csv linea 3", str(caught.exception))

    def test_threshold_comes_from_the_profile_or_the_default(self) -> None:
        profile = self.dir / "issuer-profile.yaml"
        profile.write_text("materiality:\n"
                           "  critical_pct_target_price: 10\n"
                           "  capture_unlinked_pct: 2.5   # C9\n", encoding="utf-8")
        self.assertEqual(M.materiality_pct_from_profile(profile), (2.5, "perfil"))
        profile.write_text("materiality:\n  critical_pct_target_price: 10\n", encoding="utf-8")
        self.assertEqual(M.materiality_pct_from_profile(profile), (0.5, "default"))
        self.assertEqual(M.materiality_pct_from_profile(self.dir / "no-existe.yaml"),
                         (0.5, "default"))
        self.assertEqual(M.materiality_pct_from_profile(None), (0.5, "default"))

    def test_period_axis_reports_the_header_row(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Model")
        styler.period_header(ws, 3, 3, B.PeriodHeader(2020, 2025, 2022))
        self.assertEqual({period.row for period in M._period_axis(ws)}, {3})


class CaptureCheckTests(_DemoCase):
    """C9 en los dos sentidos: el demo lo pasa, cada libro roto lo falla."""

    def _c9(self, mutate: Callable[[Workbook], None] | None = None,
            **kwargs: object) -> B.Finding:
        (finding,) = M.audit_capture(self._beside_demo(mutate), **kwargs)  # type: ignore[arg-type]
        return finding

    def _rewrite_capture(self, transform: Callable[[list[dict[str, str]]],
                                                   list[dict[str, str]]]) -> None:
        path = self.dir / "inputs" / "canonical_quarterly.csv"
        with path.open(encoding="utf-8", newline="") as fh:
            rows = list(csv.DictReader(fh))
        with path.open("w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=CAPTURE_COLUMNS)
            writer.writeheader()
            writer.writerows(transform(rows))

    @staticmethod
    def _is(row: dict[str, str], line_item: str, period: str | None = None) -> bool:
        return row["line_item"] == line_item and period in (None, row["period"])

    def test_demo_matches_its_capture_layer(self) -> None:
        finding = self._c9()
        self.assertTrue(finding.ok, finding.detail)
        self.assertFalse(finding.warning)
        self.assertIn("128 cifras en 16 filas coinciden", finding.detail)
        self.assertIn("umbral 0.5%", finding.detail)

    def test_linked_figure_that_differs_from_the_capture_always_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Operating"]["D12"] = 999
        # Ni un umbral del 100% perdona una divergencia.
        finding = self._c9(mutate, materiality_pct=100.0)
        self.assertTrue(finding.blocking)
        self.assertIn("DIVERGENCIAS", finding.detail)
        self.assertIn("Operating!D12 (is_ns_total 2Q2023)", finding.detail)

    def test_material_figure_without_a_link_fails(self) -> None:
        finding = self._c9(_unlinked_row(50))
        self.assertTrue(finding.blocking)
        self.assertIn("SIN RESPALDO sobre el umbral", finding.detail)
        self.assertIn("Operating!fila 200 (Otros activos): 1 cifras sin vinculo", finding.detail)

    def test_immaterial_figure_without_a_link_is_a_warning(self) -> None:
        finding = self._c9(_unlinked_row(0.2))
        self.assertFalse(finding.ok)
        self.assertTrue(finding.warning)
        self.assertFalse(finding.blocking)
        self.assertIn("bajo el umbral", finding.detail)
        self.assertIn("Operating!fila 200", finding.detail)

    def test_threshold_decides_between_failure_and_warning(self) -> None:
        self.assertTrue(self._c9(_unlinked_row(50)).blocking)                       # 50 > 0.5
        self.assertTrue(self._c9(_unlinked_row(50), materiality_pct=60.0).warning)  # 50 <= 60

    def test_linked_row_missing_a_period_in_the_capture_fails(self) -> None:
        self._rewrite_capture(lambda rows: [r for r in rows if not self._is(r, "bs_ar", "3Q2023")])
        finding = self._c9()
        self.assertTrue(finding.blocking)
        self.assertIn("Operating!fila 24", finding.detail)
        self.assertIn("bs_ar sin periodo en la captura (desde 3Q2023)", finding.detail)

    def test_guidance_row_does_not_back_a_historical_figure(self) -> None:
        def retag(tag: str) -> Callable[[list[dict[str, str]]], list[dict[str, str]]]:
            return lambda rows: [dict(r, tag=tag) if self._is(r, "bs_ar", "3Q2023") else r
                                 for r in rows]
        self._rewrite_capture(retag("guidance"))
        self.assertTrue(self._c9().blocking)
        self._rewrite_capture(retag("derivado"))
        self.assertTrue(self._c9().ok)

    def test_scale_converts_model_units_to_source_units(self) -> None:
        def rescale(wb: Workbook) -> None:
            wb["Operating"].defined_names["CAPTURE.bs_cash"].comment = "scale=1000.0"
        # Solo cambia la escala del vinculo: el libro ya no coincide con el CSV.
        self.assertIn("DIVERGENCIAS", self._c9(rescale).detail)
        # El CSV pasa a la unidad de la fuente (x1000): vuelve a coincidir.
        self._rewrite_capture(lambda rows: [
            dict(r, value=str(float(r["value"]) * 1000)) if self._is(r, "bs_cash") else r
            for r in rows])
        self.assertTrue(self._c9(rescale).ok)

    def test_unbacked_figure_without_materiality_reference_fails(self) -> None:
        self._rewrite_capture(lambda rows: [r for r in rows if not self._is(r, "bs_ta")])
        finding = self._c9(materiality_pct=100.0)
        self.assertTrue(finding.blocking)
        self.assertIn("Operating!fila 27", finding.detail)
        self.assertIn("sin referencia de materialidad", finding.detail)
        self.assertIn("la captura no trae is_ns_total o bs_ta de 1Q2023", finding.detail)

    def test_sheet_without_reference_links_says_what_to_link(self) -> None:
        def drop_reference_links(wb: Workbook) -> None:
            names = wb["Operating"].defined_names
            for line in B.CAPTURE_REFERENCE_LINES:
                del names[B.CAPTURE_NAME_PREFIX + line]
        finding = self._c9(drop_reference_links, materiality_pct=100.0)
        self.assertTrue(finding.blocking)
        self.assertIn("la hoja no vincula is_ns_total ni bs_ta", finding.detail)

    def test_assumptions_section_is_out_of_scope(self) -> None:
        def rename_section(wb: Workbook) -> None:
            wb["Operating"]["B5"] = "Supuestos"
        # Sin la banda "Assumptions", sus constantes tecleadas (365 dias) cuentan.
        finding = self._c9(rename_section)
        self.assertTrue(finding.blocking)
        self.assertIn("Operating!fila 7", finding.detail)

    def test_missing_capture_folder_fails_unless_the_model_is_external(self) -> None:
        elsewhere = self._mutated(lambda wb: None)       # roto/ no tiene inputs/
        (strict,) = M.audit_capture(elsewhere)
        self.assertTrue(strict.blocking)
        self.assertIn("sin capa de captura", strict.detail)
        (external,) = M.audit_capture(elsewhere, external=True)
        self.assertTrue(external.ok)
        self.assertIn("n/a", external.detail)
        (pointed,) = M.audit_capture(elsewhere, inputs_dir=str(self.dir / "inputs"))
        self.assertTrue(pointed.ok, pointed.detail)

    def test_unreadable_capture_is_reported_as_a_failure_not_a_traceback(self) -> None:
        capture = self.dir / "inputs" / "canonical_quarterly.csv"
        capture.write_text("line_item,period,value\nbs_ap,1Q2023,20\n", encoding="utf-8")
        finding = self._c9()
        self.assertTrue(finding.blocking)
        self.assertIn("faltan columnas tag", finding.detail)

    def test_detail_is_ascii(self) -> None:
        for finding in (self._c9(), self._c9(_unlinked_row(50)), self._c9(_unlinked_row(0.2))):
            self.assertTrue(finding.detail.isascii(), finding.detail)


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

    def test_external_flag_turns_missing_stamp_into_a_warning(self) -> None:
        def drop_stamp(wb: Workbook) -> None:
            wb.properties.keywords = ""
        unstamped = self._mutated(drop_stamp)
        code, out = self._run(unstamped)
        self.assertEqual(code, 1)
        self.assertIn("[x] F10", out)
        code, out = self._run(unstamped, "--external")
        # 3 y no 1: el aviso no bloquea; queda el pendiente de recalculo (S10).
        self.assertEqual(code, 3)
        self.assertIn("[aviso] F10", out)
        self.assertIn("1 avisos", out)

    def test_c9_is_part_of_the_single_report(self) -> None:
        _, out = self._run(str(self.demo))
        self.assertIn("[ok] C9 historico = capa de captura", out)

    def test_inputs_flag_points_at_a_capture_layer_elsewhere(self) -> None:
        elsewhere = self._mutated(lambda wb: None)       # roto/ no tiene inputs/
        code, out = self._run(elsewhere)
        self.assertEqual(code, 1)
        self.assertIn("[x] C9", out)
        _, out = self._run(elsewhere, "--inputs", str(self.dir / "inputs"))
        self.assertIn("[ok] C9", out)

    def test_profile_threshold_turns_a_material_row_into_a_warning(self) -> None:
        book = self._beside_demo(_unlinked_row(50))
        code, out = self._run(book)                       # sin perfil: default 0.5%
        self.assertEqual(code, 1)
        self.assertIn("[x] C9", out)
        profile = self.dir / "issuer-profile.yaml"
        profile.write_text("materiality:\n  capture_unlinked_pct: 60\n", encoding="utf-8")
        code, out = self._run(book, "--profile", str(profile))
        # 3 y no 1: el aviso no bloquea; queda el pendiente de recalculo (S10).
        self.assertEqual(code, 3)
        self.assertIn("[aviso] C9", out)
        self.assertIn("umbral 60%", out)
        self.assertIn("perfil", out)

    def test_capture_layer_is_resolved_from_the_original_model_path(self) -> None:
        # --recalc audita una COPIA en un directorio temporal: C9 debe seguir
        # leyendo la capa de captura que vive junto al modelo ORIGINAL.
        copied_elsewhere = self._mutated(lambda wb: None)    # roto/ no tiene inputs/
        options = M.capture_options(self.demo, None, None)
        self.assertEqual(Path(options.inputs_dir), (self.dir / "inputs").resolve())
        groups = dict(M.audit_groups(copied_elsewhere, None, None, False, capture=options))
        c9 = groups["Contables por escaneo (C)"][0]
        self.assertTrue(c9.check.startswith("C9 "))
        self.assertTrue(c9.ok, c9.detail)

    def test_missing_file_is_reported_without_traceback(self) -> None:
        code, out = self._run(str(self.dir / "no-existe.xlsx"))
        self.assertEqual(code, 2)
        self.assertIn("[x] no existe el modelo", out)


if __name__ == "__main__":
    unittest.main()
