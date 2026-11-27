"""Pruebas del audit de formato de tools/xlsx_builder.py (stdlib unittest).

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import csv
import io
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Callable
from unittest import mock

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.workbook import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import xlsx_builder as B  # noqa: E402

Mutation = Callable[[Workbook], None]

# Colores y formatos FUERA de las whitelists del builder.
ALIEN_FONT_COLOR = "FF7030A0"
ALIEN_NUMFMT = "0.000"
LEGACY_FILL = B.Color.ORANGE.value  # paleta CFI que salio de la whitelist


def _set(sheet: str, coord: str, value: object) -> Mutation:
    def mutate(wb: Workbook) -> None:
        wb[sheet][coord] = value
    return mutate


def _gridlines_on(wb: Workbook) -> None:
    wb["Annual"].sheet_view.showGridLines = True


def _alien_font(wb: Workbook) -> None:
    wb["Operating"]["B12"].font = Font(name="Calibri", size=11)


def _alien_font_color(wb: Workbook) -> None:
    wb["Operating"]["B12"].font = Font(name=B.FONT_NAME, size=11, color=ALIEN_FONT_COLOR)


def _legacy_fill(wb: Workbook) -> None:
    wb["Operating"]["C12"].fill = PatternFill("solid", fgColor=LEGACY_FILL)


def _alien_number_format(wb: Workbook) -> None:
    wb["Operating"]["C12"].number_format = ALIEN_NUMFMT


def _no_freeze(wb: Workbook) -> None:
    wb["Operating"].freeze_panes = None


def _ungroup_balance_section(wb: Workbook) -> None:
    ws = wb["Operating"]
    for row in range(23, 33):  # filas de la seccion Balance del demo
        ws.row_dimensions[row].outlineLevel = 0


def _junk_sheet(wb: Workbook) -> None:
    wb.create_sheet("Hoja1").sheet_view.showGridLines = False


def _no_period_formats(wb: Workbook) -> None:
    for cell in wb["Annual"][3]:
        cell.number_format = "General"


def _no_stamp(wb: Workbook) -> None:
    wb.properties.keywords = ""


def _split_series_row(wb: Workbook) -> None:
    ws = wb["Annual"]
    ws["B12"] = "Ventas forecast"   # solo las columnas E (2025, 2026) pobladas
    ws["E12"] = 1
    ws["F12"] = 2


def _rename_roic(wb: Workbook) -> None:
    ws = wb["Operating"]
    for row in range(1, ws.max_row + 1):
        cell = ws.cell(row=row, column=2)
        # La fila de la razon ("ROIC [UDM]"), no la banda "ROIC y economic profit".
        if isinstance(cell.value, str) and cell.value.startswith("ROIC") and not cell.font.bold:
            cell.value = "Retorno sobre el capital"
            return
    raise AssertionError("el demo ya no trae la fila ROIC")


def _input_in_annual(wb: Workbook) -> None:
    wb["Annual"]["C6"].fill = PatternFill("solid", fgColor=B.Color.INPUT_FILL.value)


def _cycle(wb: Workbook) -> None:
    ws = wb["Operating"]
    ws["T50"] = "=T51"
    ws["T51"] = "=T50"


# Un libro roto por check: (mutacion sobre el demo, texto que el detalle debe nombrar).
# CONTRIBUTING 4.2: "un libro que debe fallarlo y el demo que debe pasarlo".
BREAKING_MUTATIONS: dict[str, tuple[Mutation, str]] = {
    "F1": (_gridlines_on, "Annual"),
    "F2": (_alien_font, "Calibri"),
    "F3": (_alien_font_color, ALIEN_FONT_COLOR),
    "F4": (_legacy_fill, LEGACY_FILL),
    "F5": (_alien_number_format, ALIEN_NUMFMT),
    "F6": (_no_freeze, "Operating"),
    "F7": (_ungroup_balance_section, "Balance"),
    "F8": (_junk_sheet, "Hoja1"),
    "F9": (_no_period_formats, "sin formato"),
    "F10": (_no_stamp, "sin sello"),
    "F11": (_set("Operating", "M9", None), "Operating!fila 9"),      # hueco en fila de input
    "F12": (_split_series_row, "Annual!fila 12 (solo forecast)"),
    "F13": (_rename_roic, "faltan: ROIC"),
    "F14": (_input_in_annual, "celdas de INPUT"),
    "F15": (_set("Operating", "L13", "n/d"), "TEXTO"),               # texto en una serie
    "F16": (_set("Operating", "B10", "fila pegada al header"), "Operating!fila 11"),
    "F17": (_set("Operating", "N15", None), "Operating!fila 15"),    # forecast incompleto
    "F18": (_cycle, "CICLOS"),
    # ABS admite un argumento: Excel repara el libro y borra la formula (issue #33).
    "F21": (_set("Operating", "T60", "=MAX(ABS(C12,D12,E12))"), "Operating!T60: ABS"),
}
# Checks cuya falla no se puede provocar mutando el demo; tienen su propia prueba
# en AuditTests (F19 necesita valores calculados; F20, bajar el limite de escaneo).
TESTED_OUTSIDE_THE_TABLE = {"F19", "F20"}


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

    def test_demo_writes_its_capture_layer_next_to_the_book(self) -> None:
        path = Path(self.tmp.name) / "demo.xlsx"
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(str(path))
        capture = path.parent / "inputs" / "canonical_quarterly.csv"
        with capture.open(encoding="utf-8", newline="") as fh:
            rows = list(csv.DictReader(fh))
        names = set(load_workbook(path)["Operating"].defined_names)
        self.assertEqual(len(rows), 16 * 8)     # 16 filas observadas x 8 trimestres A
        self.assertEqual({row["tag"] for row in rows}, {"observado"})
        self.assertEqual(rows[0]["period"], "1Q2023")
        lines = {row["line_item"] for row in rows}
        for reserved in B.CAPTURE_REFERENCE_LINES:
            self.assertIn(reserved, lines)
            self.assertIn(B.CAPTURE_NAME_PREFIX + reserved, names)


class BrokenBookTests(unittest.TestCase):
    """Cada check F falla sobre un libro roto a proposito (y solo por eso)."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        self.demo = self.dir / "demo.xlsx"
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(str(self.demo))

    def _audit(self, mutate: Mutation | None = None,
               brand: dict[str, str] | None = None) -> dict[str, B.Finding]:
        wb = load_workbook(self.demo)
        if mutate is not None:
            mutate(wb)
        path = str(self.dir / "caso.xlsx")
        wb.save(path)
        return {f.check.split()[0]: f for f in B.audit_format(path, brand=brand)}

    def test_demo_reloaded_and_saved_untouched_stays_green(self) -> None:
        # Linea base: si esto falla, las mutaciones de abajo no prueban nada.
        broken = {k: f.detail for k, f in self._audit().items() if not f.ok}
        self.assertEqual(broken, {})

    def test_every_format_check_has_a_breaking_mutation(self) -> None:
        emitted = set(self._audit())
        self.assertEqual(set(BREAKING_MUTATIONS) | TESTED_OUTSIDE_THE_TABLE, emitted,
                         "check F sin libro que lo falle (o mutacion de un check que ya no existe)")

    def test_each_mutation_fails_its_check_and_names_the_culprit(self) -> None:
        for check_id, (mutate, culprit) in BREAKING_MUTATIONS.items():
            with self.subTest(check=check_id):
                finding = self._audit(mutate)[check_id]
                self.assertFalse(finding.ok, finding.detail)
                self.assertFalse(finding.pending, finding.detail)
                self.assertIn(culprit, finding.detail)

    def test_f4_brand_slot_extends_the_fill_whitelist(self) -> None:
        self.assertFalse(self._audit(_legacy_fill)["F4"].ok)
        branded = self._audit(_legacy_fill, brand={"brand_accent": LEGACY_FILL})
        self.assertTrue(branded["F4"].ok, branded["F4"].detail)

    def test_f8_schedule_as_its_own_sheet_fails(self) -> None:
        finding = self._audit(lambda wb: wb.create_sheet("Sch_WC") and None)["F8"]
        self.assertFalse(finding.ok)
        self.assertIn("Sch_WC", finding.detail)

    def test_f14_quarter_column_in_annual_fails(self) -> None:
        finding = self._audit(_set("Annual", "G3", "1Q2027E"))["F14"]
        self.assertFalse(finding.ok)
        self.assertIn("columnas #Q", finding.detail)

    def test_f15_missing_formula_in_the_middle_of_a_series_fails(self) -> None:
        finding = self._audit(_set("Operating", "L14", None))["F15"]
        self.assertFalse(finding.ok)
        self.assertIn("Operating!fila 14", finding.detail)

    def test_f17_discontinued_series_is_exempt_from_full_forecast(self) -> None:
        def discontinue(wb: Workbook) -> None:
            ws = wb["Operating"]
            ws["B15"] = "D&A (serie descontinuada por la emisora)"
            for col in range(11, 19):  # tramo E del demo (K..R)
                ws.cell(row=15, column=col).value = None
        self.assertTrue(self._audit(discontinue)["F17"].ok)


class LiveChecksTests(unittest.TestCase):
    """Issue #33: un check de formula viva que nunca se evaluo no es verde."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.path = str(Path(tmp.name) / "vivos.xlsx")

    def _book(self, formula: object) -> B.ModelStyler:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Checks", freeze=None)
        ws["B5"] = "C1 Balance cuadra"
        if formula is not None:
            styler.check_result(ws, 5, 3, formula)
        return styler

    def test_formulas_without_cached_values_are_pending_and_say_so(self) -> None:
        self._book('=IF(ABS(D5)>0.0001,"Error","OK")').save(self.path)
        finding = B.live_checks_finding(self.path)
        self.assertTrue(finding.pending)
        self.assertIn("SIN evaluar", finding.detail)
        self.assertIn("1 de 1", finding.detail)

    def test_checks_tab_without_formulas_is_not_applicable(self) -> None:
        self._book(None).save(self.path)
        finding = B.live_checks_finding(self.path)
        self.assertTrue(finding.ok)
        self.assertFalse(finding.pending)

    def test_cli_audit_is_not_green_while_live_checks_are_unevaluated(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(self.path)
        wb = load_workbook(self.path)
        wb["Checks"]["C7"] = '=IF(ABS(Operating!C27)>0,"OK","Error")'
        wb["Checks"]["C7"].font = Font(name=B.FONT_NAME, size=11)   # F2 sigue verde
        wb.save(self.path)
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = B.main(["xlsx_builder.py", "audit", self.path])
        self.assertEqual(code, 3)
        self.assertIn("[pendiente] Checks vivos", out.getvalue())
        self.assertIn("SIN evaluar", out.getvalue())

    def test_evaluated_error_is_a_failure_and_names_the_check(self) -> None:
        # Valores cacheados simulados: se parchea la lectura data_only.
        self._book('=IF(ABS(D5)>0.0001,"Error","OK")').save(self.path)
        cached = B.Workbook()
        ws = cached.active
        ws.title = "Checks"
        real_load = B.load_workbook

        def fake_load(path: str, data_only: bool = False, **kwargs: object) -> Workbook:
            return cached if data_only else real_load(path, data_only=False, **kwargs)

        for value, expect_ok in (("Error", False), ("#NAME?", False), ("OK", True)):
            with self.subTest(value=value):
                ws["C5"] = value
                with mock.patch.object(B, "load_workbook", side_effect=fake_load):
                    finding = B.live_checks_finding(self.path)
                self.assertEqual(finding.ok, expect_ok, finding.detail)
                self.assertFalse(finding.pending)
                if not expect_ok:
                    self.assertIn("C1 Balance cuadra", finding.detail)

    def _evaluate(self, styler: B.ModelStyler, cached_values: dict[str, object]) -> B.Finding:
        styler.save(self.path)
        cached = B.Workbook()
        ws = cached.active
        ws.title = "Checks"
        for coord, value in cached_values.items():
            ws[coord] = value
        real_load = B.load_workbook

        def fake_load(path: str, data_only: bool = False, **kwargs: object) -> Workbook:
            return cached if data_only else real_load(path, data_only=False, **kwargs)

        with mock.patch.object(B, "load_workbook", side_effect=fake_load):
            return B.live_checks_finding(self.path)

    def test_evaluated_review_is_a_warning_that_names_the_check(self) -> None:
        finding = self._evaluate(self._book('=IF(D5>D6,"REVISAR","OK")'), {"C5": "REVISAR"})
        self.assertTrue(finding.warning, finding.detail)
        self.assertFalse(finding.blocking)
        self.assertFalse(finding.pending)
        self.assertIn("REVISAR", finding.detail)
        self.assertIn("C1 Balance cuadra", finding.detail)

    def test_error_still_fails_when_another_check_says_review(self) -> None:
        styler = self._book('=IF(ABS(D5)>0.0001,"Error","OK")')
        checks = styler.wb["Checks"]
        checks["B6"] = "D5 convergencia de metodos"
        styler.check_result(checks, 6, 3, '=IF(D6>D7,"REVISAR","OK")')
        finding = self._evaluate(styler, {"C5": "Error", "C6": "REVISAR"})
        self.assertTrue(finding.blocking, finding.detail)
        self.assertIn("C1 Balance cuadra", finding.detail)
        self.assertFalse(finding.warning)
        self.assertFalse(finding.pending)

    def test_review_does_not_hide_unevaluated_check(self) -> None:
        styler = self._book('=IF(ABS(D5)>0.0001,"Error","OK")')
        checks = styler.wb["Checks"]
        checks["B6"] = "D5 convergencia de metodos"
        styler.check_result(checks, 6, 3, '=IF(D6>D7,"REVISAR","OK")')
        finding = self._evaluate(styler, {"C6": "REVISAR"})   # C5 sin cachear
        self.assertTrue(finding.pending, finding.detail)
        self.assertFalse(finding.warning)
        self.assertIn("SIN evaluar", finding.detail)

    def test_warning_label_is_ascii_for_the_console(self) -> None:
        styler = self._book('=IF(D5>D6,"REVISAR","OK")')
        styler.wb["Checks"]["B5"] = "D5 convergencia de métodos ✓"
        finding = self._evaluate(styler, {"C5": "REVISAR"})
        self.assertTrue(finding.warning, finding.detail)
        self.assertIn("D5 convergencia de m?todos ?", finding.detail)
        self.assertTrue(finding.detail.isascii(), finding.detail)

    def _cli_audit_with_cached(self, review_cell: str | None) -> tuple[int, str]:
        """Audita el demo con todos los checks vivos cacheados en OK (y uno en
        REVISAR si se pide), parcheando la lectura data_only."""
        real = load_workbook(self.path)
        cached = B.Workbook()
        cached.remove(cached.active)
        for sheet in B.LIVE_CHECK_SHEETS:
            if sheet not in real.sheetnames:
                continue
            out_ws = cached.create_sheet(sheet)
            for row in real[sheet].iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        out_ws[cell.coordinate] = "OK"
        if review_cell:
            sheet, coord = review_cell.split("!")
            cached[sheet][coord] = "REVISAR"
        real_load = B.load_workbook

        def fake_load(path: str, data_only: bool = False, **kwargs: object) -> Workbook:
            if not data_only:
                return real_load(path, data_only=False, **kwargs)
            # Las demas tabs (F19 lee Annual/Operating) conservan su lectura real.
            book = real_load(path, data_only=True, **kwargs)
            for sheet in cached.sheetnames:
                for row in cached[sheet].iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            book[sheet][cell.coordinate] = cell.value
            return book

        out = io.StringIO()
        with mock.patch.object(B, "load_workbook", side_effect=fake_load),                 contextlib.redirect_stdout(out):
            code = B.main(["xlsx_builder.py", "audit", self.path])
        return code, out.getvalue()

    def test_cli_audit_lists_review_as_warning_and_exits_green(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(self.path)
        wb = load_workbook(self.path)
        for coord, formula in (("C7", '=IF(ABS(Operating!C27)>0,"OK","Error")'),
                               ("C8", '=IF(Operating!C27>1,"REVISAR","OK")')):
            wb["Checks"][coord] = formula
            wb["Checks"][coord].font = Font(name=B.FONT_NAME, size=11)   # F2 sigue verde
        wb.save(self.path)
        target = "Checks!C8"
        base_code, base_out = self._cli_audit_with_cached(None)
        code, out = self._cli_audit_with_cached(target)
        self.assertNotIn("[aviso] Checks vivos", base_out, base_out)
        self.assertIn("[aviso] Checks vivos", out)
        self.assertEqual(code, base_code, out)
        self.assertNotIn(code, (1, 3), out)


class ExternalModelTests(unittest.TestCase):
    """F10 en un modelo que no construyo el builder: falla en el pipeline del
    plugin, [aviso] cuando quien audita declara que el modelo es externo."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.unstamped = str(Path(tmp.name) / "externo.xlsx")
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(self.unstamped)
        wb = load_workbook(self.unstamped)
        _no_stamp(wb)
        wb.save(self.unstamped)

    def _f10(self, **kwargs: bool) -> B.Finding:
        return next(f for f in B.audit_format(self.unstamped, **kwargs)
                    if f.check.startswith("F10 "))

    def _cli(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = B.main(["xlsx_builder.py", *args])
        return code, out.getvalue()

    def test_missing_stamp_blocks_by_default(self) -> None:
        f10 = self._f10()
        self.assertFalse(f10.ok)
        self.assertTrue(f10.blocking)

    def test_missing_stamp_on_declared_external_model_is_a_warning(self) -> None:
        f10 = self._f10(external=True)
        self.assertFalse(f10.ok)          # el hecho no cambia: no hay sello
        self.assertTrue(f10.warning)
        self.assertFalse(f10.blocking)

    def test_stamped_model_is_plain_ok_even_if_declared_external(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "demo.xlsx")
            with contextlib.redirect_stdout(io.StringIO()):
                B._demo(path)
            f10 = next(f for f in B.audit_format(path, external=True)
                       if f.check.startswith("F10 "))
        self.assertTrue(f10.ok)
        self.assertFalse(f10.warning)

    def test_cli_external_flag_reports_warning_and_exits_green(self) -> None:
        code, out = self._cli("audit", self.unstamped, "--external")
        self.assertEqual(code, 0)
        self.assertIn("[aviso] F10", out)
        self.assertIn("1 avisos", out)

    def test_cli_without_flag_still_fails_on_missing_stamp(self) -> None:
        code, out = self._cli("audit", self.unstamped)
        self.assertEqual(code, 1)
        self.assertIn("[x] F10", out)


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

    def test_warning_never_changes_the_exit_code(self) -> None:
        ok = B.Finding("F1", True, "")
        warn = B.Finding("F10", False, "", warning=True)
        pending = B.Finding("F19", False, "", pending=True)
        fail = B.Finding("F2", False, "")
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(B._print_report([ok, warn]), 0)
            self.assertEqual(B._print_report([warn, pending]), 3)
            self.assertEqual(B._print_report([warn, fail]), 1)

    def test_link_capture_writes_a_sheet_scoped_name_with_row_and_scale(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.link_capture(ws, 91, "bs_ap", scale=1_000_000)
        defined = ws.defined_names["CAPTURE.bs_ap"]
        self.assertEqual(defined.attr_text, "'Operating'!$91:$91")
        self.assertEqual(defined.comment, "scale=1000000.0")
        # Alcance de hoja: el nombre no existe a nivel libro.
        self.assertNotIn("CAPTURE.bs_ap", styler.wb.defined_names)

    def test_link_capture_survives_save_and_load(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.link_capture(ws, 91, "bs_ap", scale=1_000_000)
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "t.xlsx")
            styler.save(path)
            reloaded = load_workbook(path)["Operating"].defined_names["CAPTURE.bs_ap"]
        self.assertEqual(reloaded.attr_text, "'Operating'!$91:$91")
        self.assertEqual(reloaded.comment, "scale=1000000.0")

    def test_link_capture_allows_the_same_line_once_per_sheet(self) -> None:
        styler = B.ModelStyler()
        operating = styler.new_sheet("Operating")
        annual = styler.new_sheet("Annual")
        styler.link_capture(operating, 91, "bs_ap")
        styler.link_capture(annual, 40, "bs_ap")        # otra hoja: permitido
        with self.assertRaises(ValueError):
            styler.link_capture(operating, 92, "bs_ap")  # misma hoja: rechazado

    def test_link_capture_rejects_bad_line_item_and_scale(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        for line_item in ("bs ap", "1bs", "bs-ap", ""):
            with self.subTest(line_item=line_item), self.assertRaises(ValueError):
                styler.link_capture(ws, 10, line_item)
        for scale in (0, -1, True, float("nan")):
            with self.subTest(scale=scale), self.assertRaises(ValueError):
                styler.link_capture(ws, 10, "bs_ap", scale=scale)


if __name__ == "__main__":
    unittest.main()
