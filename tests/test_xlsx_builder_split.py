"""Red de seguridad de la particion de tools/xlsx_builder.py en modulos.

``xlsx_builder.py`` es una fachada sobre ``xlsx_contract``, ``xlsx_formulas``,
``xlsx_styler``, ``xlsx_scan`` y ``xlsx_audit``. Estas pruebas fijan lo que la
particion no puede cambiar:

1. Paridad del demo: el libro que escribe ``_demo`` es identico, celda por
   celda, a la referencia ``tests/fixtures/demo_reference.xlsx`` (generada con
   el builder de un solo archivo, antes de partirlo): valores, formulas,
   formato numerico, fuente, relleno, bordes, alineacion, comentarios, nombres
   definidos, agrupacion de filas, anchos, paneles y propiedades del libro.
2. API congelada: todo nombre que el modulo exponia sigue en la fachada.
3. Un parche sobre la fachada (``mock.patch.object(B, ...)``) alcanza al
   codigo de los submodulos que lee ese nombre.
4. CLI: mismos textos y exit codes que la referencia.

Si el demo cambia A PROPOSITO, regenera las referencias y revisa el diff:
    python tests/test_xlsx_builder_split.py --regenerate

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import csv
import io
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
REFERENCE_BOOK = FIXTURES / "demo_reference.xlsx"
REFERENCE_CAPTURE = FIXTURES / "demo_reference_capture.csv"
REFERENCE_DEMO_OUT = FIXTURES / "demo_reference_demo.txt"
REFERENCE_AUDIT_OUT = FIXTURES / "demo_reference_audit.txt"
BOOK_TOKEN = "<BOOK>"

sys.path.insert(0, str(TOOLS))

import xlsx_builder as B  # noqa: E402

# Namespace de xlsx_builder.py ANTES de partirlo (98 nombres sin dunder). Los
# scripts de build de coberturas viven fuera del repo: cualquiera de estos
# nombres puede estar en uso, privados incluidos.
API_BEFORE_SPLIT = frozenset("""
Alignment ArrayFormula BRAND_SLOTS BUILDER_STAMP_KEY BUILDER_STAMP_VALUE Border
CAPTURE_NAME_PREFIX CAPTURE_REFERENCE_LINES CASH_LABEL_PREFIXES CalcProperties
Callable CellRole Color Comment Enum FONT_NAME FROZEN_SHEET_PREFIXES
FUNCTION_ARITY Finding Font Iterable JUNK_SHEET_NAMES LIVE_CHECK_SHEETS
ModelStyler NOT_MEANINGFUL NO_GROWTH_BELOW_WACC NO_REAL_SOLUTION NumFmt Optional
Path PatternFill PeriodHeader RATIO_FLOW_CANONS REQUIRED_RATIO_LABELS
REVIEW_FLAG Side TRANSITION_STAMP_KEY Token Tokenizer TokenizerError Workbook
Worksheet XLFN_FUNCTIONS _ALLOWED_FILLS _ALLOWED_FONT_COLORS _ALLOWED_NUMFMTS
_ARRAY_NATIVE _CASH_CLOSE_WORDS _CASH_OPEN_WORDS _DOUBLE _EXCEL_ERRORS _FN_NAME
_LOCAL_REF _MANY _MAX_SCAN_COLS _MAX_SCAN_ROWS _QUARTER_HDR _ROLE_FONT_COLOR
_SHEET_REF _THIN _UNSET _XL_PREFIXES _cash_roll_violations _demo _find_cycles
_formula_deps _function_name _gap_violations _has_quarter_header _is_header_row
_is_multicell _period_columns _print_report _re_mod _scan_fonts_fills_formats
_series_continuity_violations _split_series_violations _write_demo_capture
audit_format auto cash_roll_bounds center_check_formula column_index_from_string
comps_market_cap_formula csv dataclass date dcf_per_share_formula
get_column_letter lint_formula live_checks_finding load_brand load_workbook main
multiple_range_flag_formula sys terminal_multiple_alarm_formula
transition_year_of
""".split())


# --- Foto del libro -----------------------------------------------------------

def _color(color: object) -> object:
    if color is None:
        return None
    return (color.type, color.rgb if color.type == "rgb" else color.value, color.tint)


def _side(side: object) -> object:
    return None if side is None else (side.style, _color(side.color))


def _value(value: object) -> object:
    if isinstance(value, ArrayFormula):
        return ("ArrayFormula", value.ref, value.text)
    return (type(value).__name__, value)


def _cell(cell: object) -> tuple:
    font, fill, border, align = cell.font, cell.fill, cell.border, cell.alignment
    return (
        _value(cell.value),
        cell.number_format,
        (font.name, font.sz, font.b, font.i, font.u, font.strike, _color(font.color)),
        (fill.fill_type, _color(fill.fgColor), _color(fill.bgColor)),
        tuple(_side(getattr(border, s)) for s in ("left", "right", "top", "bottom")),
        (align.horizontal, align.vertical, align.indent, align.wrap_text),
        cell.protection.locked,
        None if cell.comment is None else (cell.comment.text, cell.comment.author),
        None if cell.hyperlink is None else cell.hyperlink.target,
    )


def book_snapshot(path: Path) -> dict[str, object]:
    """Todo lo observable del libro, en tipos comparables con ``==``."""
    wb = load_workbook(path)
    snap: dict[str, object] = {
        "sheets": tuple(ws.title for ws in wb.worksheets),
        "active": wb.active.title,
        "properties": (wb.properties.keywords, wb.properties.title,
                       wb.properties.creator),
        "calc": (wb.calculation.fullCalcOnLoad, wb.calculation.calcMode),
        "names": tuple(sorted((n, d.attr_text, d.localSheetId)
                              for n, d in wb.defined_names.items())),
    }
    for ws in wb.worksheets:
        key = f"sheet:{ws.title}"
        snap[key + ":view"] = (ws.freeze_panes, ws.sheet_view.showGridLines,
                               ws.sheet_view.zoomScale, ws.sheet_properties.tabColor
                               and _color(ws.sheet_properties.tabColor),
                               ws.sheet_properties.outlinePr
                               and (ws.sheet_properties.outlinePr.summaryBelow,
                                    ws.sheet_properties.outlinePr.summaryRight),
                               ws.sheet_format.outlineLevelRow)
        snap[key + ":names"] = tuple(sorted((n, d.attr_text)
                                            for n, d in ws.defined_names.items()))
        snap[key + ":merged"] = tuple(sorted(str(r) for r in ws.merged_cells.ranges))
        snap[key + ":rows"] = tuple(sorted(
            (r, d.outlineLevel, bool(d.hidden), d.height)
            for r, d in ws.row_dimensions.items()
            if d.outlineLevel or d.hidden or d.height is not None))
        snap[key + ":cols"] = tuple(sorted(
            (c, d.width, d.outlineLevel, bool(d.hidden))
            for c, d in ws.column_dimensions.items()))
        snap[key + ":validations"] = tuple(
            (str(v.sqref), v.type, v.formula1) for v in ws.data_validations.dataValidation)
        snap[key + ":conditional"] = tuple(
            (str(cf.sqref), tuple((r.type, tuple(r.formula or ())) for r in cf.rules))
            for cf in ws.conditional_formatting)
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None and not cell.has_style and cell.comment is None:
                    continue
                snap[f"{key}!{cell.coordinate}"] = _cell(cell)
    return snap


def snapshot_diff(expected: dict[str, object], actual: dict[str, object],
                  limit: int = 15) -> list[str]:
    keys = sorted(set(expected) | set(actual))
    out = [f"{k}: esperado {expected.get(k)!r} / obtenido {actual.get(k)!r}"
           for k in keys if expected.get(k) != actual.get(k)]
    return out[:limit] + ([f"... y {len(out) - limit} diferencias mas"] if len(out) > limit else [])


def _build_demo(folder: Path) -> tuple[Path, str]:
    book = folder / "demo.xlsx"
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        B._demo(str(book))
    return book, buf.getvalue()


def _run_cli(*args: str) -> subprocess.CompletedProcess[str]:
    env = dict(os.environ, PYTHONIOENCODING="ascii:strict")
    return subprocess.run([sys.executable, "tools/xlsx_builder.py", *args], cwd=ROOT,
                          env=env, capture_output=True, text=True, timeout=120)


# --- Pruebas ------------------------------------------------------------------

class DemoParityTests(unittest.TestCase):
    """El demo de hoy es identico al de la referencia previa a la particion."""

    def test_demo_workbook_matches_reference_cell_by_cell(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book, _ = _build_demo(Path(tmp))
            expected, actual = book_snapshot(REFERENCE_BOOK), book_snapshot(book)
        self.assertGreater(len(expected), 1000, "la referencia parece vacia")
        self.assertEqual(snapshot_diff(expected, actual), [])

    def test_demo_capture_layer_matches_reference(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            _build_demo(Path(tmp))
            text = (Path(tmp) / "inputs" / "canonical_quarterly.csv").read_text(encoding="utf-8")
        self.assertEqual(text, REFERENCE_CAPTURE.read_text(encoding="utf-8"))

    def test_snapshot_sees_a_one_cell_format_change(self) -> None:
        """La foto no es ciega: un solo formato numerico distinto aparece."""
        with tempfile.TemporaryDirectory() as tmp:
            book, _ = _build_demo(Path(tmp))
            before = book_snapshot(book)
            wb = load_workbook(book)
            wb["Operating"]["D10"].number_format = "0.000"
            wb.save(book)
            diff = snapshot_diff(before, book_snapshot(book))
        self.assertEqual(len(diff), 1, diff)
        self.assertIn("Operating!D10", diff[0])


class FacadeApiTests(unittest.TestCase):

    def test_every_name_before_the_split_is_still_exported(self) -> None:
        missing = sorted(n for n in API_BEFORE_SPLIT if not hasattr(B, n))
        self.assertEqual(missing, [], "nombres que la fachada dejo de exportar")

    def test_documented_import_line_still_works(self) -> None:
        namespace: dict[str, object] = {}
        exec("from xlsx_builder import CellRole, ModelStyler, NumFmt, PeriodHeader, load_brand",
             namespace)
        self.assertIs(namespace["ModelStyler"], B.ModelStyler)


class FacadePatchTests(unittest.TestCase):
    """``mock.patch.object(B, ...)`` alcanza al codigo que lee el nombre."""

    def test_patching_scan_limit_reaches_audit_and_is_restored(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book, _ = _build_demo(Path(tmp))
            with mock.patch.object(B, "_MAX_SCAN_ROWS", 3):
                f20 = {f.check.split()[0]: f for f in B.audit_format(str(book))}["F20"]
            self.assertFalse(f20.ok, f20.detail)
            f20 = {f.check.split()[0]: f for f in B.audit_format(str(book))}["F20"]
        self.assertTrue(f20.ok, "el parche no se restauro en los submodulos")
        self.assertEqual(B._MAX_SCAN_ROWS, 5000)

    def test_patching_load_workbook_reaches_live_checks(self) -> None:
        calls: list[bool] = []
        real = B.load_workbook

        def spy(path: str, data_only: bool = False, **kwargs: object) -> object:
            calls.append(data_only)
            return real(path, data_only=data_only, **kwargs)

        with tempfile.TemporaryDirectory() as tmp:
            book, _ = _build_demo(Path(tmp))
            with mock.patch.object(B, "load_workbook", side_effect=spy):
                B.live_checks_finding(str(book))
                B.audit_format(str(book))
        self.assertIn(True, calls)
        self.assertIn(False, calls)
        self.assertIs(B.load_workbook, real)

    def test_patching_lint_reaches_the_styler(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        with mock.patch.object(B, "lint_formula", return_value=["parche visto"]):
            with self.assertRaisesRegex(ValueError, "parche visto"):
                styler.set_cell(ws, "C5", "=SUM(C1:C4)", B.CellRole.FORMULA)
        styler.set_cell(ws, "C5", "=SUM(C1:C4)", B.CellRole.FORMULA)


class CliParityTests(unittest.TestCase):
    """Mismos textos y exit codes que el builder de un solo archivo."""

    def test_demo_and_audit_print_the_reference_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book = str(Path(tmp) / "demo.xlsx")
            for args, reference, code in ((("demo", book), REFERENCE_DEMO_OUT, 0),
                                          (("audit", book), REFERENCE_AUDIT_OUT, 0)):
                with self.subTest(cmd=args[0]):
                    done = _run_cli(*args)
                    self.assertEqual(done.returncode, code, done.stderr)
                    self.assertEqual(done.stdout.replace(book, BOOK_TOKEN),
                                     reference.read_text(encoding="utf-8"))

    def test_bad_usage_prints_usage_and_exits_2(self) -> None:
        for args in ((), ("demo",), ("demo", "x.xlsx", "--external"), ("audit",),
                     ("otro", "x.xlsx")):
            with self.subTest(args=args):
                done = _run_cli(*args)
                self.assertEqual(done.returncode, 2)
                self.assertEqual(done.stdout,
                                 "uso: python xlsx_builder.py audit <path.xlsx> "
                                 "[brand/DESIGN.md] [--external] | demo <path.xlsx>\n")


def _regenerate() -> None:
    """Reescribe las referencias con el builder actual (solo a proposito)."""
    FIXTURES.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        book = str(Path(tmp) / "demo.xlsx")
        demo = _run_cli("demo", book)
        audit = _run_cli("audit", book)
        if demo.returncode != 0:
            raise SystemExit(f"[x] demo fallo: {demo.stderr}")
        REFERENCE_BOOK.write_bytes(Path(book).read_bytes())
        REFERENCE_CAPTURE.write_bytes((Path(tmp) / "inputs" / "canonical_quarterly.csv").read_bytes())
        for out, done in ((REFERENCE_DEMO_OUT, demo), (REFERENCE_AUDIT_OUT, audit)):
            with out.open("w", encoding="utf-8", newline="") as fh:
                fh.write(done.stdout.replace(book, BOOK_TOKEN))
    print(f"[ok] referencias reescritas en {FIXTURES}")


if __name__ == "__main__":
    if sys.argv[1:] == ["--regenerate"]:
        _regenerate()
    else:
        unittest.main()
