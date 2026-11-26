"""Pruebas de tools/deck_builder.py: numeros con procedencia, bloqueo de
cifras sueltas, auditoria contra el modelo y medicion de espacio.

El libro es sintetico (valores, sin datos de emisoras). El navegador se
simula; ``BrowserTests`` lo corre de verdad solo con RA_BROWSER_TESTS=1.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import deck_builder as D  # noqa: E402
import xlsx_builder as B  # noqa: E402

AS_OF = "2026-01-01"
MODEL_REL = f"model/x-model_{AS_OF}_v1.xlsx"
REPORT_REL = f"research/industry/industry-report_{AS_OF}.md"
FACTS_REL = f"research/notes/x-facts_{AS_OF}.json"
YEARS = (2022, 2023, 2024, 2025, 2026, 2027)   # observado hasta 2024


def _write_model(path: Path) -> None:
    """Libro minimo de valores con la forma que lee el tool."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Annual"
    for i, year in enumerate(YEARS):
        cell = ws.cell(row=4, column=3 + i, value=year)
        cell.number_format = (B.NumFmt.YEAR_A if year <= 2024 else B.NumFmt.YEAR_E).value

    def row(r: int, label: str, values: tuple[object, ...]) -> None:
        ws.cell(row=r, column=2, value=label)
        for i, value in enumerate(values):
            ws.cell(row=r, column=3 + i, value=value)

    def band(r: int, title: str) -> None:
        ws.cell(row=r, column=2, value=title).font = Font(bold=True, size=16)

    row(6, "Ingresos", (100.0, 120.0, 150.0, 180.0, 210.0, 240.0))
    row(7, "Utilidad neta", (10.0, 15.0, 20.0, 25.0, 30.0, 36.0))
    band(9, "Balance")
    row(10, "Total", (1.0, 2.0, 3.0, 4.0, 5.0, 6.0))
    band(12, "Flujos")
    row(13, "Total", (7.0, 8.0, 9.0, 10.0, 11.0, 12.0))
    band(14, "Otros")
    row(15, "Sin calcular", ("=C6*2",))
    row(16, "Parcial", (5.0, 6.0, 7.0))
    row(17, "Cambio", (4.0, -2.0, 3.0, -1.0, 2.0, 1.0))
    dcf = wb.create_sheet("DCF")
    for r, (label, values) in enumerate((("WACC", (0.09,)), ("Valor por accion", (120.5,)),
                                         ("Dos", (1.0, 2.0)), ("Texto", ("n/a",)),
                                         ("Bajo", (100.0,)), ("Alto", (140.0,)),
                                         ("Precio", (110.0,))), start=5):
        dcf.cell(row=r, column=2, value=label)
        for i, value in enumerate(values):
            dcf.cell(row=r, column=3 + i, value=value)
    sens = wb.create_sheet("Val Sens")
    for r in range(3):
        for c in range(3):
            sens.cell(row=6 + r, column=3 + c, value=100.0 + 10 * r + c)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _coverage(root: Path) -> None:
    """Cobertura sintetica: modelo, un documento de texto y sus hechos."""
    _write_model(root / MODEL_REL)
    report = root / REPORT_REL
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text("Participacion: 31.5% del mercado.\n", encoding="utf-8")
    (root / "filings").mkdir()
    (root / "filings" / "anual.pdf").write_bytes(b"%PDF-1.4")
    facts = root / FACTS_REL
    facts.parent.mkdir(parents=True, exist_ok=True)
    facts.write_text(json.dumps({"as_of": AS_OF, "facts": [
        _fact("share", 31.5, REPORT_REL, "31.5%"),
        _fact("plants", 12, "filings/anual.pdf", "12 plantas", unit="plantas"),
    ]}), encoding="utf-8")


def _fact(key: str, value: float, doc: str, quote: str, unit: str = "%",
          tag: str = "observado") -> dict[str, object]:
    return {"key": key, "value": value, "unit": unit, "tag": tag, "doc": doc,
            "locator": "p. 1", "quote": quote}


class _Case(unittest.TestCase):

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        _coverage(self.root)

    def deck(self, **kwargs: object) -> D.Deck:
        options: dict[str, object] = {"root": self.root, "model": MODEL_REL, "ticker": "X",
                                      "as_of": AS_OF, "facts": FACTS_REL}
        options.update(kwargs)
        return D.Deck(**options)  # type: ignore[arg-type]


class NumTests(unittest.TestCase):

    def test_arithmetic_keeps_the_whole_expression(self) -> None:
        a, b = D.Num(120.0, "Annual!D6"), D.Num(100.0, "Annual!C6")
        growth = a / b - 1
        self.assertAlmostEqual(growth.value, 0.2)
        self.assertEqual(growth.src, "((Annual!D6/Annual!C6)-1)")

    def test_constants_on_the_left_are_kept_in_order(self) -> None:
        a = D.Num(4.0, "S!A1")
        self.assertEqual((1 - a).src, "(1-S!A1)")
        self.assertEqual((2 * a).value, 8.0)
        self.assertEqual((a / 1e6).src, "(S!A1/1000000)")
        self.assertEqual((-a).src, "(-S!A1)")
        self.assertEqual(abs(D.Num(-3.0, "S!A2")).value, 3.0)

    def test_division_by_zero_is_a_deck_error(self) -> None:
        with self.assertRaises(D.DeckError):
            D.Num(1.0, "S!A1") / D.Num(0.0, "S!A2")

    def test_formatting_renders_a_span_with_its_origin(self) -> None:
        n = D.Num(0.384, "((Annual!D6/Annual!C6)-1)")
        rendered = D.render_markers(f"<p>Crece {n:+.0%} en el periodo</p>")
        self.assertEqual(
            rendered,
            '<p>Crece <span class="n" data-src="((Annual!D6/Annual!C6)-1)" '
            'data-fmt="+.0%">+38%</span> en el periodo</p>')

    def test_alignment_format_specs_do_not_break_the_markup(self) -> None:
        n = D.Num(5.0, "S!A1")
        rendered = D.render_markers(f"<p>{n:>6.1f}|{n:<6.1f}</p>")
        self.assertIn('data-fmt="&gt;6.1f">   5.0</span>', rendered)
        self.assertIn('data-fmt="&lt;6.1f">5.0   </span>', rendered)

    def test_inside_svg_the_number_is_a_tspan(self) -> None:
        n = D.Num(7.0, "S!A1")
        rendered = D.render_markers(f"<svg><text>{n:.0f}</text></svg><p>{n:.0f}</p>")
        self.assertIn('<text><tspan class="n" data-src="S!A1" data-fmt=".0f">7</tspan></text>',
                      rendered)
        self.assertIn('<p><span class="n"', rendered)

    def test_inside_an_attribute_the_number_is_plain_text(self) -> None:
        n = D.Num(7.0, "S!A1")
        self.assertEqual(D.render_markers(f'<p title="vale {n:.0f}">x</p>'),
                         '<p title="vale 7">x</p>')

    def test_lit_needs_a_reason_and_is_marked(self) -> None:
        with self.assertRaises(D.DeckError):
            D.lit("4", why=" ")
        rendered = D.render_markers(f"<p>{D.lit('4', why='pasos del metodo')} pasos</p>")
        self.assertEqual(rendered,
                         '<p><span class="lit" data-why="pasos del metodo">4</span> pasos</p>')

    def test_escape_keeps_markers_and_escapes_the_rest(self) -> None:
        n = D.Num(1.0, "'Val Sens'!C6")
        rendered = D.render_markers("<p>" + D._esc(f"A & B <{n:.0f}>") + "</p>")
        self.assertIn("A &amp; B &lt;", rendered)
        self.assertIn('data-src="&#x27;Val Sens&#x27;!C6"', rendered)


class ModelBookTests(_Case):

    def setUp(self) -> None:
        super().setUp()
        self.book = D.ModelBook(self.root / MODEL_REL)

    def test_series_reads_a_row_over_the_period_axis(self) -> None:
        rev = self.book.series("Annual", "Ingresos")
        self.assertEqual(rev.keys, list(YEARS))
        self.assertEqual(rev[2027].value, 240.0)
        self.assertEqual(rev[2027].src, "Annual!H6")
        self.assertTrue(rev.is_actual(2024))
        self.assertFalse(rev.is_actual(2025))

    def test_unknown_label_is_an_error(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "no hay fila con etiqueta 'Ventas'"):
            self.book.series("Annual", "Ventas")

    def test_duplicate_label_needs_a_section(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "desambigua con section="):
            self.book.series("Annual", "Total")
        self.assertEqual(self.book.series("Annual", "Total", section="Flujos")[2022].value, 7.0)
        self.assertEqual(self.book.series("Annual", "Total", section="Balance")[2022].value, 1.0)

    def test_formula_without_cached_value_stops_with_the_recalc_hint(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "recalc_validate"):
            self.book.series("Annual", "Sin calcular")

    def test_missing_period_is_an_error_on_access(self) -> None:
        partial = self.book.series("Annual", "Parcial")
        self.assertEqual(partial.keys, [2022, 2023, 2024])
        with self.assertRaisesRegex(D.DeckError, "no tiene valor en 2027"):
            partial[2027]

    def test_sheet_without_period_header_is_an_error(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "encabezado de periodos"):
            self.book.series("DCF", "WACC")

    def test_growth_over_and_cagr_are_derived_numbers(self) -> None:
        rev = self.book.series("Annual", "Ingresos")
        ni = self.book.series("Annual", "Utilidad neta")
        self.assertEqual(rev.growth().keys, list(YEARS[1:]))
        self.assertAlmostEqual(rev.growth()[2023].value, 0.2)
        self.assertEqual(rev.growth()[2023].src, "((Annual!D6/Annual!C6)-1)")
        self.assertFalse(rev.growth().is_actual(2025))
        self.assertAlmostEqual(ni.over(rev)[2022].value, 0.1)
        self.assertAlmostEqual(rev.cagr(2022, 2024).value, (150.0 / 100.0) ** 0.5 - 1)
        with self.assertRaises(D.DeckError):
            rev.cagr(2024, 2022)

    def test_cell_reads_the_only_numeric_cell_of_a_row(self) -> None:
        wacc = self.book.cell("DCF", "WACC")
        self.assertEqual((wacc.value, wacc.src), (0.09, "DCF!C5"))
        with self.assertRaisesRegex(D.DeckError, "indica column="):
            self.book.cell("DCF", "Dos")
        self.assertEqual(self.book.cell("DCF", "Dos", column="d").value, 2.0)
        with self.assertRaisesRegex(D.DeckError, "no es un numero"):
            self.book.cell("DCF", "Texto", column="C")

    def test_range_reads_a_grid_and_quotes_sheet_names_with_spaces(self) -> None:
        grid = self.book.range("Val Sens", "C6:E8")
        self.assertEqual((len(grid), len(grid[0])), (3, 3))
        self.assertEqual(grid[2][1].value, 121.0)
        self.assertEqual(grid[2][1].src, "'Val Sens'!D8")

    def test_axis_years_and_missing_pieces(self) -> None:
        self.assertEqual(self.book.axis_years(), set(YEARS))
        with self.assertRaisesRegex(D.DeckError, "no tiene la hoja"):
            self.book.cell("Nope", "WACC")
        with self.assertRaisesRegex(D.DeckError, "no existe el modelo"):
            D.ModelBook(self.root / "model" / "otro.xlsx")
