"""Pruebas de tools/deck_builder.py: numeros con procedencia, bloqueo de
cifras sueltas, auditoria contra el modelo y medicion de espacio.

El libro es sintetico (valores, sin datos de emisoras). El navegador se
simula; ``BrowserTests`` lo corre de verdad solo con RA_BROWSER_TESTS=1.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import json
import os
import re
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

    def test_format_spec_is_restricted_to_sign_grouping_precision_and_type(self) -> None:
        n = D.Num(240.0, "Annual!H6")
        for spec in ("9>6,.0f", ">8.1f", "<6.1f", "08.1f", "10", ".25f", "s", "zz"):
            with self.subTest(spec=spec), self.assertRaisesRegex(D.DeckError, "formato"):
                format(n, spec)
        for spec in ("", ",.0f", "+.1%", ".2f", "_.0f", " .1f", ".3g", "e"):
            with self.subTest(spec=spec):
                self.assertIn("\ue000", format(n, spec))

    def test_results_outside_the_real_numbers_are_deck_errors(self) -> None:
        for operation in (lambda: D.Num(-4.0, "S!A1") ** 0.5,
                          lambda: D.Num(1e200, "S!A1") * 1e200,
                          lambda: D.Num(10.0, "S!A1") ** 4000):
            with self.assertRaises(D.DeckError):
                operation()

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

    def test_str_is_readable_ascii_and_the_fstring_still_binds(self) -> None:
        n = D.Num(240.0, "Annual!H6")
        self.assertEqual(str(n), "240.0 [Annual!H6]")
        str(D.Num(1.0, "'A\u00f1o'!C5")).encode("cp1252")          # no revienta
        self.assertTrue(str(D.Num(1.0, "'Hoja \u4e2d'!C5")).isascii())
        self.assertEqual(str(D.Num(1e16, "Annual!H6")), "10000000000000000 [Annual!H6]")
        self.assertEqual(str(D.Num(1e-7, "Annual!H6")), "0.0000001 [Annual!H6]")
        for text in (f"{n}", f"{n:,.0f}", "{}".format(n), format(n, "")):
            self.assertIn('data-src="Annual!H6"', D.render_markers(f"<p>{text}</p>"))
        self.assertEqual(repr(n), "Num(value=240.0, src='Annual!H6')")
        self.assertTrue(repr(D.Num(1.0, "'\u0394 WC'!C5")).isascii())

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

    def test_section_is_matched_by_the_start_of_its_band_title(self) -> None:
        # En un modelo real la banda dice 'Assumptions (unica zona de input del libro)'.
        path = self.root / MODEL_REL
        wb = load_workbook(path)
        wb["Annual"]["B12"] = "Flujos (valores en TWD millones)"
        wb.save(path)
        book = D.ModelBook(path)
        self.assertEqual(book.series("Annual", "Total", section="flujos")[2022].value, 7.0)
        self.assertEqual(book.series("Annual", "Total", section="Balance")[2022].value, 1.0)
        with self.assertRaisesRegex(D.DeckError, "no existe la seccion 'Caja'"):
            book.series("Annual", "Total", section="Caja")

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

    def test_apply_derives_a_series_point_by_point(self) -> None:
        rev = self.book.series("Annual", "Ingresos")
        index = rev.apply(lambda n: n / rev[2022] * 100, "Ingresos (indice)")
        self.assertEqual(index.keys, list(YEARS))
        self.assertAlmostEqual(index[2024].value, 150.0)
        self.assertEqual(index[2024].src, "((Annual!E6/Annual!C6)*100)")
        self.assertFalse(index.is_actual(2026))
        with self.assertRaisesRegex(D.DeckError, "debe devolver un Num"):
            rev.apply(lambda n: n.value, "mal")

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


class FactsTests(_Case):

    def _load(self, facts: list[dict[str, object]]) -> dict[str, D.Fact]:
        path = self.root / FACTS_REL
        path.write_text(json.dumps({"facts": facts}), encoding="utf-8")
        return D.load_facts(path, self.root)

    def test_text_source_with_the_quote_is_verified(self) -> None:
        facts = D.load_facts(self.root / FACTS_REL, self.root)
        self.assertEqual(facts["share"].value, 31.5)
        self.assertFalse(facts["share"].manual)

    def test_non_text_source_is_accepted_as_manual(self) -> None:
        self.assertTrue(D.load_facts(self.root / FACTS_REL, self.root)["plants"].manual)

    def test_html_source_is_matched_on_its_visible_text(self) -> None:
        # Un 20-F en iXBRL parte cada cifra con etiquetas: la cita se coteja
        # contra el texto que ve el lector, no contra el HTML crudo.
        page = self.root / "filings" / "anual.htm"
        page.write_text("<p>customers accounted for <ix:nonFraction>70</ix:nonFraction>% ,\n"
                        " <span>76</span>% and <b>78</b>% of net revenue &amp; more</p>",
                        encoding="utf-8")
        facts = self._load([_fact("top10", 78, "filings/anual.htm",
                                  "70%, 76% and 78% of net revenue & more")])
        self.assertFalse(facts["top10"].manual)
        with self.assertRaisesRegex(D.DeckError, "no aparece en"):
            self._load([_fact("top10", 79, "filings/anual.htm", "70%, 76% and 79%")])

    def test_quote_may_span_wrapped_lines(self) -> None:
        (self.root / REPORT_REL).write_text("TSMC subio de 70.4% (4T25)\n   a 72.5% (2T26).\n",
                                            encoding="utf-8")
        facts = self._load([_fact("share", 72.5, REPORT_REL, "de 70.4% (4T25) a 72.5%")])
        self.assertEqual(facts["share"].value, 72.5)

    def test_quote_absent_from_a_text_source_is_an_error(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "no aparece en"):
            self._load([_fact("share", 72.5, REPORT_REL, "72.5%")])

    def test_missing_document_is_an_error(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "no existe el documento"):
            self._load([_fact("share", 31.5, "research/industry/nope.md", "31.5%")])

    def test_schema_violations_are_errors(self) -> None:
        bad_tag = _fact("share", 31.5, REPORT_REL, "31.5%", tag="supuesto")
        no_locator = dict(_fact("share", 31.5, REPORT_REL, "31.5%"), locator="")
        text_value = dict(_fact("share", 31.5, REPORT_REL, "31.5%"), value="31.5")
        ok = _fact("share", 31.5, REPORT_REL, "31.5%")
        for facts, message in (([bad_tag], "tag debe ser"), ([no_locator], "faltan campos"),
                               ([text_value], "value debe ser un numero"),
                               ([ok, ok], "duplicado"),
                               ([dict(ok, key="a b")], "la clave solo admite"),
                               ([dict(ok, key="a-b")], "la clave solo admite"),
                               ([dict(ok, key="a.b")], "la clave solo admite"),
                               ([dict(ok, quote=" ")], "cita demasiado corta"),
                               ([dict(ok, doc=7)], "doc debe ser texto"),
                               (["share"], "cada hecho debe ser un objeto")):
            with self.subTest(message=message), self.assertRaisesRegex(D.DeckError, message):
                self._load(facts)

    def test_unreadable_file_is_a_deck_error(self) -> None:
        path = self.root / FACTS_REL
        path.write_text("{no es json", encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "hechos ilegibles"):
            D.load_facts(path, self.root)


class DesignTests(_Case):

    def _design(self, text: str) -> D.Design:
        path = self.root / "brand" / "DESIGN.md"
        path.parent.mkdir(exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return D.load_design(path, self.root)

    def test_no_file_gives_the_neutral_style(self) -> None:
        self.assertEqual(D.load_design(None), D.Design())
        self.assertEqual(D.load_design(self.root / "brand" / "DESIGN.md"), D.Design())

    def test_slots_are_read_and_the_rest_is_prose(self) -> None:
        design = self._design(
            "# Marca\n\nbrand_primary: #132E57\n"
            "deck_background: #0b1f3a   # fondo oscuro\n"
            "deck_text: #FFFFFF\ndeck_series_2: #ffcc00\n"
            "deck_font_heading: Source Serif 4\ndeck_margin: 64\ndeck_gap_max = 120\n")
        self.assertEqual(design.background, "#0B1F3A")
        self.assertEqual(design.text, "#FFFFFF")
        self.assertEqual(design.series[1], "#FFCC00")
        self.assertEqual(design.series[0], D.Design().series[0])
        self.assertEqual((design.font_heading, design.margin, design.gap_max),
                         ("Source Serif 4", 64, 120))

    def test_malformed_slot_is_an_error(self) -> None:
        for text, message in (("deck_text: azul\n", "se espera #RRGGBB"),
                              ("deck_margin: ancho\n", "se espera un entero"),
                              ("deck_font_body: Arial; x\n", "nombre de fuente invalido"),
                              ("deck_logo: brand/nope.svg\n", "se espera un .svg o .png")):
            with self.subTest(text=text), self.assertRaisesRegex(D.DeckError, message):
                self._design(text)

    def test_logo_is_embedded_as_a_data_uri(self) -> None:
        (self.root / "brand").mkdir()
        (self.root / "brand" / "logo.svg").write_text("<svg/>", encoding="utf-8")
        design = self._design("deck_logo: brand/logo.svg\n")
        self.assertTrue(design.logo.startswith("data:image/svg+xml;base64,"))

    def test_contrast_follows_wcag(self) -> None:
        self.assertAlmostEqual(D.contrast("#000000", "#FFFFFF"), 21.0)
        self.assertAlmostEqual(D.contrast("#FFFFFF", "#FFFFFF"), 1.0)

    def test_pale_series_color_is_flagged_as_weak(self) -> None:
        self.assertEqual(D.weak_series(D.Design()), set())
        self.assertEqual(D.weak_series(self._design("deck_series_2: #F2F2F2\n")), {1})

    def test_design_findings_warn_and_never_block(self) -> None:
        self.assertTrue(all(f.ok for f in D.design_findings(D.Design())))
        text, cvd = D.design_findings(self._design(
            "deck_text: #BBBBBB\ndeck_series_1: #C0392B\ndeck_series_2: #6B8E23\n"))
        self.assertTrue(text.warning and not text.blocking)
        self.assertTrue(cvd.warning and not cvd.blocking)
        self.assertIn("1-2", cvd.detail)


class ScanTests(unittest.TestCase):

    def test_loose_numbers_rule(self) -> None:
        years = {2025, 2026}
        cases = {
            "crece 38% al anio": ["38%"],
            "multiplo de 11x y 2.5x": ["11x", "2.5x"],
            "precio de US$5 y NT$2,150": ["US$5", "NT$2,150"],
            "nodo N2 de 3nm, forma 20-F y 10-K": [],
            "en 2025 y 2026": [],
            "en 2019": ["2019"],
            "rango 10-15% (aprox.)": ["10", "15%"],
            "de 2025-2026": [],
            "1Q2026E fue mejor": [],
            "tres plantas": [],
            "cae -4.2 pp.": ["-4.2"],
            "entre 10%-12%, $5-$6 y 5x-6x": ["10%", "12%", "$5", "$6", "5x", "6x"],
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(D.loose_numbers(text, years), expected)

    def test_symbols_units_and_odd_digits_do_not_hide_a_number(self) -> None:
        blocked = ("~38%", ">38%", "<38%", "\u224838%", "\u00b15%", "38%*", "n=30", "-$5",
                   "+$5", "5\u20ac", "5$", "USD5", "$3B", "1.5B", "180pbs", "11\u00d7",
                   "5\u2030", "12mil", "5mdd", "\uff13\uff18\uff05", "38\u200b%", "3\u00ad8%",
                   "10\u00b2", "1'000", "$2025", "A$5")
        for word in blocked:
            with self.subTest(word=ascii(word)):
                self.assertTrue(D.loose_numbers(f"vale {word} hoy", {2025}))

    def test_identifiers_and_bare_axis_years_still_pass(self) -> None:
        for word in ("N2", "A14", "3nm", "20-F", "10-K", "COVID-19", "1Q2026E", "4T2025",
                     "IFRS16", "2025", "(2025)", "2025."):
            with self.subTest(word=word):
                self.assertEqual(D.loose_numbers(f"el {word} sigue", {2025}), [])

    def test_reader_follows_what_a_browser_would_show(self) -> None:
        doc = D.read_deck(
            '<html><head><title>X 99</title></head><body>'
            '<section class="slide raw" data-n="x">'
            '<p class="auto">nota<p>Margen 38%</p>'
            '<span class="lit">44</span>'
            '<svg><title>Vale 55</title><text><![CDATA[66]]></text></svg>'
            '<span class="n" data-src="Annual!C6" data-fmt=".0f">1<i class="auto">00</i></span>'
            '<p class="source">Fuente: <span class="n" data-src="Annual!D6" data-fmt=".0f">'
            '120</span> en p. 45</p>'
            '</section></body></html>')
        self.assertEqual([number for _, number, _ in D.unbound_numbers(doc, set())],
                         ["38%", "44", "55", "66"])
        self.assertEqual([n.text for n in doc.numbers], ["100", "120"])
        self.assertEqual(doc.literals, [])

    def test_document_separates_bound_literal_and_plain_text(self) -> None:
        n = D.Num(0.38, "Annual!D6")
        doc = D.read_deck(D.render_markers(
            '<html data-gap-max="75"><head><title>X 2019</title><style>p{margin:8px}</style>'
            '</head><body><section class="slide content" data-n="3">'
            f"<h2>Crece {n:.0%} y vale 99</h2>"
            f"<p>{D.lit('4', why='pasos')} pasos<br>segunda linea</p>"
            '<p class="source">Fuente: 20-F 2019, p. 45</p>'
            f'<svg><text class="auto">1,500</text><text>{n:.0%}</text><text>77</text>'
            '<rect width="5"/></svg>'
            "</section><script>var a=12;</script></body></html>"))
        self.assertEqual(doc.gap_max, 75)
        self.assertEqual([(b.slide, b.text, b.src, b.fmt) for b in doc.numbers],
                         [(3, "38%", "Annual!D6", ".0%"), (3, "38%", "Annual!D6", ".0%")])
        self.assertEqual(doc.literals, [(3, "4", "pasos")])
        self.assertEqual([(s, n_) for s, n_, _ in D.unbound_numbers(doc, set())],
                         [(3, "99"), (3, "77")])
        self.assertEqual(doc.blocks[0].text, "Crece 38% y vale 99")


class DeckTests(_Case):

    def _minimal(self, deck: D.Deck, title: str = "Ventas al alza") -> None:
        rev = deck.series("Annual", "Ingresos")
        deck.cover("Emisora X", "Inicio de cobertura")
        deck.section("Crecimiento", "Que la mueve")
        deck.slide(title=title, lede="Entradilla.", visual="<p>visual</p>",
                   note=f"Las ventas llegan a {rev[2027]:,.0f}.", source="Fuente: modelo.",
                   takeaway="Cierre.")
        deck.legal()

    def test_build_writes_the_dated_file_under_research_notes(self) -> None:
        deck = self.deck()
        self._minimal(deck)
        out = deck.build()
        self.assertEqual(out, self.root / "research" / "notes" / f"x-deck_{AS_OF}.html")
        text = out.read_text(encoding="utf-8")
        self.assertIn('data-gap-max="90"', text)
        self.assertIn('<section class="slide content" data-n="3" data-kind="content">', text)
        self.assertIn('<span class="n" data-src="Annual!H6" data-fmt=",.0f">240</span>', text)
        self.assertIn('<div class="num auto">01</div>', text)
        self.assertIn("data-measure", text)

    def test_act_number_uses_the_body_font(self) -> None:
        # La fuente de encabezado por defecto (Georgia) trae numerales de estilo
        # antiguo: en el separador de acto "01" se lee "O1".
        deck = self.deck()
        self._minimal(deck)
        css = deck.render().split("<style>", 1)[1].split("</style>", 1)[0]
        rule = re.search(r"\.act \.num\{([^}]*)\}", css)
        self.assertIsNotNone(rule)
        assert rule is not None
        self.assertIn("font-family:Arial", rule.group(1))
        self.assertNotIn("Georgia", rule.group(1))

    def test_same_input_gives_the_same_bytes(self) -> None:
        outputs = []
        for _ in range(2):
            deck = self.deck()
            self._minimal(deck)
            outputs.append(deck.build().read_bytes())
        self.assertEqual(outputs[0], outputs[1])

    def test_loose_number_blocks_the_build(self) -> None:
        deck = self.deck()
        self._minimal(deck, title="Ventas crecen 38%")
        with self.assertRaisesRegex(D.DeckError, r"1 numero\(s\) sin origen") as caught:
            deck.build()
        self.assertIn("lamina 3: '38%'", str(caught.exception))
        self.assertFalse(deck.output_path.exists())

    def test_a_number_passed_through_str_is_not_bound_and_blocks(self) -> None:
        for value in (240.0, 1e16, 1e-7, 2027.0):
            for show in (str, repr):
                deck = self.deck()
                deck.raw(f"<p>Ventas de {show(D.Num(value, 'Annual!H6'))}.</p>",
                         source="Fuente: propia.")
                with self.subTest(value=value, show=show.__name__), \
                        self.assertRaisesRegex(D.DeckError, "sin origen"):
                    deck.build()

    def test_loose_number_inside_custom_svg_blocks_too(self) -> None:
        deck = self.deck()
        deck.raw("<svg viewBox='0 0 10 10'><text>72.5</text></svg>", source="Fuente: propia.")
        with self.assertRaisesRegex(D.DeckError, "sin origen"):
            deck.build()

    def test_lit_fact_and_axis_years_pass(self) -> None:
        deck = self.deck()
        deck.raw(f"<p>En 2027, {D.lit('3', why='condiciones de la tesis')} condiciones y "
                 f"{deck.fact('share'):.1f}% del mercado.</p>", source="Fuente: 20-F 2019, p. 4.")
        text = deck.build().read_text(encoding="utf-8")
        self.assertIn('data-src="fact:share"', text)

    def test_content_slide_requires_note_and_source(self) -> None:
        deck = self.deck()
        with self.assertRaisesRegex(D.DeckError, "falta note"):
            deck.slide(title="T", lede="", visual="<p>v</p>", note=" ", source="Fuente: x.")
        with self.assertRaisesRegex(D.DeckError, "falta source"):
            deck.slide(title="T", lede="", visual="<p>v</p>", note="Nota.", source="")
        with self.assertRaisesRegex(D.DeckError, "raw: falta source"):
            deck.raw("<p>x</p>", source="")

    def test_text_fields_are_escaped(self) -> None:
        deck = self.deck()
        deck.slide(title="A & B <c>", lede="", visual="<p>v</p>", note="Nota.",
                   source="Fuente: x.")
        self.assertIn("<h2>A &amp; B &lt;c&gt;</h2>", deck.render())

    def test_constructor_validates_its_arguments(self) -> None:
        with self.assertRaisesRegex(D.DeckError, "as_of debe ser"):
            self.deck(as_of="26/11/2026")
        with self.assertRaisesRegex(D.DeckError, "ticker invalido"):
            self.deck(ticker="X Y")
        with self.assertRaisesRegex(D.DeckError, "no esta en el registro"):
            self.deck().fact("nope")
        with self.assertRaisesRegex(D.DeckError, "no tiene laminas"):
            self.deck().build()


class ChartTests(_Case):

    def _bound(self, deck: D.Deck, visual: str) -> D.DeckDocument:
        deck.slide(title="Titulo", lede="", visual=visual, note="Nota.", source="Fuente: x.")
        doc = D.read_deck(deck.render())
        self.assertEqual(D.unbound_numbers(doc, deck.book.axis_years()), [])
        return doc

    def test_nice_ticks_cover_the_range_with_round_steps(self) -> None:
        self.assertEqual(D._nice_ticks(0.0, 240.0), [0.0, 100.0, 200.0, 300.0])
        self.assertEqual(D._nice_ticks(-2.0, 4.0), [-2.0, 0.0, 2.0, 4.0])
        self.assertEqual(D._nice_ticks(5.0, 5.0)[0], 5.0)

    def test_bars_label_only_the_requested_periods_and_mark_the_forecast(self) -> None:
        deck = self.deck()
        svg = deck.bars(deck.series("Annual", "Ingresos"), show=(2024, 2027))
        self.assertEqual(svg.count("Proyectado"), 1)
        self.assertEqual(svg.count(f'fill="{deck.design.series[0]}"'), 3)   # observados
        self.assertEqual(svg.count(f'fill="{deck.design.series[1]}"'), 3)   # proyectados
        doc = self._bound(deck, svg)
        self.assertEqual([n.text for n in doc.numbers], ["150", "240"])

    def test_bars_scale_and_negative_color(self) -> None:
        deck = self.deck()
        change = deck.series("Annual", "Cambio")
        svg = deck.bars(change, show=change.keys, fmt="+.1f", scale=2.0)
        self.assertEqual(svg.count(f'fill="{deck.design.negative}"'), 2)
        doc = self._bound(deck, svg)
        self.assertEqual(doc.numbers[1].text, "-1.0")
        self.assertEqual(doc.numbers[1].src, "(Annual!D17/2)")

    def test_weak_series_color_forces_every_label(self) -> None:
        (self.root / "brand").mkdir()
        (self.root / "brand" / "DESIGN.md").write_text("deck_series_2: #F2F2F2\n", encoding="utf-8")
        deck = self.deck(design="brand/DESIGN.md")
        doc = self._bound(deck, deck.bars(deck.series("Annual", "Ingresos")))
        self.assertEqual(len(doc.numbers), len(YEARS))

    def test_lines_name_each_series_and_shade_the_forecast(self) -> None:
        deck = self.deck()
        rev, ni = deck.series("Annual", "Ingresos"), deck.series("Annual", "Utilidad neta")
        svg = deck.lines([("Margen neto", ni.over(rev))], show=(2022, 2027), fmt=".1%")
        self.assertIn("Margen neto</text>", svg)
        self.assertEqual(svg.count("<polyline"), 1)
        self.assertEqual(svg.count("Proyectado"), 1)
        self.assertEqual([n.text for n in self._bound(deck, svg).numbers], ["10.0%", "15.0%"])

    def test_football_binds_low_high_mid_and_marker(self) -> None:
        deck = self.deck()
        svg = deck.football([("DCF", deck.cell("DCF", "Bajo"), deck.cell("DCF", "Alto"),
                              deck.cell("DCF", "Valor por accion"))],
                            marker=("Precio", deck.cell("DCF", "Precio")), fmt=",.1f")
        doc = self._bound(deck, svg)
        self.assertEqual(sorted(n.text for n in doc.numbers), ["100.0", "110.0", "120.5", "140.0"])

    def test_football_row_without_a_range_shows_its_value_once(self) -> None:
        deck = self.deck()
        value = deck.cell("DCF", "Valor por accion")
        doc = self._bound(deck, deck.football([("Un solo punto", value, value, value)],
                                              fmt=",.1f"))
        self.assertEqual([n.text for n in doc.numbers], ["120.5"])

    def _football_rows(self, deck: D.Deck) -> list[tuple[str, D.Num, D.Num, D.Num]]:
        low, high = deck.cell("DCF", "Bajo"), deck.cell("DCF", "Alto")
        mid = deck.cell("DCF", "Valor por accion")
        return [("Flujos", low, high, mid), ("Multiplos", low, mid, low)]

    def test_football_spreads_its_rows_over_the_height(self) -> None:
        deck = self.deck()
        rows = self._football_rows(deck)

        def centers(svg: str) -> list[float]:
            return [float(v) for v in re.findall(r'<circle cx="[0-9.]+" cy="([0-9.]+)"', svg)]

        def bar_heights(svg: str) -> list[float]:
            return [float(v) for v in re.findall(r'height="([0-9.]+)" fill=', svg)]

        default = deck.football(rows)
        self.assertIn('viewBox="0 0 1100 180"', default)
        self.assertEqual(centers(default), [56.0, 116.0])
        self.assertEqual(bar_heights(default), [24.0, 24.0])
        tall = deck.football(rows, h=420)
        self.assertIn('viewBox="0 0 1100 420"', tall)
        self.assertEqual(centers(tall), [116.0, 296.0])     # paso de 180: (420 - 60) / 2
        self.assertTrue(all(height > 24.0 for height in bar_heights(tall)))
        self._bound(deck, tall)

    def test_football_rejects_a_height_that_does_not_fit_its_rows(self) -> None:
        deck = self.deck()
        with self.assertRaisesRegex(D.DeckError, r"football: h=100 .* 2 fila\(s\); minimo 168"):
            deck.football(self._football_rows(deck), h=100)
        # En el minimo, la etiqueta del primer punto no toca el rotulo del marcador.
        value = deck.cell("DCF", "Valor por accion")
        svg = deck.football([("Uno", value, value, value)], marker=("Precio", value), h=114)
        label_y = float(re.search(r'<text class="val" x="[0-9.]+" y="([0-9.]+)" '
                                  r'text-anchor="middle">(?!Precio)', svg).group(1))
        marker_y = float(re.search(r'y="([0-9.]+)" text-anchor="middle">Precio', svg).group(1))
        self.assertGreaterEqual(label_y - marker_y, 16.0)

    def test_table_text_columns_are_not_right_aligned(self) -> None:
        deck = self.deck()
        wacc = deck.cell("DCF", "WACC")
        table = deck.table(["Concepto", "Unidad", "Valor", "Lectura"],
                           [["Tasa de descuento", "tasa anual", f"{wacc:.1%}", "Sin cambio"],
                            ["Condiciones", "conteo", D.lit("3", why="condiciones"), "n/a"]])
        self.assertIn('<th class="t">Unidad</th>', table)
        self.assertIn('<td class="t">tasa anual</td>', table)
        self.assertIn('<td class="t">Sin cambio</td>', table)
        self.assertIn("<th>Valor</th>", table)             # columna de cifras: a la derecha
        self.assertIn("<th>Concepto</th>", table)          # la primera ya va a la izquierda
        deck.slide(title="Titulo", lede="", visual=table, note="Nota.", source="Fuente: x.")
        css = deck.render().split("<style>", 1)[1].split("</style>", 1)[0]
        self.assertIn("table.grid th.t,table.grid td.t{text-align:left}", css)

    def test_table_accepts_rows_from_an_iterator(self) -> None:
        deck = self.deck()
        table = deck.table(["Concepto", "Valor"], zip(["Uno", "Dos"], ["a", "b"]))
        self.assertEqual(table.count("<td"), 4)

    def test_forecast_tag_sits_above_the_plot_clear_of_value_labels(self) -> None:
        import re
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        for svg in (deck.bars(rev, show=rev.keys), deck.lines([("Ventas", rev)], show=rev.keys)):
            tag_y = float(re.search(r'y="([0-9.]+)">Proyectado', svg).group(1))
            label_ys = [float(y) for y in re.findall(r'class="val" x="[0-9.]+" y="([0-9.]+)"', svg)]
            self.assertTrue(all(tag_y < y - 6 for y in label_ys), (tag_y, label_ys))

    def test_line_labels_at_the_ends_stay_off_the_axis_and_the_series_name(self) -> None:
        import re
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        svg = D.render_markers(deck.lines([("Ventas", rev)], show=(2022, 2024, 2027)))
        anchors = re.findall(r'<text class="val" x="[0-9.]+" y="[0-9.]+" text-anchor="(\w+)"', svg)
        self.assertEqual(anchors, ["start", "middle", "end"])

    def test_short_charts_carry_fewer_axis_marks(self) -> None:
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        marks = lambda svg: svg.count('text-anchor="end">')      # noqa: E731
        self.assertEqual(marks(deck.bars(rev, h=320)), 4)
        self.assertEqual(marks(deck.bars(rev, h=100)), 2)
        # En lineas el eje no parte de cero: un paso redondo puede pedir una marca mas.
        self.assertLessEqual(marks(deck.lines([("Ventas", rev)], h=110)), 3)
        self.assertEqual(D._tick_count(300.0), 4)
        self.assertEqual(D._tick_count(20.0), 1)

    def test_second_series_is_labelled_below_its_line(self) -> None:
        import re
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        svg = D.render_markers(deck.lines([("Arriba", rev), ("Abajo", rev)], show=(2024,)))
        first, second = (float(y) for y in re.findall(
            r'<text class="val" x="[0-9.]+" y="([0-9.]+)"', svg))
        self.assertGreater(second - first, 20.0)
        # Pegada al piso de la grafica, la etiqueta de abajo caeria sobre los
        # rotulos de periodo: ahi vuelve arriba de la linea.
        floor = D.Series("Piso", {k: rev[2022] * 0 for k in rev.keys},
                         {k: True for k in rev.keys})
        svg = D.render_markers(deck.lines([("Arriba", rev), ("Piso", floor)], show=(2024,),
                                          h=320))
        ys = [float(y) for y in re.findall(r'<text class="val" x="[0-9.]+" y="([0-9.]+)"', svg)]
        self.assertLess(max(ys), 320 - 30)

    def test_period_labels_do_not_crowd_the_last_one(self) -> None:
        import re
        deck = self.deck()
        keys = [f"{q}Q{y}" for y in range(2017, 2032) for q in (1, 2, 3, 4)]
        quarterly = D.Series("Margen", {k: D.Num(0.4 + i / 1000, f"Annual!C{6 + i}")
                                        for i, k in enumerate(keys)}, {k: True for k in keys})
        svg = deck.lines([("Margen", quarterly)], h=320)
        xs = sorted(float(x) for x in re.findall(
            r'<text class="auto" x="([0-9.]+)" y="310" text-anchor="middle">', svg))
        self.assertGreaterEqual(len(xs), 6)
        self.assertGreaterEqual(min(b - a for a, b in zip(xs, xs[1:])), 60.0)
        self.assertIn(">4Q2031</text>", svg)

    def test_waterfall_ends_in_the_total_read_from_the_model(self) -> None:
        deck = self.deck()
        low, high = deck.cell("DCF", "Bajo"), deck.cell("DCF", "Alto")
        svg = deck.waterfall([("Base", low), ("Ajuste", high - low)], total=("Total", high))
        self.assertEqual(svg.count("<rect"), 3)
        self.assertEqual(self._bound(deck, svg).numbers[-1].src, "DCF!C10")

    def test_heatmap_and_table_bind_their_cells(self) -> None:
        deck = self.deck()
        grid = deck.range("Val Sens", "C6:E8")
        wacc = deck.cell("DCF", "WACC")
        heat = deck.heatmap(grid, rows=[wacc - 0.01, wacc, wacc + 0.01],
                            cols=["Bajo", "Base", "Alto"], corner="WACC", center=(1, 1))
        self.assertEqual(heat.count("<td"), 9)
        self.assertEqual(heat.count('class="hl"'), 1)
        table = deck.table(["Concepto", "Valor"], [["WACC", f"{wacc:.1%}"]])
        doc = self._bound(deck, heat + table)
        self.assertEqual(len(doc.numbers), 9 + 3 + 1)
        self.assertEqual(doc.numbers[-1].text, "9.0%")


class AuditTests(_Case):

    def _built(self, extra: str = "") -> Path:
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        deck.cover("Emisora X")
        deck.slide(title=f"Ventas crecen {rev.growth()[2027]:+.1%} en 2027",
                   lede="Menos que antes. Sigue creciendo.",
                   visual=deck.bars(rev, show=(2027,)),
                   note=f"Participa con {deck.fact('share'):.1f}% del mercado.{extra}",
                   source="Fuente: modelo y reporte de industria.")
        return deck.build()

    def _audit(self, html_path: Path) -> dict[str, B.Finding]:
        findings, _, _ = D.audit_deck(html_path, self.root / MODEL_REL, self.root / FACTS_REL)
        return {f.check.split()[0]: f for f in findings}

    def test_fresh_deck_is_green(self) -> None:
        result = self._audit(self._built())
        self.assertTrue(all(f.ok for f in result.values()), [f.detail for f in result.values()])
        self.assertIn("3 cifras verificadas", result["P1"].detail)

    def test_altered_number_fails_p1(self) -> None:
        path = self._built()
        path.write_text(path.read_text(encoding="utf-8").replace(">+14.3%<", ">+18.0%<"),
                        encoding="utf-8")
        result = self._audit(path)
        self.assertTrue(result["P1"].blocking)
        self.assertIn("muestra '+18.0%', el origen da '+14.3%'", result["P1"].detail)

    def test_model_changed_after_the_build_fails_p1(self) -> None:
        path = self._built()
        wb = load_workbook(self.root / MODEL_REL)
        wb["Annual"]["H6"] = 300.0
        wb.save(self.root / MODEL_REL)
        result = self._audit(path)
        self.assertTrue(result["P1"].blocking)
        self.assertIn("2 de 3 no coinciden", result["P1"].detail)

    def test_number_typed_into_the_html_fails_p2(self) -> None:
        path = self._built()
        path.write_text(path.read_text(encoding="utf-8").replace(
            "Menos que antes.", "Menos que el 22% anterior."), encoding="utf-8")
        result = self._audit(path)
        self.assertTrue(result["P2"].blocking)
        self.assertIn("lamina 2: '22%'", result["P2"].detail)

    def test_literals_and_manual_facts_are_warnings(self) -> None:
        deck = self.deck()
        deck.raw(f"<p>{D.lit('3', why='condiciones')} condiciones en "
                 f"{deck.fact('plants'):.0f} plantas.</p>", source="Fuente: informe anual.")
        result = self._audit(deck.build())
        self.assertTrue(result["P3"].warning and not result["P3"].blocking)
        self.assertTrue(result["P4"].warning and not result["P4"].blocking)
        self.assertIn("plants", result["P4"].detail)

    def test_evaluate_only_accepts_the_origin_grammar(self) -> None:
        book = D.ModelBook(self.root / MODEL_REL)
        facts = D.load_facts(self.root / FACTS_REL, self.root)
        self.assertAlmostEqual(D.evaluate("((Annual!D6/Annual!C6)-1)", book, facts), 0.2)
        self.assertEqual(D.evaluate("abs((-'Val Sens'!C6))", book, facts), 100.0)
        self.assertEqual(D.evaluate("(fact:share*2)", book, facts), 63.0)
        for src in ("__import__('os')", "Annual!C6 if 1 else 2", "max(Annual!C6)", "x+1",
                    "Annual!C6 +", "fact:nope"):
            with self.subTest(src=src), self.assertRaises(D.DeckError):
                D.evaluate(src, book, facts)

    def test_claims_list_every_sentence_in_order(self) -> None:
        doc = D.read_deck(self._built().read_text(encoding="utf-8"))
        self.assertEqual(D.claims(doc), [
            (1, "texto", "Emisora X"),
            (2, "texto", "Ventas crecen +14.3% en 2027"),
            (2, "texto", "Menos que antes."),
            (2, "texto", "Sigue creciendo."),
            (2, "texto", "Participa con 31.5% del mercado."),
        ])
        table = D.claims_table(doc)
        self.assertIn("| L\u00e1mina | D\u00f3nde | Frase | Evidencia |", table)
        self.assertIn("| 2 | texto | Sigue creciendo. |  |", table)

    def _rich(self) -> Path:
        deck = self.deck()
        rev, ni = deck.series("Annual", "Ingresos"), deck.series("Annual", "Utilidad neta")
        wacc = deck.cell("DCF", "WACC")
        deck.cover("Emisora X", figures=[("Ventas del periodo", f"{rev[2024]:,.0f}")])
        deck.slide(title="Margen estable", lede="", note="Nota de la lamina.",
                   source="Fuente: modelo.",
                   visual=deck.lines([("Margen neto", ni.over(rev))], show=(2027,), fmt=".1%")
                   + deck.football([("Flujos", deck.cell("DCF", "Bajo"), deck.cell("DCF", "Alto"),
                                     deck.cell("DCF", "Valor por accion"))],
                                   marker=("Precio", deck.cell("DCF", "Precio")))
                   + deck.table(["Concepto", "Valor"],
                                [["Tasa de descuento", f"{wacc:.1%}"],
                                 ["Ventas", f"{rev[2027]:,.0f} mm"]]))
        deck.raw("<div>Texto en un div.</div><svg viewBox='0 0 10 10'>"
                 "<text>Etapa uno</text></svg>", source="Fuente: propia.")
        return deck.build()

    def test_claims_include_tables_cover_labels_and_chart_text_with_their_origin(self) -> None:
        listed = D.claims(D.read_deck(self._rich().read_text(encoding="utf-8")))
        for expected in ((1, "portada", "150 Ventas del periodo"),
                         (2, "gr\u00e1fica", "Margen neto"),
                         (2, "gr\u00e1fica", "Flujos"),
                         (2, "gr\u00e1fica", "Precio 110"),
                         (2, "tabla", "Concepto"), (2, "tabla", "Valor"),
                         (2, "tabla", "Tasa de descuento"), (2, "tabla", "240 mm"),
                         (2, "texto", "Nota de la lamina."),
                         (3, "texto", "Texto en un div."),
                         (3, "gr\u00e1fica", "Etapa uno")):
            self.assertIn(expected, listed)
        texts = [text for _, _, text in listed]
        for absent in ("9.0%", "Proyectado", "Nota", "15.0%", "120", "2027"):
            self.assertNotIn(absent, texts)
        # En orden de lectura: la portada antes que la lamina de contenido.
        self.assertEqual([slide for slide, _, _ in listed], sorted(s for s, _, _ in listed))

    def _report(self, path: Path) -> tuple[Path, tuple[int, int, int]]:
        findings, doc, facts = D.audit_deck(path, self.root / MODEL_REL, self.root / FACTS_REL)
        report = self.root / "research" / "notes" / f"x-deck-verify_{AS_OF}.md"
        return report, D.write_report(report, path, [("Cifras", findings)], doc, facts,
                                      model=self.root / MODEL_REL)

    def _carry(self, deck: Path, *rows: str) -> tuple[str, int, int, int]:
        previous = ("| L\u00e1mina | D\u00f3nde | Frase | Evidencia |\n|---|---|---|---|\n"
                    + "".join(row + "\n" for row in rows))
        return D.claims_with_evidence(D.read_deck(deck.read_text(encoding="utf-8")), previous)

    def test_evidence_follows_the_same_claim_not_the_same_words(self) -> None:
        deck = self._rich()
        # Mismas palabras en otro lugar (rotulo de grafica contra encabezado de tabla).
        table, kept, lost, _ = self._carry(deck, "| 2 | gr\u00e1fica | Concepto | de la grafica |")
        self.assertEqual((kept, lost), (0, 1))
        self.assertIn("| 2 | tabla | Concepto |  |", table)
        # La lamina cambio de numero y la frase es unica: es la misma afirmacion.
        table, kept, lost, _ = self._carry(deck, "| 7 | tabla | Tasa de descuento | DCF C5 |")
        self.assertEqual((kept, lost), (1, 0))
        self.assertIn("| 2 | tabla | Tasa de descuento | DCF C5 |", table)

    def test_repeated_claims_keep_evidence_only_on_their_own_slide(self) -> None:
        deck = self.deck()
        for title in ("Primera", "Segunda"):
            deck.slide(title=title, lede="", note="Nota.", source="Fuente: x.",
                       visual=deck.table(["Concepto", "Valor"], [["Fila", "celda"]]))
        path = deck.build()
        table, kept, lost, _ = self._carry(path, "| 1 | tabla | Valor | de la primera |",
                                           "| 2 | tabla | Valor | de la segunda |")
        self.assertEqual((kept, lost), (2, 0))
        self.assertIn("| 1 | tabla | Valor | de la primera |", table)
        self.assertIn("| 2 | tabla | Valor | de la segunda |", table)
        # Renumeradas y repetidas: no se sabe cual es cual, no se adivina.
        table, kept, lost, _ = self._carry(path, "| 5 | tabla | Valor | de la primera |",
                                           "| 6 | tabla | Valor | de la segunda |")
        self.assertEqual((kept, lost), (0, 2))
        self.assertNotIn("de la primera", table)

    def test_evidence_text_survives_pipes_backslashes_and_odd_rows(self) -> None:
        deck = self._built()
        table, kept, lost, _ = self._carry(
            deck, "| 2 | texto | Menos que antes. | serie a | b, sin escapar |",
            "| 2 | texto | Sigue creciendo. | ver C:\\modelos\\|",
            "| 1 | texto | -- | evidencia de una frase rara |")
        self.assertEqual((kept, lost), (2, 1))
        self.assertIn("| 2 | texto | Menos que antes. | serie a | b, sin escapar |", table)
        self.assertIn("| 2 | texto | Sigue creciendo. | ver C:\\modelos\\| |", table)

    def test_evidence_written_against_another_model_is_marked_for_review(self) -> None:
        path = self._built()
        report, _ = self._report(path)
        text = report.read_text(encoding="utf-8")
        self.assertRegex(text, r"Evidencia cotejada contra: modelo `[0-9a-f]{12}`, "
                               r"hechos `[0-9a-f]{12}`\.")
        row = "| 2 | texto | Menos que antes. |  |"
        report.write_text(text.replace(row, row[:-3] + " crecimiento menor al observado |"),
                          encoding="utf-8")
        self.assertEqual(self._report(path)[1], (1, 0, 0))      # mismo modelo: sin marca
        wb = load_workbook(self.root / MODEL_REL)
        wb["Annual"]["C7"] = 11.0          # una celda que el deck no muestra
        wb.save(self.root / MODEL_REL)
        report, counts = self._report(path)
        self.assertEqual(counts, (1, 0, 1))                     # conservada, pero marcada
        marked = "| 2 | texto | Menos que antes. | [revisar] crecimiento menor al observado |"
        self.assertIn(marked, report.read_text(encoding="utf-8"))
        self.assertEqual(self._report(path)[1], (1, 0, 0))      # no se marca dos veces
        self.assertIn(marked, report.read_text(encoding="utf-8"))
        self.assertNotIn("[revisar] [revisar]", report.read_text(encoding="utf-8"))

    def test_report_keeps_the_evidence_of_sentences_that_did_not_change(self) -> None:
        report, counts = self._report(self._built())
        self.assertEqual(counts, (0, 0, 0))
        text = report.read_text(encoding="utf-8")
        for sentence, evidence in (("Sigue creciendo.", "Annual!H6 > Annual!G6"),
                                   ("Menos que antes.", "crecimiento 2024 \\| 2027"),
                                   ("Participa con 31.5% del mercado.", "reporte, p. 1")):
            row = f"| 2 | texto | {sentence} |  |"
            self.assertIn(row, text)
            text = text.replace(row, f"| 2 | texto | {sentence} | {evidence} |")
        report.write_text(text, encoding="utf-8")
        # Misma lamina, otra entradilla: una frase cambia, las demas no.
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        deck.cover("Emisora X")
        deck.slide(title=f"Ventas crecen {rev.growth()[2027]:+.1%} en 2027",
                   lede="Menos que en el tramo observado. Sigue creciendo.",
                   visual=deck.bars(rev, show=(2027,)),
                   note=f"Participa con {deck.fact('share'):.1f}% del mercado.",
                   source="Fuente: modelo y reporte de industria.")
        report, counts = self._report(deck.build())
        self.assertEqual(counts, (2, 1, 0))     # conservadas, perdidas, marcadas
        text = report.read_text(encoding="utf-8")
        self.assertIn("| 2 | texto | Sigue creciendo. | Annual!H6 > Annual!G6 |", text)
        self.assertIn("| 2 | texto | Participa con 31.5% del mercado. | reporte, p. 1 |", text)
        self.assertIn("| 2 | texto | Menos que en el tramo observado. |  |", text)
        self.assertNotIn("crecimiento 2024", text)

    def test_evidence_is_read_from_a_report_of_the_previous_format(self) -> None:
        path = self._built()
        report, _ = self._report(path)
        text = report.read_text(encoding="utf-8")
        head, _, _ = text.partition("## Afirmaciones")
        report.write_text(head + "## Afirmaciones\n\n| L\u00e1mina | Frase | Evidencia |\n"
                          "|---|---|---|\n| 2 | Sigue creciendo. | celda H6 |\n",
                          encoding="utf-8")
        report, counts = self._report(path)
        # Sin huella del modelo en el reporte viejo: se conserva, marcada para revisar.
        self.assertEqual(counts, (1, 0, 1))
        self.assertIn("| 2 | texto | Sigue creciendo. | [revisar] celda H6 |",
                      report.read_text(encoding="utf-8"))

    def test_space_section_is_added_and_replaced_without_touching_the_rest(self) -> None:
        path = self._built()
        report, _ = self._report(path)
        before = report.read_text(encoding="utf-8")
        D.write_space_section(report, path, D.measure_findings([_slide(2, 40.0)], 90))
        text = report.read_text(encoding="utf-8")
        self.assertIn("## Espacio", text)
        self.assertIn("| [ok] | P6 sin lamina vacia | 1 laminas medidas, umbral 90 px |", text)
        self.assertLess(text.index("## Espacio"), text.index("## Cifras y su origen"))
        self.assertLess(text.index("## Cifras\n"), text.index("## Espacio"))
        start = text.index("## Espacio")
        self.assertEqual(text[:start] + text[text.index("## Cifras y su origen"):], before)
        D.write_space_section(report, path, D.measure_findings([_slide(2, 140.0)], 90))
        text = report.read_text(encoding="utf-8")
        self.assertEqual(text.count("## Espacio"), 1)
        self.assertIn("| [x] | P6 sin lamina vacia |", text)
        self.assertNotIn("| [ok] | P6", text)
        self.assertEqual(text[:text.index("## Espacio")]
                         + text[text.index("## Cifras y su origen"):], before)

    def test_space_section_creates_the_report_when_there_is_none(self) -> None:
        path = self._built()
        report = self.root / "research" / "notes" / "nuevo.md"
        D.write_space_section(report, path, D.measure_findings([_slide(2, 40.0)], 90))
        text = report.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("# Verificaci\u00f3n del deck"))
        self.assertIn("| [ok] | P5 sin contenido encimado |", text)

    def test_audit_report_keeps_the_space_section_only_for_the_deck_it_measured(self) -> None:
        path = self._built()
        report, _ = self._report(path)
        D.write_space_section(report, path, D.measure_findings([_slide(2, 40.0)], 90))
        report, _ = self._report(path)             # mismo deck: la medicion sigue valiendo
        text = report.read_text(encoding="utf-8")
        self.assertIn("| [ok] | P6 sin lamina vacia |", text)
        self.assertEqual(text.count("## Espacio"), 1)
        report, _ = self._report(self._built(extra=" Otra frase."))   # el deck cambio
        text = report.read_text(encoding="utf-8")
        self.assertEqual(text.count("## Espacio"), 1)
        self.assertNotIn("| [ok] | P6", text)
        self.assertIn("| [pendiente] | P5-P8 medicion de espacio |", text)
        self.assertIn("measure --report", text)

    def test_report_holds_the_deterministic_sections(self) -> None:
        path = self._built()
        findings, doc, facts = D.audit_deck(path, self.root / MODEL_REL, self.root / FACTS_REL)
        report = self.root / "research" / "notes" / f"x-deck-verify_{AS_OF}.md"
        D.write_report(report, path, [("Cifras", findings)], doc, facts,
                       model=self.root / MODEL_REL)
        text = report.read_text(encoding="utf-8")
        for expected in ("| [ok] | P1 cifras iguales a su origen |",
                         "| 2 | +14.3% | `((Annual!H6/Annual!G6)-1)` | `+.1%` |",
                         "| share | 31.5 % | observado |", "## Afirmaciones",
                         "| 2 | texto | Participa con 31.5% del mercado. |  |"):
            self.assertIn(expected, text)


class LabelsTests(_Case):

    def test_labels_list_rows_bands_numeric_span_and_axis(self) -> None:
        text = D.labels_report(D.ModelBook(self.root / MODEL_REL))
        self.assertTrue(text.startswith(f"# Etiquetas del modelo \u2014 x-model_{AS_OF}_v1.xlsx"))
        annual, dcf = text.split("## Annual", 1)[1].split("## DCF", 1)
        self.assertIn("Periodos (6): `2022` a `2027`, columnas C a H; proyectado desde `2025`.",
                      annual)
        for row in ("| 6 | Ingresos |  | 6 | C a H |  |",
                    "| 9 | Balance | secci\u00f3n | 0 |  |  |",
                    "| 10 | Total |  | 6 | C a H | repetida: usa `section=` |",
                    "| 13 | Total |  | 6 | C a H | repetida: usa `section=` |",
                    "| 15 | Sin calcular |  | 0 |  | 1 f\u00f3rmula(s) sin valor: recalcula |",
                    "| 16 | Parcial |  | 3 | C a E |  |"):
            self.assertIn(row, annual)
        self.assertIn("Sin encabezado de periodos", dcf)
        self.assertIn("| 5 | WACC |  | 1 | C |  |", dcf)
        self.assertIn("| 7 | Dos |  | 2 | C a D |  |", dcf)
        self.assertIn("| 8 | Texto |  | 0 |  |  |", dcf)

    def test_label_with_a_line_break_and_repeats_that_section_cannot_split(self) -> None:
        model = self.root / MODEL_REL
        wb = load_workbook(model)
        ws = wb["Annual"]
        ws["B18"], ws["C18"] = "Deuda\nneta", 5.0
        ws["B19"], ws["B20"] = "Doble", "Doble"         # las dos en la seccion Otros
        ws["B5"] = "Total"                              # antes de cualquier banda
        wb.save(model)
        book = D.ModelBook(model)
        text = D.labels_report(book)
        self.assertIn("| 18 | Deuda\\nneta |  | 1 | C | salto de l\u00ednea: "
                      "escr\u00edbelo `\\n` |", text)
        self.assertEqual(book.cell("Annual", "Deuda\nneta").value, 5.0)
        self.assertIn("| 19 | Doble |  | 0 |  | repetida en su secci\u00f3n: usa `deck.range` |",
                      text)
        self.assertIn("| 5 | Total |  | 0 |  | repetida y fuera de secci\u00f3n: usa "
                      "`deck.range` |", text)
        self.assertIn("| 10 | Total |  | 6 | C a H | repetida: usa `section=` |", text)

    def test_band_rule_is_the_one_section_lookup_uses(self) -> None:
        book = D.ModelBook(self.root / MODEL_REL)
        text = D.labels_report(book)
        for title in ("Balance", "Flujos", "Otros"):
            self.assertIsNotNone(D._section_bounds(book._sheet("Annual"), title))
            self.assertRegex(text, rf"\| \d+ \| {title} \| secci\u00f3n \|")
        self.assertIsNone(D._section_bounds(book._sheet("Annual"), "Ingresos"))

    def test_quarterly_axis_is_shown_with_the_keys_the_script_uses(self) -> None:
        model = self.root / MODEL_REL
        wb = load_workbook(model)
        ws = wb.create_sheet("Quarterly")
        for i, key in enumerate(("1Q2026A", "2Q2026A", "3Q2026E", "4Q2026E")):
            ws.cell(row=4, column=3 + i, value=key)
        ws.cell(row=6, column=2, value="Ingresos")
        wb.save(model)
        text = D.labels_report(D.ModelBook(model))
        self.assertIn("Periodos (4): `1Q2026` a `4Q2026`, columnas C a F; "
                      "proyectado desde `3Q2026`.", text)


class FactsCheckTests(_Case):

    def test_every_fact_is_listed_with_its_verification(self) -> None:
        findings = D.check_facts(self.root / FACTS_REL, self.root)
        self.assertEqual([(f.check, f.mark) for f in findings],
                         [("hecho share", "[ok]"), ("hecho plants", "[aviso]")])
        self.assertIn("cita encontrada", findings[0].detail)
        self.assertIn(REPORT_REL, findings[0].detail)
        self.assertIn("manual", findings[1].detail)
        self.assertFalse(any(f.blocking for f in findings))

    def test_every_bad_fact_is_reported_not_only_the_first(self) -> None:
        path = self.root / "hechos.json"
        path.write_text(json.dumps({"facts": [
            _fact("share", 31.5, REPORT_REL, "31.5%"),
            _fact("lejos", 1.0, REPORT_REL, "no esta"),
            _fact("share", 2.0, REPORT_REL, "31.5%"),
            {"key": "corto"}, "suelto"]}), encoding="utf-8")
        findings = D.check_facts(path, self.root)
        self.assertEqual([f.mark for f in findings], ["[ok]", "[x]", "[x]", "[x]", "[x]"])
        self.assertIn("no aparece en", findings[1].detail)
        self.assertNotIn("hecho 'lejos'", findings[1].detail)    # la clave ya va en el check
        self.assertIn("duplicado", findings[2].detail)
        self.assertIn("faltan campos", findings[3].detail)
        # load_facts sigue deteniendose en el primero, con su clave.
        with self.assertRaisesRegex(D.DeckError, "hecho 'lejos': la cita"):
            D.load_facts(path, self.root)

    def test_a_broken_file_is_still_a_deck_error(self) -> None:
        path = self.root / "hechos.json"
        path.write_text("[]", encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "lista 'facts'"):
            D.check_facts(path, self.root)


def _slide(n: int, gap: float, **extra: object) -> dict[str, object]:
    data: dict[str, object] = {"n": n, "kind": "content", "gap": gap, "visual_h": 300.0,
                               "overflow": [], "svg_out": []}
    data.update(extra)
    return data


class MeasureTests(_Case):

    def _marks(self, data: list[dict[str, object]], gap_max: int = 90) -> dict[str, B.Finding]:
        return {f.check.split()[0]: f for f in D.measure_findings(data, gap_max)}

    def test_clean_slides_are_green_and_fixed_templates_are_exempt(self) -> None:
        marks = self._marks([{"n": 1, "kind": "cover"}, _slide(2, 40.0), _slide(3, 89.0)])
        self.assertTrue(all(f.ok for f in marks.values()))
        self.assertIn("2 laminas medidas", marks["P5"].detail)

    def test_each_defect_fails_its_own_check(self) -> None:
        marks = self._marks([
            _slide(2, -12.0), _slide(3, 140.0),
            _slide(4, 30.0, overflow=["h2: Titulo muy largo"]),
            _slide(5, 30.0, svg_out=["1,380"]),
        ])
        self.assertIn("lamina 2: -12 px", marks["P5"].detail)
        self.assertIn("lamina 3: 140 px libres (el visual mide 300 px)", marks["P6"].detail)
        self.assertIn("lamina 4: h2: Titulo muy largo", marks["P7"].detail)
        self.assertIn("lamina 5: '1,380'", marks["P8"].detail)
        self.assertTrue(all(f.blocking for f in marks.values()))

    def test_overlapping_chart_labels_fail_p8(self) -> None:
        marks = self._marks([_slide(2, 30.0, svg_overlap=["55% / 60%"])])
        self.assertTrue(marks["P8"].blocking)
        self.assertIn("lamina 2: encimadas '55% / 60%'", marks["P8"].detail)

    def test_threshold_comes_from_the_deck(self) -> None:
        self.assertTrue(self._marks([_slide(2, 100.0)], gap_max=120)["P6"].ok)

    def _deck_file(self) -> Path:
        deck = self.deck()
        deck.slide(title="Titulo", lede="", visual="<p>visual</p>", note="Nota.",
                   source="Fuente: x.")
        return deck.build()

    def test_without_a_browser_the_result_is_pending_never_green(self) -> None:
        with mock.patch.object(D, "find_browser", return_value=None):
            findings = D.measure_deck(self._deck_file())
        self.assertEqual(len(findings), 1)
        self.assertTrue(findings[0].pending and not findings[0].ok)

    def test_measure_reads_what_the_browser_dumps(self) -> None:
        dump = ('<html data-measure="[{&quot;n&quot;:1,&quot;kind&quot;:&quot;content&quot;,'
                '&quot;gap&quot;:-5,&quot;visual_h&quot;:20,&quot;overflow&quot;:[],'
                '&quot;svg_out&quot;:[]}]"><body></body></html>')
        with mock.patch.object(D, "find_browser", return_value="navegador"), \
                mock.patch.object(D, "_run_browser", return_value=dump) as run:
            findings = D.measure_deck(self._deck_file())
        self.assertIn("--dump-dom", run.call_args.args[1])
        self.assertTrue(findings[0].blocking)
        self.assertIn("lamina 1: -5 px", findings[0].detail)

    def test_a_browser_that_does_not_measure_is_a_failure(self) -> None:
        with mock.patch.object(D, "find_browser", return_value="navegador"), \
                mock.patch.object(D, "_run_browser", return_value="<html></html>"):
            findings = D.measure_deck(self._deck_file())
        self.assertTrue(findings[0].blocking)
        self.assertIn("no dejo medicion", findings[0].detail)

    def test_forced_browser_path_that_does_not_exist_means_no_browser(self) -> None:
        with mock.patch.dict(os.environ, {"RA_BROWSER": str(self.root / "nope.exe")}):
            self.assertIsNone(D.find_browser())


# El navegador de verdad es opcional: en CI puede existir y no arrancar (sandbox).
REAL_BROWSER = os.environ.get("RA_BROWSER_TESTS") == "1" and bool(D.find_browser())
REAL_BROWSER_REASON = "requiere RA_BROWSER_TESTS=1 y Edge, Chrome o Chromium"


@unittest.skipUnless(REAL_BROWSER, REAL_BROWSER_REASON)
class BrowserTests(_Case):
    """El navegador de verdad: cada defecto de espacio falla su check."""

    def test_overfull_empty_and_overflowing_slides_fail(self) -> None:
        deck = self.deck()
        tall = "<div style='height:700px'>alto</div>"
        deck.slide(title="Encimada", lede="", visual=tall, note="Nota.", source="Fuente: x.")
        deck.slide(title="Vacia", lede="", visual="<p>poco</p>", note="Nota.", source="Fuente: x.")
        deck.slide(title="Ancha", lede="", note="Nota.", source="Fuente: x.",
                   visual="<div style='width:1500px;height:330px'>ancho</div>")
        deck.slide(title="Etiqueta", lede="", note="Nota.", source="Fuente: x.",
                   visual="<svg viewBox='0 0 1100 330'><text x='1090' y='40'>etiqueta "
                          "larga fuera</text></svg>")
        marks = {f.check.split()[0]: f for f in D.measure_deck(deck.build())}
        self.assertIn("lamina 1", marks["P5"].detail)
        self.assertIn("lamina 2", marks["P6"].detail)
        self.assertIn("lamina 3", marks["P7"].detail)
        self.assertIn("lamina 4", marks["P8"].detail)


def _cli(args: list[str], env: dict[str, str] | None = None) -> subprocess.CompletedProcess[bytes]:
    full = dict(os.environ, PYTHONIOENCODING="ascii:strict", **(env or {}))
    return subprocess.run([sys.executable, str(ROOT / "tools" / "deck_builder.py"), *args],
                          cwd=ROOT, env=full, capture_output=True, timeout=180)


class CliTests(_Case):

    def test_demo_builds_audits_green_and_prints_ascii(self) -> None:
        done = _cli(["demo", str(self.root)])
        self.assertEqual(done.returncode, 0, done.stdout.decode("ascii", "replace")
                         + done.stderr.decode("ascii", "replace"))
        self.assertTrue(done.stdout.isascii())
        self.assertIn(b"Veredicto: VERDE", done.stdout)
        self.assertTrue((self.root / "demo" / "research" / "notes"
                         / f"demo-deck_{D.DEMO_AS_OF}.html").is_file())

    def test_audit_exit_code_follows_the_verdict_and_writes_the_report(self) -> None:
        self.assertEqual(_cli(["demo", str(self.root)]).returncode, 0)
        notes = self.root / "demo" / "research" / "notes"
        deck = notes / f"demo-deck_{D.DEMO_AS_OF}.html"
        model = self.root / "demo" / "model" / f"demo-model_{D.DEMO_AS_OF}_v1.xlsx"
        facts = notes / f"demo-facts_{D.DEMO_AS_OF}.json"
        report = notes / f"demo-deck-verify_{D.DEMO_AS_OF}.md"
        args = ["audit", str(deck), str(model), "--facts", str(facts), "--report", str(report)]
        self.assertEqual(_cli(args).returncode, 0)
        self.assertIn("## Afirmaciones", report.read_text(encoding="utf-8"))
        deck.write_text(deck.read_text(encoding="utf-8").replace(">1,380<", ">1,999<"),
                        encoding="utf-8")
        broken = _cli(args)
        self.assertEqual(broken.returncode, 1)
        self.assertIn(b"[x] P1 cifras iguales a su origen", broken.stdout)

    def test_audit_with_design_lists_contrast_warnings_without_failing(self) -> None:
        design = self.root / "DESIGN.md"
        design.write_text("deck_text: #BBBBBB" + chr(10), encoding="utf-8")
        deck = self.deck(design="DESIGN.md")
        deck.slide(title="Titulo", lede="", visual="<p>visual</p>", note="Nota.",
                   source="Fuente: x.")
        done = _cli(["audit", str(deck.build()), str(self.root / MODEL_REL),
                     "--facts", str(self.root / FACTS_REL), "--design", str(design)])
        self.assertEqual(done.returncode, 0)
        self.assertIn(b"-- Diseno", done.stdout)
        self.assertIn(b"[aviso] P9 contraste del texto", done.stdout)

    def test_measure_without_a_browser_exits_pending(self) -> None:
        self.assertEqual(_cli(["demo", str(self.root)]).returncode, 0)
        deck = self.root / "demo" / "research" / "notes" / f"demo-deck_{D.DEMO_AS_OF}.html"
        done = _cli(["measure", str(deck)], env={"RA_BROWSER": str(self.root / "nope.exe")})
        self.assertEqual(done.returncode, 3)
        self.assertIn(b"[pendiente] P5-P8 medicion de espacio", done.stdout)
        self.assertTrue(done.stdout.isascii())

    def test_claims_writes_the_table(self) -> None:
        self.assertEqual(_cli(["demo", str(self.root)]).returncode, 0)
        deck = self.root / "demo" / "research" / "notes" / f"demo-deck_{D.DEMO_AS_OF}.html"
        out = self.root / "frases.md"
        done = _cli(["claims", str(deck), str(out)])
        self.assertEqual(done.returncode, 0)
        self.assertTrue(done.stdout.isascii())
        row = "| 3 | texto | El crecimiento se modera, no se detiene. |  |"
        self.assertIn(row, out.read_text(encoding="utf-8"))
        out.write_text(out.read_text(encoding="utf-8").replace(
            row, row[:-3] + " serie de ventas |"), encoding="utf-8")
        again = _cli(["claims", str(deck), str(out)])       # no pisa evidencia ya escrita
        self.assertEqual(again.returncode, 1)
        self.assertIn(b"[x]", again.stdout)
        self.assertIn(b"audit --report", again.stdout)
        self.assertIn("| 3 | texto | El crecimiento se modera, no se detiene. | serie de ventas |",
                      out.read_text(encoding="utf-8"))
        other = self.root / "notas.md"
        other.write_text("Notas del analista.\n", encoding="utf-8")
        self.assertEqual(_cli(["claims", str(deck), str(other)]).returncode, 1)
        self.assertEqual(other.read_text(encoding="utf-8"), "Notas del analista.\n")

    def test_labels_writes_the_model_map_and_keeps_the_console_ascii(self) -> None:
        model = self.root / MODEL_REL
        wb = load_workbook(model)
        wb["Annual"]["B20"] = "Depreciaci\u00f3n"
        wb.save(model)
        out = self.root / "etiquetas.md"
        done = _cli(["labels", str(model), str(out)])
        self.assertEqual(done.returncode, 0, done.stdout.decode("ascii", "replace"))
        self.assertTrue(done.stdout.isascii())
        self.assertIn(b"[ok] 18 etiquetas en 3 hojas", done.stdout)
        text = out.read_text(encoding="utf-8")
        self.assertIn("| 20 | Depreciaci\u00f3n |  | 0 |  |  |", text)
        missing = _cli(["labels", str(self.root / "nope.xlsx"), str(out)])
        self.assertEqual(missing.returncode, 1)
        self.assertIn(b"[x]", missing.stdout)
        self.assertEqual(_cli(["labels", str(model), str(out)]).returncode, 0)   # su propia salida
        journal = self.root / "journal.md"
        journal.write_text("# Journal\n", encoding="utf-8")
        clobber = _cli(["labels", str(model), str(journal)])
        self.assertEqual(clobber.returncode, 1)
        self.assertEqual(journal.read_text(encoding="utf-8"), "# Journal\n")

    def test_facts_validates_the_register_without_building_a_deck(self) -> None:
        facts = self.root / FACTS_REL
        done = _cli(["facts", str(facts), "--root", str(self.root)])
        self.assertEqual(done.returncode, 0, done.stdout.decode("ascii", "replace"))
        self.assertIn(b"[ok] hecho share: cita encontrada", done.stdout)
        self.assertIn(b"[aviso] hecho plants: manual", done.stdout)
        # Sin --root: la cobertura es la carpeta que contiene research/notes/.
        self.assertEqual(_cli(["facts", str(facts)]).returncode, 0)
        facts.write_text(json.dumps({"facts": [
            _fact("share", 31.5, REPORT_REL, "31.5%"),
            _fact("lejos", 1.0, REPORT_REL, "no est\u00e1 en el reporte"),
            _fact("nodoc", 2.0, "filings/nope.md", "dos")]}), encoding="utf-8")
        broken = _cli(["facts", str(facts), "--root", str(self.root)])
        self.assertEqual(broken.returncode, 1)
        self.assertTrue(broken.stdout.isascii())
        self.assertIn(b"[ok] hecho share", broken.stdout)
        self.assertIn(b"[x] hecho lejos: la cita", broken.stdout)
        self.assertIn(b"[x] hecho nodoc: no existe el documento", broken.stdout)

    def test_facts_never_reads_green_with_nothing_validated(self) -> None:
        facts = self.root / FACTS_REL
        for payload in ({"hechos": [_fact("share", 31.5, REPORT_REL, "31.5%")]}, {"facts": []}):
            facts.write_text(json.dumps(payload), encoding="utf-8")
            done = _cli(["facts", str(facts), "--root", str(self.root)])
            with self.subTest(payload=list(payload)):
                self.assertNotEqual(done.returncode, 0, done.stdout.decode("ascii", "replace"))
                self.assertNotIn(b"VERDE", done.stdout)
        facts.write_text(json.dumps({"facts": [_fact("nan", float("nan"), REPORT_REL, "31.5%")]}),
                         encoding="utf-8")
        done = _cli(["facts", str(facts), "--root", str(self.root)])
        self.assertEqual(done.returncode, 1)
        self.assertIn(b"[x] hecho nan: value debe ser un numero", done.stdout)

    def test_facts_outside_research_notes_needs_the_root(self) -> None:
        loose = self.root / "hechos.json"
        loose.write_text((self.root / FACTS_REL).read_text(encoding="utf-8"), encoding="utf-8")
        done = _cli(["facts", str(loose)])
        self.assertEqual(done.returncode, 1)
        self.assertIn(b"--root", done.stdout)
        self.assertEqual(_cli(["facts", str(loose), "--root", str(self.root)]).returncode, 0)

    def test_facts_dumps_the_visible_text_of_a_source(self) -> None:
        filing = self.root / "filings" / "anual.htm"
        filing.write_text("<html><body><table><tr><td>Ventas</td><td><ix:x>2,192,931</ix:x>"
                          "</td><td>58%</td></tr></table>"
                          + "<p>Parrafo de relleno con acento: operaci\u00f3n.</p>" * 30
                          + "</body></html>", encoding="utf-8")
        facts = self.root / FACTS_REL
        facts.write_text(json.dumps({"facts": [_fact("x", 1.0, "filings/anual.htm", "nope")]}),
                         encoding="utf-8")
        out = self.root / "scratch" / "anual.txt"
        done = _cli(["facts", str(facts), "--root", str(self.root),
                     "--text", "filings/anual.htm", str(out)])
        # El volcado sale aunque el registro no valide: sirve para corregir la cita.
        self.assertEqual(done.returncode, 1)
        self.assertTrue(done.stdout.isascii())
        self.assertIn(b"[ok] texto visible", done.stdout)
        text = out.read_text(encoding="utf-8")
        self.assertIn("Ventas 2,192,931 58%", text)
        self.assertNotIn("<td>", text)
        self.assertIn("operaci\u00f3n", text)
        self.assertLessEqual(max(len(line) for line in text.splitlines()), 100)
        self.assertEqual(D._flatten(text), D._visible_text(filing))
        pdf = _cli(["facts", str(facts), "--root", str(self.root),
                    "--text", "filings/anual.pdf", str(out)])
        self.assertEqual(pdf.returncode, 1)
        self.assertIn(b"no es una fuente de texto", pdf.stdout)
        # Repetir el volcado sobre si mismo si; sobre otro archivo, nunca.
        again = _cli(["facts", str(facts), "--root", str(self.root),
                      "--text", "filings/anual.htm", str(out)])
        self.assertIn(b"[ok] texto visible", again.stdout)
        register = facts.read_text(encoding="utf-8")
        clobber = _cli(["facts", str(facts), "--root", str(self.root),
                        "--text", "filings/anual.htm", str(facts)])
        self.assertEqual(clobber.returncode, 1)
        self.assertNotIn(b"[ok] texto visible", clobber.stdout)
        self.assertEqual(facts.read_text(encoding="utf-8"), register)

    def test_measure_report_writes_the_space_section(self) -> None:
        self.assertEqual(_cli(["demo", str(self.root)]).returncode, 0)
        notes = self.root / "demo" / "research" / "notes"
        deck = notes / f"demo-deck_{D.DEMO_AS_OF}.html"
        report = notes / f"demo-deck-verify_{D.DEMO_AS_OF}.md"
        self.assertEqual(_cli(["audit", str(deck), str(self.root / "demo" / "model"
                                                       / f"demo-model_{D.DEMO_AS_OF}_v1.xlsx"),
                               "--facts", str(notes / f"demo-facts_{D.DEMO_AS_OF}.json"),
                               "--report", str(report)]).returncode, 0)
        done = _cli(["measure", str(deck), "--report", str(report)],
                    env={"RA_BROWSER": str(self.root / "nope.exe")})
        self.assertEqual(done.returncode, 3)
        self.assertTrue(done.stdout.isascii())
        text = report.read_text(encoding="utf-8")
        self.assertIn("## Espacio", text)
        self.assertIn("| [pendiente] | P5-P8 medicion de espacio |", text)
        self.assertNotIn("Deck medido", text)        # sin navegador no se midio nada
        self.assertIn("## Afirmaciones", text)

    def test_audit_design_is_found_in_the_coverage_and_must_be_the_deck_palette(self) -> None:
        (self.root / "brand").mkdir()
        (self.root / "brand" / "DESIGN.md").write_text("deck_text: #BBBBBB\n", encoding="utf-8")
        deck = self.deck(design="brand/DESIGN.md")
        deck.slide(title="Titulo", lede="", visual="<p>visual</p>", note="Nota.",
                   source="Fuente: x.")
        path = deck.build()
        args = ["audit", str(path), str(self.root / MODEL_REL), "--facts",
                str(self.root / FACTS_REL), "--design"]
        # Ruta relativa a la cobertura, como en Deck(design=...), desde otra carpeta.
        relative = _cli(args + ["brand/DESIGN.md"])
        self.assertEqual(relative.returncode, 0, relative.stdout.decode("ascii", "replace"))
        self.assertIn(b"[aviso] P9 contraste del texto", relative.stdout)
        # Un archivo que no existe NO da verde a una paleta que el deck no usa.
        wrong = _cli(args + ["brand/NOPE.md"])
        self.assertEqual(wrong.returncode, 1)
        self.assertIn(b"[x]", wrong.stdout)
        self.assertNotIn(b"[ok] P9", wrong.stdout)
        self.assertIn(b"otra paleta", wrong.stdout)

    def test_printing_a_number_survives_a_cp1252_console(self) -> None:
        code = ("import sys; sys.path.insert(0, 'tools'); import deck_builder as D; "
                "n = D.Num(240.0, 'Annual!H6') / 2; print(n); print(str(n)); print([n]); "
                "print([D.Num(1.0, chr(39) + chr(916) + ' WC' + chr(39) + '!C5')])")
        done = subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True,
                              env=dict(os.environ, PYTHONIOENCODING="cp1252:strict"), timeout=60)
        self.assertEqual(done.returncode, 0, done.stderr.decode("ascii", "replace"))
        self.assertIn(b"120.0 [(Annual!H6/2)]", done.stdout)

    def test_usage_errors_exit_two_and_missing_files_exit_one(self) -> None:
        self.assertEqual(_cli([]).returncode, 2)
        self.assertEqual(_cli(["audit", "solo-un-argumento"]).returncode, 2)
        self.assertEqual(_cli(["--help"]).returncode, 0)
        missing = _cli(["audit", str(self.root / "nope.html"), str(self.root / MODEL_REL)])
        self.assertEqual(missing.returncode, 1)
        self.assertIn(b"[x]", missing.stdout)


class HardeningTests(_Case):
    """Defectos que encontro la revision independiente del PR."""

    def _audit(self, html_path: Path) -> B.Finding:
        findings, _, _ = D.audit_deck(html_path, self.root / MODEL_REL, self.root / FACTS_REL)
        return findings[0]

    def test_ticks_end_on_equal_and_nearly_equal_values(self) -> None:
        for lo, hi in ((0.3, 0.30000000000000004), (0.0, 3e-11), (5.0, 5.0), (-2.0, -2.0),
                       (1e9, 1e9 + 1e-6)):
            with self.subTest(lo=lo, hi=hi):
                ticks = D._nice_ticks(lo, hi)
                self.assertTrue(2 <= len(ticks) <= 12, ticks)
                self.assertLessEqual(ticks[0], lo)
                self.assertGreaterEqual(ticks[-1], hi)

    def test_flat_series_can_be_charted(self) -> None:
        deck = self.deck()
        flat = D.Series("Margen", {y: D.Num(0.1 + 0.2 if y == 2022 else 0.3, f"Annual!C{y}")
                                   for y in (2022, 2023)}, {2022: True, 2023: True})
        self.assertIn("<polyline", deck.lines([("Margen", flat)]))
        self.assertIn("<rect", deck.bars(flat))

    def test_fact_on_the_left_of_an_operation_audits_green(self) -> None:
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        deck.raw(f"<p>Diferencia {deck.fact('share') - rev[2022]:.1f} y "
                 f"{deck.fact('share') - 1:.1f}.</p>", source="Fuente: x.")
        first = self._audit(deck.build())
        self.assertTrue(first.ok, first.detail)

    def test_number_in_the_source_line_audits_green(self) -> None:
        deck = self.deck()
        rev = deck.series("Annual", "Ingresos")
        deck.raw("<p>Texto.</p>", source=f"Fuente: modelo, ventas de {rev[2027]:,.0f}.")
        first = self._audit(deck.build())
        self.assertTrue(first.ok, first.detail)

    def test_cell_references_are_normalised_or_rejected(self) -> None:
        book = D.ModelBook(self.root / MODEL_REL)
        self.assertEqual(book.num("DCF", "c5").src, "DCF!C5")
        self.assertEqual(book.num("DCF", "$C$5").src, "DCF!C5")
        for bad in ("hola", "C", "5", "C0", "A99999999"):
            with self.subTest(coord=bad), self.assertRaisesRegex(D.DeckError, "celda invalida"):
                book.num("DCF", bad)
        with self.assertRaisesRegex(D.DeckError, "column debe ser"):
            book.cell("DCF", "WACC", column="5")
        for bad in ("C5", "C", "C5:D"):
            with self.subTest(ref=bad), self.assertRaisesRegex(D.DeckError, "rango invalido"):
                book.range("DCF", bad)

    def test_origin_without_a_reference_is_rejected_at_build_and_audit(self) -> None:
        deck = self.deck()
        deck.raw(f"<p>Vale {D.Num(0.38, '0.38'):.0%}.</p>", source="Fuente: x.")
        with self.assertRaisesRegex(D.DeckError, "sin referencia"):
            deck.build()
        good = self.deck()
        good.raw(f"<p>Vale {good.cell('DCF', 'WACC'):.0%}.</p>", source="Fuente: x.")
        path = good.build()
        path.write_text(path.read_text(encoding="utf-8").replace(
            'data-src="DCF!C5"', 'data-src="0.09"'), encoding="utf-8")
        first = self._audit(path)
        self.assertTrue(first.blocking)
        self.assertIn("sin referencia", first.detail)

    def test_build_checks_every_number_against_its_origin(self) -> None:
        fake = self.deck()
        fake.raw(f"<p>Vale {D.Num(0.38, 'Annual!C6'):.0%}.</p>", source="Fuente: x.")
        with self.assertRaisesRegex(D.DeckError, "no coinciden con su origen"):
            fake.build()
        edited = self.deck()
        rev = edited.series("Annual", "Ingresos")
        edited.raw("<p>Vale " + f"{rev[2027]:,.1f}".replace("240", "999") + ".</p>",
                   source="Fuente: x.")
        with self.assertRaisesRegex(D.DeckError, "muestra '999.0', el origen da '240.0'"):
            edited.build()
        self.assertFalse(edited.output_path.exists())

    def test_embedded_documents_are_not_slide_content(self) -> None:
        for content in ("<iframe srcdoc='38%'></iframe>", "<object data='x.svg'></object>",
                        "<embed src='x.svg'>"):
            with self.subTest(content=content[:8]), \
                    self.assertRaisesRegex(D.DeckError, "style ni script"):
                self.deck().raw(content, source="Fuente: x.")

    def test_audit_reports_hostile_origins_and_formats_instead_of_crashing(self) -> None:
        deck = self.deck()
        deck.raw(f"<p>Vale {deck.cell('DCF', 'WACC'):.0%}.</p>", source="Fuente: x.")
        path = deck.build()
        original = path.read_text(encoding="utf-8")
        cases = (('data-fmt=".0%"', 'data-fmt="9&gt;6,.0f"'),
                 ('data-fmt=".0%"', 'data-fmt="999999999999"'),
                 ('data-fmt=".0%"', 'data-fmt="zz"'),
                 ('data-src="DCF!C5"', 'data-src="DCF!C5**9**9**9"'),
                 ('data-src="DCF!C5"', 'data-src="Annual!A99999999"'),
                 ('data-src="DCF!C5"', 'data-src="DCF!C5' + "+1" * 3000 + '"'),
                 ('data-src="DCF!C5"', 'data-src="(0-DCF!C5)**0.5"'))
        for old, new in cases:
            with self.subTest(new=new[:40]):
                self.assertIn(old, original)
                path.write_text(original.replace(old, new), encoding="utf-8")
                self.assertTrue(self._audit(path).blocking)

    def test_paths_cannot_leave_the_coverage(self) -> None:
        cover = self.root / "cov"
        _coverage(cover)
        (self.root / "fuera.md").write_text("Dato: 42 unidades.", encoding="utf-8")
        (self.root / "fuera.svg").write_text("<svg/>", encoding="utf-8")
        facts = cover / FACTS_REL
        for doc in ("../fuera.md", str(self.root / "fuera.md")):
            with self.subTest(doc=doc[:12]):
                facts.write_text(json.dumps({"facts": [_fact("x", 42, doc, "42 unidades")]}),
                                 encoding="utf-8")
                with self.assertRaisesRegex(D.DeckError, "sale de la cobertura"):
                    D.load_facts(facts, cover)
        (cover / "brand").mkdir()
        design = cover / "brand" / "DESIGN.md"
        design.write_text("deck_logo: ../fuera.svg" + chr(10), encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "sale de la cobertura"):
            D.load_design(design, cover)

    def test_facts_file_must_be_an_object_with_a_list(self) -> None:
        path = self.root / FACTS_REL
        path.write_text("[]", encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "se espera un objeto"):
            D.load_facts(path, self.root)
        path.write_bytes("{\"facts\": \"a\u00f1o\"}".encode("cp1252"))
        with self.assertRaisesRegex(D.DeckError, "hechos ilegibles"):
            D.load_facts(path, self.root)

    def test_unknown_design_slot_is_an_error(self) -> None:
        path = self.root / "DESIGN.md"
        path.write_text("deck_backgroud: #FFFFFF" + chr(10), encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "slot desconocido"):
            D.load_design(path, self.root)

    def test_slide_content_cannot_carry_style_or_script(self) -> None:
        deck = self.deck()
        with self.assertRaisesRegex(D.DeckError, "style ni script"):
            deck.raw("<STYLE>p::after{content:'38%'}</STYLE><p>x</p>", source="Fuente: x.")
        with self.assertRaisesRegex(D.DeckError, "style ni script"):
            deck.slide(title="T", lede="", visual="<script>1</script>", note="Nota.",
                       source="Fuente: x.")

    def test_charts_reject_degenerate_input(self) -> None:
        deck = self.deck()
        low, high = deck.cell("DCF", "Bajo"), deck.cell("DCF", "Alto")
        grid = deck.range("Val Sens", "C6:E8")
        single = D.Series("Uno", {2022: low}, {2022: True})
        for call in (lambda: deck.lines([]), lambda: deck.football([]),
                     lambda: deck.football([("DCF", high, low, low)]),
                     lambda: deck.heatmap([], rows=[], cols=[]),
                     lambda: deck.heatmap(grid, rows=["a"], cols=["a", "b", "c"]),
                     lambda: single.growth(),
                     lambda: single.over(D.Series("Otro", {2023: high}, {2023: True}))):
            with self.assertRaises(D.DeckError):
                call()

    def test_negative_bar_label_stays_clear_of_the_period_label(self) -> None:
        deck = self.deck()
        change = deck.series("Annual", "Cambio")
        svg = D.render_markers(deck.bars(change, show=change.keys, fmt="+.1f", h=320))
        import re
        ys = [float(y) for y in re.findall(r'<text class="val" x="[0-9.]+" y="([0-9.]+)"', svg)]
        self.assertEqual(len(ys), len(YEARS))
        self.assertTrue(all(y < 320 - 30 for y in ys), ys)

    def test_unreadable_inputs_are_deck_errors(self) -> None:
        fake = self.root / "model" / "roto.xlsx"
        fake.write_text("no es un xlsx", encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "modelo ilegible"):
            D.ModelBook(fake)
        other = self.root / "model" / "modelo.txt"
        other.write_text("x", encoding="utf-8")
        with self.assertRaisesRegex(D.DeckError, "modelo ilegible"):
            D.ModelBook(other)
        deck = self.deck()
        deck.raw("<p>Texto.</p>", source="Fuente: x.")
        path = deck.build()
        path.write_bytes(path.read_text(encoding="utf-8").replace("Texto", "A\u00f1o")
                         .encode("cp1252"))
        with self.assertRaisesRegex(D.DeckError, "no esta en UTF-8"):
            D.audit_deck(path, self.root / MODEL_REL)
        with self.assertRaisesRegex(D.DeckError, "no esta en UTF-8"):
            D.measure_deck(path)

    def test_deck_errors_are_ascii_by_construction(self) -> None:
        with self.assertRaises(D.DeckError) as caught:
            D.ModelBook(self.root / MODEL_REL).series("Annual", "Utilidad de operaci\u00f3n")
        self.assertTrue(str(caught.exception).isascii())

    def test_clipped_fixed_slide_fails_p5(self) -> None:
        findings = D.measure_findings([{"n": 1, "kind": "cover", "clipped": True},
                                       _slide(2, 40.0, clipped=False)], 90)
        self.assertTrue(findings[0].blocking)
        self.assertIn("lamina 1: contenido cortado", findings[0].detail)


class CliHardeningTests(_Case):

    def _demo(self) -> tuple[Path, Path]:
        self.assertEqual(_cli(["demo", str(self.root)]).returncode, 0)
        return (self.root / "demo" / "research" / "notes" / f"demo-deck_{D.DEMO_AS_OF}.html",
                self.root / "demo" / "model" / f"demo-model_{D.DEMO_AS_OF}_v1.xlsx")

    def test_missing_design_file_audits_the_neutral_palette(self) -> None:
        # Deck(design=...) acepta que el archivo falte y usa el estilo neutro: audit
        # evalua esa misma paleta, y dice que lo hizo.
        deck, model = self._demo()
        facts = ["--facts", str(deck.parent / f"demo-facts_{D.DEMO_AS_OF}.json")]
        done = _cli(["audit", str(deck), str(model), *facts,
                     "--design", str(self.root / "NOPE.md")])
        self.assertEqual(done.returncode, 0, done.stdout.decode("ascii", "replace"))
        self.assertIn(b"-- Diseno (paleta neutra: no existe NOPE.md)", done.stdout)
        self.assertIn(b"[ok] P9 contraste del texto", done.stdout)
        self.assertIn(b"[ok] P10 series distinguibles con daltonismo", done.stdout)
        folder = _cli(["audit", str(deck), str(model), *facts, "--design", str(self.root)])
        self.assertEqual(folder.returncode, 1)      # una carpeta no es "sin archivo"
        self.assertIn(b"[x] --design no es un archivo", folder.stdout)

    def test_broken_inputs_never_end_in_a_traceback(self) -> None:
        deck, model = self._demo()
        corrupt = self.root / "roto.xlsx"
        corrupt.write_bytes(b"PK no es un zip")
        facts = self.root / "hechos.json"
        facts.write_text("[]", encoding="utf-8")
        ansi = self.root / "ansi.html"
        ansi.write_bytes("<p>a\u00f1o</p>".encode("cp1252"))
        for args in (["audit", str(deck), str(corrupt)],
                     ["audit", str(deck), str(model), "--facts", str(facts)],
                     ["audit", str(ansi), str(model)],
                     ["claims", str(ansi), str(self.root / "f.md")],
                     ["measure", str(ansi)]):
            with self.subTest(args=args[0] + " " + Path(args[1]).name):
                done = _cli(args)
                self.assertEqual(done.returncode, 1, done.stdout.decode("ascii", "replace"))
                self.assertIn(b"[x]", done.stdout)
                self.assertNotIn(b"Traceback", done.stderr)
                self.assertTrue(done.stdout.isascii())


@unittest.skipUnless(REAL_BROWSER, REAL_BROWSER_REASON)
class ClippedBrowserTests(_Case):
    """Portada con mas texto del que cabe: el navegador lo corta; P5 lo dice."""

    def test_overfull_cover_fails_p5(self) -> None:
        deck = self.deck()
        deck.cover("Titulo de portada demasiado largo " * 14)
        deck.raw("<p style='height:520px'>Texto.</p>", source="Fuente: x.")
        marks = {f.check.split()[0]: f for f in D.measure_deck(deck.build())}
        self.assertIn("lamina 1: contenido cortado", marks["P5"].detail)
        self.assertNotIn("lamina 2", marks["P5"].detail)

    def test_chart_labels_are_checked_for_overlap(self) -> None:
        deck = self.deck()
        keys = [f"{q}Q{y}" for y in range(2017, 2032) for q in (1, 2, 3, 4)]
        rev = deck.series("Annual", "Ingresos")
        quarterly = D.Series("Margen", {k: rev[2022] / rev[2027] for k in keys},
                             {k: i < 40 for i, k in enumerate(keys)})
        deck.slide(title="Lineas trimestrales", lede="", note="Nota.", source="Fuente: x.",
                   visual=deck.lines([("Margen", quarterly)], show=(keys[0], keys[39], keys[-1]),
                                     fmt=".0%", h=330))
        deck.slide(title="Etiquetas encimadas", lede="", note="Nota.", source="Fuente: x.",
                   visual="<svg viewBox='0 0 1100 330'><text x='100' y='60'>primera etiqueta"
                          "</text><text x='120' y='64'>segunda etiqueta</text></svg>")
        marks = {f.check.split()[0]: f for f in D.measure_deck(deck.build())}
        self.assertIn("lamina 2: encimadas 'primera etiqueta / segunda etiqueta'",
                      marks["P8"].detail)
        self.assertNotIn("lamina 1", marks["P8"].detail)


@unittest.skipUnless(REAL_BROWSER, REAL_BROWSER_REASON)
class DemoBrowserTests(_Case):
    """El deck demo, medido con el navegador de verdad."""

    def test_demo_measures_green_and_writes_one_shot_per_slide(self) -> None:
        deck, _, _ = D.build_demo(self.root)
        findings = D.measure_deck(deck, shots=self.root / "shots")
        self.assertTrue(all(f.ok for f in findings), [f.detail for f in findings])
        self.assertEqual(len(list((self.root / "shots").glob("slide-*.png"))), 7)


if __name__ == "__main__":
    unittest.main()
