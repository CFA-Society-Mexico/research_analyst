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
            (1, "Emisora X"),
            (2, "Ventas crecen +14.3% en 2027"),
            (2, "Menos que antes."),
            (2, "Sigue creciendo."),
            (2, "Participa con 31.5% del mercado."),
        ])
        self.assertIn("| 2 | Sigue creciendo. |  |", D.claims_table(doc))

    def test_report_holds_the_deterministic_sections(self) -> None:
        path = self._built()
        findings, doc, facts = D.audit_deck(path, self.root / MODEL_REL, self.root / FACTS_REL)
        report = self.root / "research" / "notes" / f"x-deck-verify_{AS_OF}.md"
        D.write_report(report, path, [("Cifras", findings)], doc, facts)
        text = report.read_text(encoding="utf-8")
        for expected in ("| [ok] | P1 cifras iguales a su origen |",
                         "| 2 | +14.3% | `((Annual!H6/Annual!G6)-1)` | `+.1%` |",
                         "| share | 31.5 % | observado |", "## Afirmaciones",
                         "| 2 | Participa con 31.5% del mercado. |  |"):
            self.assertIn(expected, text)


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
        self.assertEqual(_cli(["demo", str(self.root)]).returncode, 0)
        notes = self.root / "demo" / "research" / "notes"
        design = self.root / "DESIGN.md"
        design.write_text("deck_text: #BBBBBB" + chr(10), encoding="utf-8")
        done = _cli(["audit", str(notes / f"demo-deck_{D.DEMO_AS_OF}.html"),
                     str(self.root / "demo" / "model" / f"demo-model_{D.DEMO_AS_OF}_v1.xlsx"),
                     "--facts", str(notes / f"demo-facts_{D.DEMO_AS_OF}.json"),
                     "--design", str(design)])
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
        self.assertIn("| 3 | El crecimiento se modera, no se detiene. |  |",
                      out.read_text(encoding="utf-8"))

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

    def test_missing_design_file_is_an_error_not_a_green_palette(self) -> None:
        deck, model = self._demo()
        done = _cli(["audit", str(deck), str(model), "--design", str(self.root / "NOPE.md")])
        self.assertEqual(done.returncode, 1)
        self.assertIn(b"[x]", done.stdout)
        self.assertNotIn(b"P9", done.stdout)

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
