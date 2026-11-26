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
                               ([dict(ok, key="a b")], "la clave solo admite")):
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
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(D.loose_numbers(text, years), expected)

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
