"""Pruebas de los bloques de valuacion del builder y del anio de transicion.

Issues #31 (terminal Gordon normalizado + g implicita con guarda), #39
(cuadricula de sensibilidad sin data table), #37 (market cap de comparables
con ADS y FX) y #41 (anio de transicion). Las pruebas de estructura corren sin
Excel; ``ExcelEvaluationTests`` evalua las formulas con Excel real y se salta
donde no hay Excel (CI).

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import model_check as M  # noqa: E402
import recalc_validate as R  # noqa: E402
import xlsx_builder as B  # noqa: E402


def _formulas(ws: object) -> list[str]:
    return [c.value for row in ws.iter_rows() for c in row  # type: ignore[attr-defined]
            if isinstance(c.value, str) and c.value.startswith("=")]


class ValuationBlockStructureTests(unittest.TestCase):

    def setUp(self) -> None:
        self.styler = B.ModelStyler()
        self.ws = self.styler.new_sheet("Annual")

    def test_gordon_terminal_uses_normalized_nopat_flow_not_fcff(self) -> None:
        refs = self.styler.gordon_terminal_block(
            self.ws, 40, nopat_ref="$T$8", g_ref="G_TERMINAL", wacc_ref="WACC",
            ronic_ref="RONIC")
        self.assertEqual(refs["flow"], "$C$41")
        self.assertEqual(refs["tv"], "$C$42")
        flow = self.ws["C41"].value
        self.assertIn("$T$8*(1+G_TERMINAL)*(1-G_TERMINAL/RONIC)", flow)
        self.assertIn("$C$41/(WACC-G_TERMINAL)", self.ws["C42"].value)
        self.assertIn("WACC<=G_TERMINAL", self.ws["C42"].value)      # guarda g >= WACC
        self.assertIn("RONIC", str(self.ws["B41"].value))

    def test_implied_growth_block_is_six_rows_with_the_spurious_root_guard(self) -> None:
        refs = self.styler.implied_growth_block(
            self.ws, 50, title="Reverse DCF", tv_ref="$C$48", nopat_ref="$T$8",
            wacc_ref="WACC", ronic_ref="RONIC")
        self.assertEqual(refs["g"], "$C$55")
        self.assertEqual(refs["next_row"], "56")
        self.assertEqual(self.ws["C50"].value, "=$T$8/RONIC")                    # A
        self.assertEqual(self.ws["C51"].value, "=$C$48+$T$8*(1-1/RONIC)")        # B
        self.assertEqual(self.ws["C52"].value, "=$C$48*WACC-$T$8")               # C
        g = self.ws["C55"].value
        self.assertIn(B.NO_GROWTH_BELOW_WACC, g)
        self.assertIn(B.NO_REAL_SOLUTION, g)
        self.assertIn(">=WACC", g)

    def test_valuation_formulas_carry_no_hardcoded_literals(self) -> None:
        # S4: solo 0 y 1 estructurales; el 2 y el 4 de la cuadratica no aparecen.
        self.styler.gordon_terminal_block(self.ws, 40, nopat_ref="$T$8", g_ref="G_TERMINAL",
                                          wacc_ref="WACC", ronic_ref="RONIC")
        self.styler.implied_growth_block(self.ws, 50, title="Exit", tv_ref="$C$48",
                                         nopat_ref="$T$8", wacc_ref="WACC", ronic_ref="RONIC")
        self.styler.sensitivity_grid(
            self.ws, 60, 2, row_base_ref="WACC", col_base_ref="G_TERMINAL",
            row_step_ref="SENS_WACC_STEP", col_step_ref="SENS_G_STEP",
            cell_formula=lambda w, g: B.dcf_per_share_formula(
                wacc=w, g=g, fcff_range="$D$5:$H$5", exponent_range="$D$6:$H$6",
                nopat_n="$H$3", ronic="RONIC", tv_exponent="$C$7",
                net_debt="$C$8", shares="$C$9"))
        self.styler.equity_terminal_block(self.ws, 70, ni_ref="$T$20", g_ref="G_TERMINAL",
                                          ke_ref="KE", roe_ref="ROE_TERMINAL")
        self.styler.equity_value_block(
            self.ws, 80, title="DDM", flow_range="$D$40:$H$40",
            exponent_range="$D$30:$H$30", ke_ref="KE", tv_ref="$C$72",
            tv_exponent_ref="$C$31", shares_ref="$C$32", require_positive_flows=True)
        for formula in _formulas(self.ws):
            with self.subTest(formula=formula):
                self.assertEqual(M.hardcoded_literals(formula), [])
                self.assertEqual(B.lint_formula(formula), [])

    def test_sensitivity_grid_headers_step_from_the_model_base_without_data_table(self) -> None:
        center = self.styler.sensitivity_grid(
            self.ws, 10, 2, row_base_ref="WACC", col_base_ref="G_TERMINAL",
            row_step_ref="SENS_WACC_STEP", col_step_ref="SENS_G_STEP",
            cell_formula=lambda w, g: f"={w}-{g}", n_side=2)
        self.assertEqual(center, "$E$13")
        self.assertEqual(self.ws["E10"].value, "=G_TERMINAL")            # columna central
        self.assertEqual(self.ws["F10"].value, "=E10+SENS_G_STEP")
        self.assertEqual(self.ws["D10"].value, "=E10-SENS_G_STEP")
        self.assertEqual(self.ws["C10"].value, "=D10-SENS_G_STEP")
        self.assertEqual(self.ws["B13"].value, "=WACC")                  # fila central
        self.assertEqual(self.ws["B14"].value, "=B13+SENS_WACC_STEP")
        self.assertEqual(self.ws["E13"].value, "=$B13-E$10")
        self.assertEqual(self.ws["G15"].value, "=$B15-G$10")
        # Cuadricula completa: 5 x 5 celdas con formula, ninguna vacia.
        empties = [f"{c}{r}" for r in range(11, 16) for c in "CDEFG"
                   if self.ws[f"{c}{r}"].value is None]
        self.assertEqual(empties, [])
        self.assertFalse(any("TABLE(" in f for f in _formulas(self.ws)))

    def test_comps_market_cap_converts_ads_price_and_currency(self) -> None:
        formula = B.comps_market_cap_formula(price="C5", shares_per_ads="C6",
                                             shares="C7", fx_to_reporting="C8")
        self.assertEqual(formula, "=C5/C6*C7*C8")
        flag = B.multiple_range_flag_formula("H5", "COMP_MULT_MIN", "COMP_MULT_MAX")
        self.assertIn(B.REVIEW_FLAG, flag)
        self.assertEqual(B.lint_formula(flag), [])
        self.assertEqual(M.hardcoded_literals(flag), [])

    def test_gordon_terminal_cells_are_frozen_across_the_refactor(self) -> None:
        self.styler.gordon_terminal_block(self.ws, 40, nopat_ref="$T$8", g_ref="G_TERMINAL",
                                          wacc_ref="WACC", ronic_ref="RONIC")
        self.assertEqual(self.ws["B40"].value, "Tasa de reinversion terminal (g / RONIC)")
        self.assertEqual(self.ws["C40"].value, '=IF(RONIC<=0,"n/s",G_TERMINAL/RONIC)')
        self.assertEqual(self.ws["B41"].value,
                         "Flujo terminal normalizado (NOPAT_n x (1+g) x (1 - g/RONIC))")
        self.assertEqual(self.ws["C41"].value,
                         '=IF(RONIC<=0,"n/s",$T$8*(1+G_TERMINAL)*(1-G_TERMINAL/RONIC))')
        self.assertEqual(self.ws["B42"].value, "TV Gordon (flujo terminal / (WACC - g))")
        self.assertEqual(self.ws["C42"].value,
                         '=IF(OR(WACC<=G_TERMINAL,NOT(ISNUMBER($C$41))),"n/s",'
                         '$C$41/(WACC-G_TERMINAL))')

    def test_equity_terminal_normalizes_net_income_with_terminal_roe(self) -> None:
        refs = self.styler.equity_terminal_block(
            self.ws, 70, ni_ref="$T$20", g_ref="G_TERMINAL", ke_ref="KE",
            roe_ref="ROE_TERMINAL")
        self.assertEqual(refs["reinvestment"], "$C$70")
        self.assertEqual(refs["tv"], "$C$72")
        self.assertEqual(refs["next_row"], "73")
        self.assertEqual(refs["flow"], "$C$71")
        self.assertIn("NI_n", self.ws["B71"].value)
        self.assertEqual(self.ws["C70"].value,
                         '=IF(ROE_TERMINAL<=0,"n/s",G_TERMINAL/ROE_TERMINAL)')
        self.assertEqual(self.ws["C71"].value,
                         '=IF(ROE_TERMINAL<=0,"n/s",'
                         '$T$20*(1+G_TERMINAL)*(1-G_TERMINAL/ROE_TERMINAL))')
        self.assertEqual(self.ws["C72"].value,
                         '=IF(OR(KE<=G_TERMINAL,NOT(ISNUMBER($C$71))),"n/s",'
                         '$C$71/(KE-G_TERMINAL))')
        self.assertIn("ROE terminal", self.ws["B70"].value)
        self.assertIn("ke - g", self.ws["B72"].value)

    def _value_block(self, **kwargs: object) -> dict[str, str]:
        return self.styler.equity_value_block(
            self.ws, 80, title="DDM", flow_range="$D$40:$H$40",
            exponent_range="$D$30:$H$30", ke_ref="KE", tv_ref="$C$72",
            tv_exponent_ref="$C$31", shares_ref="$C$32", **kwargs)

    def test_equity_value_block_discounts_equity_flows_without_net_debt(self) -> None:
        refs = self._value_block()
        self.assertEqual(self.ws["C80"].value, "=SUMPRODUCT($D$40:$H$40/(1+KE)^$D$30:$H$30)")
        self.assertEqual(self.ws["C81"].value, '=IF(ISNUMBER($C$72),$C$72/(1+KE)^$C$31,"n/s")')
        self.assertEqual(self.ws["C82"].value, '=IF(ISNUMBER($C$81),$C$80+$C$81,"n/s")')
        self.assertEqual(self.ws["C83"].value,
                         '=IF(OR(NOT(ISNUMBER($C$82)),$C$32<=0),"n/s",$C$82/$C$32)')
        self.assertEqual(self.ws["C84"].value,
                         '=IF(OR(NOT(ISNUMBER($C$82)),$C$82<=0),"n/s",$C$81/$C$82)')
        self.assertEqual(refs["per_share"], "$C$83")
        self.assertEqual(refs["next_row"], "85")
        self.assertNotIn("flag", refs)
        self.assertIsNone(self.ws["C85"].value)
        self.assertTrue(str(self.ws["B83"].value).startswith("DDM: "))

    def test_ddm_flags_a_method_activated_without_dividends(self) -> None:
        refs = self._value_block(require_positive_flows=True)
        self.assertEqual(refs["flag"], "$C$85")
        self.assertEqual(refs["next_row"], "86")
        self.assertEqual(self.ws["C85"].value, '=IF(SUM($D$40:$H$40)<=0,"REVISAR","OK")')
        self.assertEqual(self.ws["B85"].value, "DDM: flujos explicitos positivos")
        self.assertEqual(self.ws["C85"].font.color.rgb, B.Color.WARN_RED.value)

    def test_equity_flow_rows_subtract_payments_stored_as_positive_magnitudes(self) -> None:
        refs = self.styler.equity_flow_rows(
            self.ws, 40, 4, 3, dividends_row=20, cfo_row=21, capex_row=22,
            issuance_rows=[23], repayment_rows=[24, 25])
        self.assertEqual(refs, {"dividends_row": "40", "fcfe_row": "41", "next_row": "42"})
        self.assertEqual(self.ws["D40"].value, "=D20")             # dividendos sin voltear
        self.assertEqual(self.ws["F40"].value, "=F20")
        self.assertEqual(self.ws["D41"].value, "=D21-D22+D23-D24-D25")
        self.assertEqual(self.ws["F41"].value, "=F21-F22+F23-F24-F25")
        self.assertIsNone(self.ws["G41"].value)                    # solo n_cols columnas
        self.assertEqual(self.ws["B40"].value, "Dividendos al accionista")

    def test_equity_flow_rows_without_debt_lines(self) -> None:
        self.styler.equity_flow_rows(self.ws, 40, 4, 1, dividends_row=20, cfo_row=21,
                                     capex_row=22)
        self.assertEqual(self.ws["D41"].value, "=D21-D22")

    def test_method_divergence_is_spread_over_median_and_only_warns(self) -> None:
        formula = B.method_divergence_formula(["$C$90", "$C$83", "$C$95"],
                                              "METHOD_DIVERGENCE_MAX")
        args = "$C$90,$C$83,$C$95"
        self.assertEqual(
            formula,
            f'=IF(COUNT({args})<=1,"n/a",IF(MEDIAN({args})<=0,"n/a",'
            f'IF((MAX({args})-MIN({args}))/MEDIAN({args})>METHOD_DIVERGENCE_MAX,'
            f'"REVISAR","OK")))')
        self.assertNotIn('"Error"', formula)                       # aviso, nunca bloqueo
        self.assertEqual(B.lint_formula(formula), [])
        self.assertEqual(M.hardcoded_literals(formula), [])

    def test_method_divergence_needs_at_least_one_value(self) -> None:
        with self.assertRaises(ValueError):
            B.method_divergence_formula([], "METHOD_DIVERGENCE_MAX")

    def test_method_consistency_block_lists_the_usual_sources_of_divergence(self) -> None:
        refs = self.styler.method_consistency_block(
            self.ws, 100, g_ref="G_TERMINAL", roe_last_ref="$H$60",
            dividends_last_ref="$H$40", ni_last_ref="$H$20", ke_equity_ref="KE",
            ke_wacc_ref="$C$12", reinvestment_ref="$C$70", net_debt_first_ref="$D$50",
            ebitda_first_ref="$D$15", net_debt_last_ref="$H$50", ebitda_last_ref="$H$15")
        self.assertEqual(refs["next_row"], "107")
        self.assertEqual(self.ws["C100"].value, '=IF($H$20<=0,"n/s",$H$60*(1-$H$40/$H$20))')
        self.assertEqual(self.ws["C101"].value, '=IF(ISNUMBER($C$100),G_TERMINAL-$C$100,"n/s")')
        self.assertEqual(self.ws["C102"].value, "=KE-$C$12")
        self.assertEqual(self.ws["C103"].value, '=IF(ISNUMBER($C$70),1-$C$70,"n/s")')
        self.assertEqual(self.ws["C104"].value, '=IF($H$20<=0,"n/s",$H$40/$H$20)')
        self.assertEqual(self.ws["C105"].value, '=IF($D$15<=0,"n/s",$D$50/$D$15)')
        self.assertEqual(self.ws["C106"].value, '=IF($H$15<=0,"n/s",$H$50/$H$15)')
        for formula in _formulas(self.ws):
            with self.subTest(formula=formula):
                self.assertEqual(M.hardcoded_literals(formula), [])
                self.assertEqual(B.lint_formula(formula), [])


class TransitionYearTests(unittest.TestCase):
    """Issue #41: formula propia en la primera columna del forecast, declarada."""

    def _model(self, stamp: bool, transition: bool = True) -> B.ModelStyler:
        styler = B.ModelStyler()
        styler.set_periodicity("annual")
        if stamp:
            styler.set_transition_year(2026)
        ws = styler.new_sheet("Model")
        styler.period_header(ws, 3, 3, B.PeriodHeader(2022, 2029, 2025))
        cols = "CDEFGHIJ"
        hist = [100, 110, 121, 133]
        forecast = [f"={cols[i - 1]}6*(1+{cols[i]}5)" for i in range(4, 8)]
        kwargs = {"transition_first": "=G8+G9"} if transition else {}
        styler.series_row(ws, 6, "Ventas", 3, hist, forecast,
                          hist_role=B.CellRole.OBSERVED,
                          forecast_role=B.CellRole.FORMULA, **kwargs)
        return styler

    def test_transition_first_requires_the_declared_year(self) -> None:
        with self.assertRaises(ValueError) as caught:
            self._model(stamp=False)
        self.assertIn("set_transition_year", str(caught.exception))

    def test_transition_first_replaces_only_the_first_forecast_cell(self) -> None:
        ws = self._model(stamp=True).wb["Model"]
        self.assertEqual(ws["G6"].value, "=G8+G9")
        self.assertEqual(ws["H6"].value, "=G6*(1+H5)")
        self.assertEqual(ws["J6"].value, "=I6*(1+J5)")
        self.assertIsNone(ws["K6"].value)                      # no corre la serie una columna

    def test_none_is_a_valid_transition_value(self) -> None:
        styler = B.ModelStyler()
        styler.set_transition_year(2026)
        ws = styler.new_sheet("Model")
        styler.series_row(ws, 6, "x", 3, [1], [2, 3], transition_first=None)
        self.assertIsNone(ws["D6"].value)
        self.assertEqual(ws["E6"].value, 3)

    def test_s5_treats_the_declared_transition_column_as_its_own_segment(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "m.xlsx")
            self._model(stamp=True).save(path)
            s5 = M._check_s5(M.load_book(path))
        self.assertTrue(s5.ok, s5.detail)
        self.assertIn("2026", s5.detail)

    def test_s5_still_fails_an_undeclared_odd_first_forecast_formula(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "m.xlsx")
            styler = self._model(stamp=False, transition=False)
            styler.wb["Model"]["G6"] = "=G8+G9"               # formula distinta sin sello
            styler.save(path)
            s5 = M._check_s5(M.load_book(path))
        self.assertFalse(s5.ok)
        self.assertIn("Model!fila 6", s5.detail)

    def test_transition_stamp_keeps_the_builder_stamp_and_periodicity(self) -> None:
        styler = self._model(stamp=True)
        keywords = styler.wb.properties.keywords
        self.assertIn(B.BUILDER_STAMP_KEY, keywords)
        self.assertIn("periodicity=annual", keywords)
        self.assertEqual(B.transition_year_of(keywords), 2026)
        self.assertIsNone(B.transition_year_of("periodicity=annual"))


@unittest.skipUnless(R.find_engine("excel") is R.Engine.EXCEL, "requiere Excel (COM)")
class ExcelEvaluationTests(unittest.TestCase):
    """Las formulas se evaluan con Excel: la matematica, no solo el texto."""

    NOPAT, WACC, RONIC, G = 100.0, 0.09, 0.15, 0.03

    @classmethod
    def setUpClass(cls) -> None:
        cls.tmp = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.tmp.cleanup)
        styler = B.ModelStyler()
        ws = styler.new_sheet("Annual")
        for coord, value in (("C5", cls.WACC), ("C6", cls.G), ("C7", cls.RONIC),
                             ("C8", cls.NOPAT), ("C9", 0.01), ("C10", 0.005),
                             ("C11", 250.0), ("C12", 40.0), ("C13", 5)):
            ws[coord] = value
        for name, coord in (("WACC", "C5"), ("G_TERMINAL", "C6"), ("RONIC", "C7"),
                            ("SENS_WACC_STEP", "C9"), ("SENS_G_STEP", "C10")):
            styler.define_constant(name, "Annual", coord)
        for i, (fcff, exponent) in enumerate(zip((60, 66, 72, 78, 84), (1, 2, 3, 4, 5))):
            ws.cell(row=15, column=4 + i, value=fcff)
            ws.cell(row=16, column=4 + i, value=exponent)
            col = B.get_column_letter(4 + i)
            styler.set_cell(ws, f"{col}17", f"={col}15/(1+WACC)^{col}16", B.CellRole.FORMULA)
        # (1) terminal normalizado y su g implicita (debe devolver G)
        tv = styler.gordon_terminal_block(ws, 20, nopat_ref="$C$8", g_ref="G_TERMINAL",
                                          wacc_ref="WACC", ronic_ref="RONIC")
        cls.refs = {"flow": "C21", "tv": "C22"}
        cls.refs["g_back"] = styler.implied_growth_block(
            ws, 24, title="cruce", tv_ref=tv["tv"], nopat_ref="$C$8",
            wacc_ref="WACC", ronic_ref="RONIC")["g"].replace("$", "")
        # (2) RONIC = WACC: raiz valida (TV que implica g = 3%) y raiz espuria
        ws["C31"] = cls.NOPAT * 1.03 / cls.WACC        # g verdadera 3% < WACC
        ws["C32"] = cls.NOPAT * 1.10 / cls.WACC        # g verdadera 10% >= WACC
        cls.refs["g_valid"] = styler.implied_growth_block(
            ws, 34, title="R=W valida", tv_ref="$C$31", nopat_ref="$C$8",
            wacc_ref="WACC", ronic_ref="WACC")["g"].replace("$", "")
        cls.refs["g_spurious"] = styler.implied_growth_block(
            ws, 41, title="R=W espuria", tv_ref="$C$32", nopat_ref="$C$8",
            wacc_ref="WACC", ronic_ref="WACC")["g"].replace("$", "")
        # (3) valor por accion del modelo, linea por linea, y la cuadricula
        styler.set_cell(ws, "C48", "=(SUM(D17:H17)+C22/(1+WACC)^C13-C11)/C12",
                        B.CellRole.FORMULA, B.NumFmt.DEC2)
        center = styler.sensitivity_grid(
            ws, 50, 2, row_base_ref="WACC", col_base_ref="G_TERMINAL",
            row_step_ref="SENS_WACC_STEP", col_step_ref="SENS_G_STEP",
            cell_formula=lambda w, g: B.dcf_per_share_formula(
                wacc=w, g=g, fcff_range="$D$15:$H$15", exponent_range="$D$16:$H$16",
                nopat_n="$C$8", ronic="RONIC", tv_exponent="$C$13",
                net_debt="$C$11", shares="$C$12"))
        cls.refs["center"] = center.replace("$", "")
        styler.set_cell(ws, "C58", B.center_check_formula(center, "$C$48"),
                        B.CellRole.WARN, B.NumFmt.GENERAL)
        # (4) comparables: ADR en USD, 5 acciones por ADS, FX 32 a moneda de reporte
        for coord, value in (("C60", 8.0), ("C61", 5.0), ("C62", 12_500.0), ("C63", 32.0),
                             ("C65", 1.0), ("C66", 2.0), ("C67", 60.0)):
            ws[coord] = value
        styler.set_cell(ws, "C64", B.comps_market_cap_formula(
            price="C60", shares_per_ads="C61", shares="C62", fx_to_reporting="C63"),
            B.CellRole.FORMULA)
        styler.set_cell(ws, "C68", B.multiple_range_flag_formula("C65", "C66", "C67"),
                        B.CellRole.WARN, B.NumFmt.GENERAL)
        source = Path(cls.tmp.name) / "valuacion.xlsx"
        styler.save(str(source))
        cls.source = source
        recalculated = R.recalculate(source, Path(cls.tmp.name) / "wd", R.Engine.EXCEL)
        cls.values = load_workbook(str(recalculated), data_only=True)["Annual"]
        cls.recalculated = recalculated

    def _v(self, key: str) -> object:
        return self.values[self.refs.get(key, key)].value

    def test_excel_keeps_every_formula_the_builder_wrote(self) -> None:
        self.assertEqual(R.formula_cells(self.source) - R.formula_cells(self.recalculated), set())
        self.assertEqual(R.error_cells(self.recalculated), [])

    def test_normalized_terminal_flow_and_value(self) -> None:
        flow = self.NOPAT * (1 + self.G) * (1 - self.G / self.RONIC)
        self.assertAlmostEqual(self._v("flow"), flow, places=6)
        self.assertAlmostEqual(self._v("tv"), flow / (self.WACC - self.G), places=6)

    def test_implied_growth_recovers_the_g_that_built_the_terminal(self) -> None:
        self.assertAlmostEqual(self._v("g_back"), self.G, places=9)

    def test_ronic_equal_wacc_returns_the_valid_root(self) -> None:
        self.assertAlmostEqual(self._v("g_valid"), 0.03, places=9)

    def test_ronic_equal_wacc_spurious_root_is_text_never_a_number(self) -> None:
        self.assertEqual(self._v("g_spurious"), B.NO_GROWTH_BELOW_WACC)

    def test_sensitivity_center_reproduces_the_model_value(self) -> None:
        self.assertAlmostEqual(self._v("center"), self._v("C48"), places=6)
        self.assertEqual(self._v("C58"), "OK")

    def test_sensitivity_grid_is_fully_evaluated_and_monotonic(self) -> None:
        grid = [[self.values.cell(row=51 + i, column=3 + j).value for j in range(5)]
                for i in range(5)]
        self.assertTrue(all(isinstance(v, (int, float)) for row in grid for v in row), grid)
        for row in grid:                                   # mas g => mas valor (RONIC > WACC)
            self.assertEqual(row, sorted(row))
        for j in range(5):                                 # mas WACC => menos valor
            column = [grid[i][j] for i in range(5)]
            self.assertEqual(column, sorted(column, reverse=True))

    def test_comps_market_cap_in_reporting_currency(self) -> None:
        # 8 USD por ADS / 5 acciones por ADS x 12,500 acciones x 32 = 640,000
        self.assertAlmostEqual(self._v("C64"), 640_000.0, places=6)
        self.assertEqual(self._v("C68"), B.REVIEW_FLAG)     # 1.0x fuera de 2x-60x


if __name__ == "__main__":
    unittest.main()
