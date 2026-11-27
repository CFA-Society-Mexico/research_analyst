"""Pruebas de tools/contract_check.py: checks K en los dos sentidos.

Cada check tiene una cobertura que lo pasa y un contrato roto que lo falla o
lo avisa. Stdlib unittest, sin red. Las pruebas de YAML piden PyYAML (lo
instala requirements.txt); el modo degradado se prueba sin el.

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

import contract_check as K  # noqa: E402
from xlsx_builder import Finding  # noqa: E402

PROFILE_OK = """\
ticker: "TICK"
name: "Emisora de prueba"
as_of: 2026-11-20
listing: "sec"
issuer_type: "non_financial"
adr: false
framework_governing: "us_gaap"
framework_confirmed_by_user: true
currency_reporting: "USD"
units: "USD millones"
fiscal_year_end: "09-27"
life_cycle_stage: "mature"
model_periodicity: "quarterly"
quarterly_history_years: 10
periodicity_confirmed_by_user: true
valuation_methods:
  dcf_fcff: true
  comps: true
industry_staleness_months: 6
materiality:
  critical_pct_target_price: 10
  capture_unlinked_pct: 0.5
mnpi_confirmed_public: true
"""

MACRO_OK = """\
updated_at: "2026-10-01"
owner: "analista"
staleness_warn_months: 3
rates:
  risk_free_10y: {value: 4.1, label: observado, source: "FRED DGS10"}
"""

COMP_OK = """\
ticker: "MSFT"
name: "Microsoft Corporation"
as_of: "2026-11-23 16:00 EST"
stale_days: 5
market_data:
  price: {value: 410.5, label: observado, source: "cierre 2026-11-23"}
  price_currency: "USD"
  shares_per_ads: {value: 1, label: observado, source: "accion ordinaria"}
  fx_to_reporting: {value: 1, source: "misma moneda", as_of: "2026-11-23"}
  shares_outstanding: {value: 7430000000, label: observado, source: "10-Q p.1"}
balance_items:
  total_debt: {value: 60000, label: observado, source: "10-Q p.5"}
  preferred_equity: {}
income_items:
  period: "ltm"
  revenue: {value: 245000, label: observado, source: "10-Q p.3"}
framework: "us_gaap"
"""

# Bloques por pais (countries:), con la forma de un workspace real: secciones
# propias del pais (labor), componentes que heredan el source del mapa que
# los contiene y una ruta con values[] citada en su mapa.
COUNTRIES_OK = """\
gate_status: "propuesta - SIN confirmar"
countries:
  mexico:
    as_of: "2026-09-07"
    rates:
      risk_free_10y: {value: 9.45, label: observado, source: "FRED IRLTLT01MXM156N"}
    equity_risk_premium:
      value: 6.74
      label: atribuido
      source: "Damodaran, tabla de paises"
      components:
        country_risk_premium: {value: 2.46, label: atribuido}
    labor:
      minimum_wage_daily_mxn: {value: 315.04, label: observado, source: "CONASAMI"}
    gdp:
      real_yoy_recent_path:
        label: observado
        source: "FRED NGDPRSAXDCMXQ"
        values:
          - {date: "2026-Q2", value: 1.93}
"""

CSV_HEADER = "line_item,period,value,source_doc,source_ref,tag\n"
QUARTERLY_OK = CSV_HEADER + (
    "is_ns_total,1Q2026,100.0,10-Q acc 1,p.3,observado\n"
    "is_ns_total,2Q2026,103.0,10-Q acc 2,p.3,observado\n"
    "is_ni,4Q2025,9.0,derivado,FY menos 3Q,derivado\n")
ANNUAL_OK = CSV_HEADER + "is_ns_total,FY2025,400.0,10-K acc 9,p.30,observado\n"


class _Coverage(unittest.TestCase):
    """Base: una cobertura valida en disco; cada prueba rompe UNA pieza."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.coverage = self.root / "TICK"
        self._write("TICK/issuer-profile.yaml", PROFILE_OK)
        self._write("macro/macro-view.yaml", MACRO_OK)
        self._write("TICK/comps/MSFT.yaml", COMP_OK)
        self._write("TICK/model/inputs/canonical_quarterly.csv", QUARTERLY_OK)
        self._write("TICK/model/inputs/canonical_annual.csv", ANNUAL_OK)

    def _write(self, relative: str, text: str) -> Path:
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    def _checks(self) -> dict[str, Finding]:
        return {f.check.split()[0]: f for f in K.audit_contracts(self.coverage)}

    def assertFails(self, finding: Finding, *fragments: str) -> None:
        self.assertTrue(finding.blocking, finding.detail)
        for fragment in fragments:
            self.assertIn(fragment, finding.detail)

    def assertWarns(self, finding: Finding, *fragments: str) -> None:
        self.assertTrue(finding.warning, finding.detail)
        self.assertFalse(finding.blocking, finding.detail)
        for fragment in fragments:
            self.assertIn(fragment, finding.detail)


class CaptureTests(_Coverage):

    def _k5(self) -> Finding:
        return K.check_capture(self.coverage / "model" / "inputs")

    def test_valid_capture_is_green(self) -> None:
        finding = self._k5()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("4 filas en 2 archivos", finding.detail)

    def test_wide_format_fails_on_its_columns(self) -> None:
        self._write("TICK/model/inputs/canonical_quarterly.csv",
                    "line_item,region,1Q2016,2Q2016\naccounts_payable,consolidado,1,2\n")
        self.assertFails(self._k5(), "canonical_quarterly.csv: faltan columnas period, value")

    def test_period_must_match_the_file(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    ANNUAL_OK + "is_ni,4Q2025,9.0,10-K,p.1,observado\n")
        self.assertFails(self._k5(), "canonical_annual.csv", "period fuera de formato", "linea 3")

    def test_non_numeric_value_fails(self) -> None:
        for bad in ("n/d", "", "nan"):
            with self.subTest(value=bad):
                self._write("TICK/model/inputs/canonical_annual.csv",
                            CSV_HEADER + f"is_ni,FY2025,{bad},10-K,p.1,observado\n")
                self.assertFails(self._k5(), "value no numerico")

    def test_duplicated_key_fails(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    ANNUAL_OK + "is_ns_total,FY2025,401.0,10-K/A,p.30,observado\n")
        self.assertFails(self._k5(), "llave (line_item, period) repetida")

    def test_tag_outside_the_contract_fails(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,10-K,p.1,estimado\n")
        self.assertFails(self._k5(), "tag fuera de (observado, derivado, guidance)")

    def test_valid_tag_with_extra_text_warns(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,10-K,p.1,observado (signo normalizado)\n")
        self.assertWarns(self._k5(), "tag con texto extra")

    def test_observed_figure_needs_its_document(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,,p.1,observado\n")
        self.assertFails(self._k5(), "observado sin source_doc")

    def test_blank_line_item_fails(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + ",FY2025,9.0,10-K,p.1,observado\n")
        self.assertFails(self._k5(), "line_item vacio")

    def test_coverage_without_capture_does_not_apply(self) -> None:
        for name in ("canonical_quarterly.csv", "canonical_annual.csv"):
            (self.coverage / "model" / "inputs" / name).unlink()
        finding = self._k5()
        self.assertTrue(finding.ok)
        self.assertIn("n/a", finding.detail)

    def test_many_bad_rows_are_counted_not_listed(self) -> None:
        rows = "".join(f"is_ni,FY{2000 + i},9.0,10-K,p.1,estimado\n" for i in range(40))
        self._write("TICK/model/inputs/canonical_annual.csv", CSV_HEADER + rows)
        finding = self._k5()
        self.assertFails(finding, "40 filas con tag fuera de")
        self.assertLess(len(finding.detail), 400)


# Mapa en linea sin cerrar: ningun parser de YAML lo acepta.
BROKEN_YAML = """\
updated_at: "2026-10-01"
rates:
  risk_free_10y: {value: 4.1, source: "FRED DGS10"
fx: 1
"""


@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class YamlParsingTests(_Coverage):

    def test_valid_coverage_passes_the_five_checks(self) -> None:
        found = self._checks()
        self.assertEqual(sorted(found), ["K1", "K2", "K3", "K4", "K5"])
        for check_id, finding in found.items():
            with self.subTest(check=check_id):
                self.assertTrue(finding.ok, finding.detail)

    def test_broken_yaml_fails_k1_with_file_and_line(self) -> None:
        self._write("macro/macro-view.yaml", BROKEN_YAML)
        found = self._checks()
        self.assertFails(found["K1"], "macro-view.yaml", "linea")
        # Lo que no parsea no se puede revisar por dentro: pendiente, no verde.
        self.assertTrue(found["K3"].pending)

    def test_yaml_that_is_not_a_mapping_fails_k1(self) -> None:
        self._write("TICK/comps/MSFT.yaml", "- solo\n- una lista\n")
        self.assertFails(self._checks()["K1"], "MSFT.yaml", "no es un mapa")

    def test_macro_flag_points_at_another_file(self) -> None:
        other = self._write("otra/macro-view.yaml", BROKEN_YAML)
        found = {f.check.split()[0]: f for f in K.audit_contracts(self.coverage, other)}
        self.assertFails(found["K1"], "macro-view.yaml")


@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class ProfileTests(_Coverage):

    def _k2(self, text: str | None = None) -> Finding:
        if text is not None:
            self._write("TICK/issuer-profile.yaml", text)
        return self._checks()["K2"]

    def test_required_field_left_blank_fails(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace('units: "USD millones"', 'units: ""')),
                         "units vacio")

    def test_required_field_missing_fails(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace('listing: "sec"\n', "")), "listing vacio")

    def test_value_outside_its_list_fails(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace('listing: "sec"', 'listing: "nyse"')),
                         "listing: 'nyse' no esta en (bmv, sec, dual, private)")

    def test_deprecated_periodicity_warns_with_its_equivalent(self) -> None:
        text = PROFILE_OK.replace('model_periodicity: "quarterly"',
                                  'model_periodicity: "annual_plus_quarterly"')
        self.assertWarns(self._k2(text), "annual_plus_quarterly", "quarterly")

    def test_unconfirmed_profile_fails(self) -> None:
        for flag in K.PROFILE_CONFIRMATIONS:
            with self.subTest(flag=flag):
                self.assertFails(self._k2(PROFILE_OK.replace(f"{flag}: true", f"{flag}: false")),
                                 f"{flag} no esta en true")

    def test_unknown_field_warns(self) -> None:
        self.assertWarns(self._k2(PROFILE_OK + 'listing_notes: "ADR nivel 3"\n'),
                         "campo desconocido: listing_notes")

    def test_dates_and_fiscal_year_end_have_a_format(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace("as_of: 2026-11-20", 'as_of: "noviembre"')),
                         "as_of sin fecha YYYY-MM-DD")

    def test_fiscal_year_end_that_is_not_a_calendar_date_warns(self) -> None:
        # Un anio fiscal de 52/53 semanas no cabe en MM-DD: aviso, no falla.
        self.assertWarns(self._k2(PROFILE_OK.replace('"09-27"', '"09-lastSat"')),
                         "fiscal_year_end no es MM-DD")

    def test_typed_fields_reject_the_wrong_type(self) -> None:
        cases = {
            "quarterly_history_years: 10": ("quarterly_history_years: -1", "quarterly_history_years"),
            "adr: false": ('adr: "no"', "adr no es booleano"),
            "  comps: true": ('  comps: "si"', "valuation_methods.comps"),
            "  capture_unlinked_pct: 0.5": ("  capture_unlinked_pct: -2", "materiality.capture_unlinked_pct"),
        }
        for good, (bad, fragment) in cases.items():
            with self.subTest(field=fragment):
                self.assertFails(self._k2(PROFILE_OK.replace(good, bad)), fragment)

    def test_optional_field_absent_is_not_a_finding(self) -> None:
        finding = self._k2(PROFILE_OK.replace("quarterly_history_years: 10\n", ""))
        self.assertTrue(finding.ok, finding.detail)

    def test_coverage_without_profile_fails(self) -> None:
        (self.coverage / "issuer-profile.yaml").unlink()
        self.assertFails(self._checks()["K2"], "no existe issuer-profile.yaml")

    def test_detail_stays_ascii_with_accented_values(self) -> None:
        finding = self._k2(PROFILE_OK.replace('listing: "sec"', 'listing: "bolsa de México"'))
        self.assertTrue(finding.detail.isascii(), finding.detail)


@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class MacroTests(_Coverage):

    def test_macro_view_without_date_fails(self) -> None:
        self._write("macro/macro-view.yaml", MACRO_OK.replace('"2026-10-01"', '""'))
        self.assertFails(self._checks()["K3"], "updated_at vacio o sin fecha")

    def test_missing_macro_view_warns(self) -> None:
        (self.root / "macro" / "macro-view.yaml").unlink()
        self.assertWarns(self._checks()["K3"], "sin macro-view.yaml")

    def _k3(self, text: str) -> Finding:
        self._write("macro/macro-view.yaml", text)
        return self._checks()["K3"]

    def test_unknown_field_and_bad_window(self) -> None:
        self._write("macro/macro-view.yaml", MACRO_OK + 'house_view: "x"\n')
        self.assertWarns(self._checks()["K3"], "campo desconocido: house_view")
        self._write("macro/macro-view.yaml",
                    MACRO_OK.replace("staleness_warn_months: 3", "staleness_warn_months: 0"))
        self.assertFails(self._checks()["K3"], "staleness_warn_months")


    # --- datos {value, source} y bloques por pais ---------------------------

    def test_country_blocks_are_green(self) -> None:
        finding = self._k3(MACRO_OK + COUNTRIES_OK)
        self.assertTrue(finding.ok, finding.detail)
        self.assertFalse(finding.warning, finding.detail)
        self.assertIn("1 pais", finding.detail)

    def test_country_datum_without_source_fails(self) -> None:
        text = MACRO_OK + COUNTRIES_OK.replace(', source: "CONASAMI"', "")
        self.assertFails(self._k3(text),
                         "countries.mexico.labor.minimum_wage_daily_mxn sin source")

    def test_component_without_any_source_up_the_tree_fails(self) -> None:
        text = MACRO_OK + COUNTRIES_OK.replace('      source: "Damodaran, tabla de paises"\n', "")
        self.assertFails(
            self._k3(text),
            "countries.mexico.equity_risk_premium sin source",
            "countries.mexico.equity_risk_premium.components.country_risk_premium sin source")

    def test_path_item_without_source_up_the_tree_fails(self) -> None:
        text = MACRO_OK + COUNTRIES_OK.replace('        source: "FRED NGDPRSAXDCMXQ"\n', "")
        self.assertFails(self._k3(text),
                         "countries.mexico.gdp.real_yoy_recent_path.values[0] sin source")

    def test_non_numeric_value_fails(self) -> None:
        text = MACRO_OK + COUNTRIES_OK.replace("value: 9.45,", 'value: "9.45%",')
        self.assertFails(self._k3(text),
                         "countries.mexico.rates.risk_free_10y sin value numerico")

    def test_null_value_with_source_warns_as_pending(self) -> None:
        text = MACRO_OK + COUNTRIES_OK + (
            "    fx:\n"
            "      usdmxn_path:\n"
            '        - {date: "2026-08-21", value: null, label: observado, source: "FRED"}\n')
        self.assertWarns(self._k3(text),
                         "countries.mexico.fx.usdmxn_path[0]: value pendiente (null)")

    def test_null_value_without_source_fails(self) -> None:
        text = MACRO_OK + COUNTRIES_OK + (
            "    fx:\n"
            "      usdmxn_spot: {value: null, label: observado}\n")
        self.assertFails(self._k3(text), "countries.mexico.fx.usdmxn_spot sin source")

    def test_top_level_data_follow_the_same_rule(self) -> None:
        self.assertFails(self._k3(MACRO_OK.replace(', source: "FRED DGS10"', "")),
                         "rates.risk_free_10y sin source")
        deck = ("commodity_decks:\n"
                '  - {commodity: cobre, unit: USD/lb, label: supuesto, source: "casa",\n'
                '     path: [{date: "2027", value: 4.2}]}\n')
        self.assertTrue(self._k3(MACRO_OK + deck).ok)
        self.assertFails(self._k3(MACRO_OK + deck.replace(', source: "casa"', "")),
                         "commodity_decks[0].path[0] sin source")

    def test_empty_template_placeholders_are_not_data(self) -> None:
        text = MACRO_OK + "equity_risk_premium: {}\ninflation:\n  mx_forecast_path: []\n"
        self.assertTrue(self._k3(text).ok)

    def test_country_block_without_capm_inputs_warns(self) -> None:
        text = MACRO_OK + COUNTRIES_OK.replace(
            '      risk_free_10y: {value: 9.45, label: observado, source: "FRED IRLTLT01MXM156N"}\n',
            '      govt_long_term: {value: 9.45, label: observado, source: "FRED"}\n')
        self.assertWarns(self._k3(text),
                         "countries.mexico sin rates.risk_free_10y (insumo del CAPM)")

    def test_country_as_of(self) -> None:
        self.assertFails(self._k3(MACRO_OK + COUNTRIES_OK.replace('"2026-09-07"', '"septiembre"')),
                         "countries.mexico.as_of sin fecha")
        self.assertWarns(self._k3(MACRO_OK + COUNTRIES_OK.replace('    as_of: "2026-09-07"\n', "")),
                         "countries.mexico sin as_of")

    def test_countries_must_be_maps(self) -> None:
        self.assertFails(self._k3(MACRO_OK + "countries: [mexico, taiwan]\n"),
                         "countries no es un mapa")
        self.assertFails(self._k3(MACRO_OK + 'countries:\n  taiwan: "pendiente"\n'),
                         "countries.taiwan no es un mapa")

    def test_empty_countries_is_valid(self) -> None:
        self.assertTrue(self._k3(MACRO_OK + "countries: {}\n").ok)

    def test_legacy_top_level_country_block_warns_with_migration_hint(self) -> None:
        legacy = ("taiwan:\n  rates:\n    risk_free_10y: "
                  '{value: 1.855, label: observado, source: "Investing.com"}\n')
        self.assertWarns(self._k3(MACRO_OK + legacy), "campo desconocido: taiwan", "countries:")

    def test_country_detail_stays_ascii(self) -> None:
        text = MACRO_OK + COUNTRIES_OK.replace("  mexico:", "  m\u00e9xico:").replace(
            ', source: "CONASAMI"', "")
        finding = self._k3(text)
        self.assertTrue(finding.blocking)
        self.assertTrue(finding.detail.isascii(), finding.detail)


@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class CompTests(_Coverage):

    def _k4(self, text: str | None = None, name: str = "MSFT.yaml") -> Finding:
        if text is not None:
            self._write(f"TICK/comps/{name}", text)
        return self._checks()["K4"]

    @staticmethod
    def _without(text: str, *keys: str) -> str:
        return "".join(line for line in text.splitlines(keepends=True)
                       if not line.strip().startswith(tuple(f"{k}:" for k in keys)))

    def test_missing_mandatory_market_fields_are_named(self) -> None:
        text = self._without(COMP_OK, "price_currency", "shares_per_ads", "fx_to_reporting")
        self.assertFails(self._k4(text), "MSFT.yaml",
                         "market_data sin price_currency, shares_per_ads, fx_to_reporting")

    def test_old_currency_key_is_reported_as_unknown(self) -> None:
        text = COMP_OK.replace('  price_currency: "USD"',
                               '  price_currency: "USD"\n  currency: "USD"')
        self.assertWarns(self._k4(text), "campo desconocido: market_data.currency")

    def test_price_currency_must_be_iso_4217(self) -> None:
        self.assertFails(self._k4(COMP_OK.replace('price_currency: "USD"', 'price_currency: "usd"')),
                         "price_currency no es ISO 4217")

    def test_captured_datum_needs_value_and_source(self) -> None:
        no_source = COMP_OK.replace('{value: 410.5, label: observado, source: "cierre 2026-11-23"}',
                                    "{value: 410.5, label: observado}")
        self.assertFails(self._k4(no_source), "market_data.price sin source")
        no_value = COMP_OK.replace("{value: 60000, label: observado", "{value: n/d, label: observado")
        self.assertFails(self._k4(no_value), "balance_items.total_debt sin value numerico")

    def test_fx_needs_its_as_of(self) -> None:
        text = COMP_OK.replace('{value: 1, source: "misma moneda", as_of: "2026-11-23"}',
                               '{value: 1, source: "misma moneda"}')
        self.assertFails(self._k4(text), "market_data.fx_to_reporting sin as_of")

    def test_snapshot_needs_ticker_and_dated_as_of(self) -> None:
        self.assertFails(self._k4(COMP_OK.replace('as_of: "2026-11-23 16:00 EST"', 'as_of: ""')),
                         "as_of sin fecha")

    def test_foreign_framework_warns(self) -> None:
        self.assertWarns(self._k4(COMP_OK.replace('framework: "us_gaap"', 'framework: "k_ifrs"')),
                         "framework", "k_ifrs")

    def test_only_the_latest_snapshot_of_a_comparable_is_checked(self) -> None:
        old = self._without(COMP_OK, "price_currency").replace("2026-11-23 16:00 EST", "2026-08-01")
        self._write("TICK/comps/MSFT_2026-08-01.yaml", old)   # historia: incompleto, no cuenta
        finding = self._k4()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("1 snapshots vigentes", finding.detail)

    def test_tie_on_as_of_goes_to_the_newer_dated_file(self) -> None:
        # Un snapshot de correccion conserva el as_of del precio. Con la fecha
        # empatada, vigente es el archivo con la fecha mas reciente en el nombre.
        broken = self._without(COMP_OK, "price_currency")
        self._write("TICK/comps/MSFT.yaml", broken)                # sin fecha: el mas viejo
        self._write("TICK/comps/MSFT_2026-11-25.yaml", broken)
        self._write("TICK/comps/MSFT_2026-11-26.yaml", COMP_OK)    # la correccion
        finding = self._k4()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("1 snapshots vigentes", finding.detail)

    def test_tie_on_as_of_does_not_hide_a_broken_correction(self) -> None:
        # El sentido contrario: si la correccion es la rota, K4 debe verla.
        self._write("TICK/comps/MSFT_2026-11-26.yaml", self._without(COMP_OK, "price_currency"))
        self.assertFails(self._k4(), "MSFT_2026-11-26.yaml", "market_data sin price_currency")

    def test_newer_as_of_beats_a_newer_file_name(self) -> None:
        # as_of manda; la fecha del nombre solo desempata.
        stale = self._without(COMP_OK, "price_currency").replace("2026-11-23 16:00 EST", "2026-11-01")
        self._write("TICK/comps/MSFT_2026-11-26.yaml", stale)
        finding = self._k4()                                       # MSFT.yaml (as_of 11-23) sigue vigente
        self.assertTrue(finding.ok, finding.detail)

    def test_coverage_without_comps_does_not_apply(self) -> None:
        (self.coverage / "comps" / "MSFT.yaml").unlink()
        finding = self._k4()
        self.assertTrue(finding.ok)
        self.assertIn("n/a", finding.detail)

    def test_text_field_inside_a_section_is_unknown_not_a_broken_datum(self) -> None:
        text = COMP_OK.replace("balance_items:\n", 'balance_items:\n  currency: "USD"\n')
        self.assertWarns(self._k4(text), "campo desconocido: balance_items.currency")

    def test_empty_optional_datum_is_not_a_finding(self) -> None:
        finding = self._k4()                                   # preferred_equity: {}
        self.assertTrue(finding.ok, finding.detail)


class WithoutPyYamlTests(_Coverage):

    def test_yaml_checks_are_pending_and_csv_still_runs(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,10-K,p.1,estimado\n")
        with mock.patch.object(K, "yaml", None):
            found = self._checks()
        for check_id in ("K1", "K2", "K3", "K4"):
            with self.subTest(check=check_id):
                self.assertTrue(found[check_id].pending)
                self.assertIn("pip install pyyaml", found[check_id].detail)
        self.assertTrue(found["K5"].blocking)      # el CSV se valida igual


class CliTests(_Coverage):

    def _run(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = K.main(["contract_check.py", *args])
        return code, out.getvalue()

    @unittest.skipIf(K.yaml is None, "PyYAML no instalado")
    def test_valid_coverage_exits_green(self) -> None:
        code, out = self._run(str(self.coverage))
        self.assertEqual(code, 0)
        self.assertIn("[ok] K5", out)
        self.assertIn("Veredicto: VERDE", out)
        self.assertTrue(out.isascii())

    @unittest.skipIf(K.yaml is None, "PyYAML no instalado")
    def test_warning_does_not_change_the_exit_code(self) -> None:
        self._write("TICK/issuer-profile.yaml", PROFILE_OK + 'listing_notes: "x"\n')
        code, out = self._run(str(self.coverage))
        self.assertEqual(code, 0)
        self.assertIn("[aviso] K2", out)
        self.assertIn("1 avisos", out)

    def test_broken_contract_exits_failure(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv", "line_item,region\nx,y\n")
        code, out = self._run(str(self.coverage))
        self.assertEqual(code, 1)
        self.assertIn("[x] K5", out)
        self.assertIn("Veredicto: FALLA", out)

    def test_without_pyyaml_exits_pending(self) -> None:
        with mock.patch.object(K, "yaml", None):
            code, out = self._run(str(self.coverage))
        self.assertEqual(code, 3)
        self.assertIn("[pendiente] K1", out)

    def test_missing_folder_is_reported_without_traceback(self) -> None:
        code, out = self._run(str(self.root / "NOEXISTE"))
        self.assertEqual(code, 2)
        self.assertIn("[x] no existe la carpeta de cobertura", out)


if __name__ == "__main__":
    unittest.main()
