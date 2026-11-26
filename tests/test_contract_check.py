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


if __name__ == "__main__":
    unittest.main()
