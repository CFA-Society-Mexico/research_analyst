"""Pruebas del diagnostico de cobertura de tools/xbrl_fetch.py.

Nace de la prueba en vivo con TSM (2026-11-25): la captura IFRS salio sin D&A,
sin impuestos pagados y sin FY2025, y el tool no decia que conceptos reporta
la emisora que el mapa no cubre. Stdlib unittest, sin red.

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
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import xbrl_fetch as X  # noqa: E402

ACCN_2024 = "0001-25-000020"
ACCN_2025 = "0001-26-000031"


def _annual(val: float, year: int, accn: str, filed: str) -> dict:
    return {"start": f"{year}-01-01", "end": f"{year}-12-31", "val": val,
            "accn": accn, "filed": filed, "form": "20-F"}


def _node(unit: str, *items: dict, label: str = "") -> dict:
    node: dict = {"units": {unit: list(items)}}
    if label:
        node["label"] = label
    return node


def _payload(extra_ifrs: dict | None = None, other: dict | None = None) -> dict:
    """Emisora IFRS en TWD con FY2024 mapeado y conceptos que el mapa no cubre."""
    ifrs = {
        "Revenue": _node("TWD", _annual(2894, 2024, ACCN_2024, "2025-04-17")),
        "ProfitLoss": _node("TWD", _annual(1173, 2024, ACCN_2024, "2025-04-17")),
        # Reportados por la emisora, fuera del mapa:
        "OtherReserves": _node(
            "TWD", _annual(600, 2023, ACCN_2024, "2025-04-17"),
            _annual(650, 2024, ACCN_2024, "2025-04-17"), label="Other reserves"),
        "MergerReserve": _node(
            "TWD", _annual(9, 2024, ACCN_2024, "2025-04-17"), label="Merger reserve"),
        "OldDiscontinuedLine": _node("TWD", _annual(1, 2019, "0001-20-000001", "2020-04-15")),
        # En el mapa pero SIN datos: no es "sin mapear".
        "CostOfSales": {"units": {}},
        **(extra_ifrs or {}),
    }
    return {"facts": {"dei": {}, "ifrs-full": ifrs, **(other or {})}}


def _fetch(payload: dict, **kwargs: object) -> tuple[str, dict[str, list[dict]]]:
    """Corre X.fetch sin red: (stdout, {nombre de archivo: filas})."""
    out = io.StringIO()
    files: dict[str, list[dict]] = {}
    with tempfile.TemporaryDirectory() as tmp, \
            mock.patch.object(X, "resolve_cik", return_value="0000000001"), \
            mock.patch.object(X, "_get_json", return_value=payload), \
            contextlib.redirect_stdout(out):
        X.fetch("TEST", Path(tmp), "Test test@example.com", **kwargs)
        for path in sorted(Path(tmp).iterdir()):
            with path.open(newline="", encoding="utf-8") as fh:
                files[path.name] = list(csv.DictReader(fh))
    return out.getvalue(), files


class UnmappedConceptsTests(unittest.TestCase):

    def setUp(self) -> None:
        facts = _payload()["facts"]["ifrs-full"]
        self.found = X.unmapped_concepts(facts, X.taxonomy_concepts("ifrs-full"))
        self.by_name = {u.concept: u for u in self.found}

    def test_lists_only_concepts_with_data_that_the_map_does_not_cover(self) -> None:
        self.assertEqual(set(self.by_name),
                         {"OtherReserves", "MergerReserve", "OldDiscontinuedLine"})

    def test_reports_fact_count_span_units_and_label(self) -> None:
        dep = self.by_name["OtherReserves"]
        self.assertEqual((dep.facts, dep.first_end, dep.last_end), (2, "2023-12-31", "2024-12-31"))
        self.assertEqual((dep.units, dep.label, dep.last_form), ("TWD", "Other reserves", "20-F"))
        self.assertEqual(dep.last_filed, "2025-04-17")

    def test_most_recent_and_most_reported_come_first(self) -> None:
        # Lo vigente antes que lo descontinuado; a igual cierre, mas hechos primero.
        self.assertEqual([u.concept for u in self.found],
                         ["OtherReserves", "MergerReserve", "OldDiscontinuedLine"])


class FetchDiagnosticsTests(unittest.TestCase):

    def test_the_count_of_unmapped_concepts_is_always_printed(self) -> None:
        out, files = _fetch(_payload())
        self.assertIn("3 conceptos con datos sin mapear", out)
        self.assertIn("--unmapped", out)
        self.assertEqual(list(files), ["xbrl_facts_TEST.csv"])       # sin bandera, sin archivo

    def test_unmapped_flag_writes_the_list_and_prints_the_top(self) -> None:
        out, files = _fetch(_payload(), unmapped=True)
        self.assertIn("xbrl_unmapped_TEST.csv", files)
        rows = files["xbrl_unmapped_TEST.csv"]
        self.assertEqual(list(rows[0]), list(X.UNMAPPED_FIELDNAMES))
        self.assertEqual([r["concept"] for r in rows],
                         ["OtherReserves", "MergerReserve", "OldDiscontinuedLine"])
        self.assertEqual(rows[0]["taxonomy"], "ifrs-full")
        self.assertEqual(rows[0]["facts"], "2")
        self.assertIn("OtherReserves", out)
        self.assertIn("xbrl_unmapped_TEST.csv", out)

    def test_no_unmapped_concepts_says_so_and_writes_an_empty_list(self) -> None:
        payload = _payload()
        for name in ("OtherReserves", "MergerReserve", "OldDiscontinuedLine"):
            del payload["facts"]["ifrs-full"][name]
        out, files = _fetch(payload, unmapped=True)
        self.assertIn("0 conceptos con datos sin mapear", out)
        self.assertEqual(files["xbrl_unmapped_TEST.csv"], [])

    def test_year_reported_only_in_unmapped_concepts_raises_a_warning(self) -> None:
        # FY2025 existe en el payload, pero solo bajo un concepto fuera del mapa:
        # el CSV termina en FY2024 y el tool lo dice en vez de callar.
        extra = {"RevenueFromContractsWithCustomers": _node(
            "TWD", _annual(3800, 2025, ACCN_2025, "2026-04-16"))}
        out, _ = _fetch(_payload(extra_ifrs=extra))
        self.assertIn("[aviso] el payload trae cierres anuales posteriores", out)
        self.assertIn("2025-12-31", out)
        self.assertIn("2024-12-31", out)
        self.assertIn("RevenueFromContractsWithCustomers", out)

    def test_no_gap_no_warning(self) -> None:
        out, _ = _fetch(_payload())
        self.assertNotIn("cierres anuales posteriores", out)

    def test_later_year_in_another_taxonomy_is_named(self) -> None:
        other = {"us-gaap": {"Assets": _node("USD", _annual(9, 2025, ACCN_2025, "2026-04-16"))}}
        out, _ = _fetch(_payload(other=other), taxonomy="ifrs-full")
        self.assertIn("cierres anuales posteriores", out)
        self.assertIn("us-gaap", out)


class TsmReportedConceptsTests(unittest.TestCase):
    """Conceptos que TSMC reporta de verdad (lista --unmapped, 2026-11-25)."""

    FLOWS = {
        "Revenue": 2894, "SalesAndMarketingExpense": 13, "GeneralAndAdministrativeExpense": 70,
        "FinanceCosts": 10, "DepreciationExpense": 650, "AmortisationExpense": 9,
        "IncomeTaxesPaidClassifiedAsOperatingActivities": 198,
        "ProceedsFromIssueOfBondsNotesAndDebentures": 40, "ProceedsFromNoncurrentBorrowings": 15,
        "RepaymentsOfBondsNotesAndDebentures": 12,
        "InterestPaidClassifiedAsFinancingActivities": 9,
    }
    STOCKS = {"PropertyPlantAndEquipment": 3200, "CurrentTradeReceivables": 270,
              "TradeAndOtherCurrentPayablesToTradeSuppliers": 74,
              "NoncurrentPortionOfNoncurrentBondsIssued": 900, "LongtermBorrowings": 30}
    EXPECTED = {
        ("is_sm", "FY2024"): 13, ("is_ga", "FY2024"): 70, ("is_finance_costs", "FY2024"): 10,
        ("cf_depreciation", "FY2024"): 650, ("cf_amortisation", "FY2024"): 9,
        ("cf_taxes_paid", "FY2024"): 198, ("cf_bonds_issued", "FY2024"): 40,
        ("cf_borrowings_lt_issued", "FY2024"): 15, ("cf_bonds_repaid", "FY2024"): 12,
        ("cf_interest_paid", "FY2024"): 9, ("bs_ppe", "4Q2024"): 3200,
        ("bs_ar", "4Q2024"): 270, ("bs_ap", "4Q2024"): 74,
        ("bs_bonds_lt", "4Q2024"): 900, ("bs_borrowings_lt", "4Q2024"): 30,
    }

    def setUp(self) -> None:
        ifrs = {name: _node("TWD", _annual(v, 2024, ACCN_2024, "2025-04-17"))
                for name, v in self.FLOWS.items()}
        for name, v in self.STOCKS.items():
            ifrs[name] = _node("TWD", {"end": "2024-12-31", "val": v, "accn": ACCN_2024,
                                       "filed": "2025-04-17", "form": "20-F"})
        ifrs["BasicEarningsLossPerShare"] = _node(
            "TWD/shares", _annual(45.25, 2024, ACCN_2024, "2025-04-17"))
        self.out, files = _fetch({"facts": {"ifrs-full": ifrs}})
        self.rows = {(r["canon"], r["period"]): r for r in files["xbrl_facts_TEST.csv"]}

    def test_each_reported_concept_lands_on_its_canon_line(self) -> None:
        got = {key: float(self.rows[key]["value"]) for key in self.EXPECTED if key in self.rows}
        self.assertEqual(got, {k: float(v) for k, v in self.EXPECTED.items()})

    def test_additive_concepts_never_share_a_canon(self) -> None:
        # Depreciacion + amortizacion y bonos + prestamos se SUMAN: si
        # compartieran canon, una pisaria a la otra dentro del mismo filing.
        canon = X.IFRS_CONCEPT_MAP
        for a, b in (("DepreciationExpense", "AmortisationExpense"),
                     ("ProceedsFromIssueOfBondsNotesAndDebentures", "ProceedsFromNoncurrentBorrowings"),
                     ("NoncurrentPortionOfNoncurrentBondsIssued", "LongtermBorrowings"),
                     ("SalesAndMarketingExpense", "GeneralAndAdministrativeExpense")):
            with self.subTest(pair=(a, b)):
                self.assertNotEqual(canon[a], canon[b])

    def test_payments_carry_the_outflow_sign_flag(self) -> None:
        for canon in ("cf_taxes_paid", "cf_bonds_repaid", "cf_interest_paid"):
            with self.subTest(canon=canon):
                self.assertEqual(self.rows[(canon, "FY2024")]["sign"], X.SIGN_OUTFLOW_POSITIVE)
        self.assertEqual(self.rows[("cf_bonds_issued", "FY2024")]["sign"], "")

    def test_basic_eps_is_never_derived_for_the_fourth_quarter(self) -> None:
        self.assertIn("is_eps_basic", X.NON_ADDITIVE)
        self.assertIn(("is_eps_basic", "FY2024"), self.rows)

    def test_console_states_how_far_companyfacts_reaches(self) -> None:
        self.assertIn("ultimo cierre anual en companyfacts (ifrs-full): 2024-12-31", self.out)


class CliTests(unittest.TestCase):

    def test_unmapped_flag_reaches_fetch(self) -> None:
        with mock.patch.object(X, "fetch") as fetch:
            X.main(["xbrl_fetch.py", "TSM", "--ua", "Test test@example.com", "--unmapped"])
            self.assertTrue(fetch.call_args.kwargs["unmapped"])
            X.main(["xbrl_fetch.py", "TSM", "--ua", "Test test@example.com"])
            self.assertFalse(fetch.call_args.kwargs["unmapped"])


if __name__ == "__main__":
    unittest.main()
