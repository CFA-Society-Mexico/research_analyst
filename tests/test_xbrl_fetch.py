"""Pruebas de tools/xbrl_fetch.py (stdlib unittest, sin red).

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

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import xbrl_fetch as X  # noqa: E402


def _label(start: str, end: str, fye: int) -> str | None:
    return X.period_label({"start": start, "end": end}, fye)


class PeriodLabelTests(unittest.TestCase):
    """Anios fiscales de 52/53 semanas: el cierre cae a veces en el mes siguiente."""

    # Cierres anuales reales de Costco (domingo mas cercano a fin de agosto).
    COSTCO_FY = [("2021-08-30", "2022-08-28"), ("2022-08-29", "2023-09-03"),
                 ("2023-09-04", "2024-09-01"), ("2024-09-02", "2025-08-31")]

    def test_costco_like_fiscal_years_keep_their_own_label(self) -> None:
        facts = {"Revenues": {"units": {"USD": [
            {"start": s, "end": e, "val": 1} for s, e in self.COSTCO_FY]}}}
        fye = X.detect_fye_month(facts)
        labels = [_label(s, e, fye) for s, e in self.COSTCO_FY]
        self.assertEqual(fye, 8)
        self.assertEqual(labels, ["FY2022", "FY2023", "FY2024", "FY2025"])

    def test_quarters_closing_early_next_month(self) -> None:
        # Starbucks FY2023: cierre de septiembre, trimestres que cierran el dia 1-2.
        cases = [("2022-10-03", "2023-01-01", "1Q2023"),
                 ("2023-01-02", "2023-04-02", "2Q2023"),
                 ("2023-04-03", "2023-07-02", "3Q2023"),
                 ("2022-10-03", "2023-10-01", "FY2023")]
        for start, end, want in cases:
            with self.subTest(end=end):
                self.assertEqual(_label(start, end, 9), want)

    def test_costco_twelve_week_quarters(self) -> None:
        cases = [("2022-08-29", "2022-11-20", "1Q2023"),
                 ("2022-11-21", "2023-02-12", "2Q2023"),
                 ("2023-02-13", "2023-05-07", "3Q2023")]
        for start, end, want in cases:
            with self.subTest(end=end):
                self.assertEqual(_label(start, end, 8), want)

    def test_regular_calendars_unchanged(self) -> None:
        cases = [("2024-09-29", "2024-12-28", 9, "1Q2025"),   # Apple
                 ("2024-09-29", "2025-09-27", 9, "FY2025"),
                 ("2025-01-01", "2025-03-31", 12, "1Q2025"),  # anio calendario
                 ("2025-02-01", "2025-04-30", 1, "1Q2026"),   # Walmart
                 ("2024-02-01", "2025-01-31", 1, "FY2025")]
        for start, end, fye, want in cases:
            with self.subTest(end=end, fye=fye):
                self.assertEqual(_label(start, end, fye), want)

    def test_balance_instant_at_year_end_is_4q(self) -> None:
        self.assertEqual(X.period_label({"end": "2023-09-03"}, 8), "4Q2023")


def _row(value: float, months: int | None, end: str = "") -> dict:
    return {"value": value, "months": months, "end": end, "tag": "observado"}


class DeaccumulateTests(unittest.TestCase):

    def test_cash_flow_ytd_becomes_quarterly_and_4q_is_derived(self) -> None:
        rows = {("cf_cfo", "1Q2025"): _row(100, 3),
                ("cf_cfo", "2Q2025"): _row(210, 6),
                ("cf_cfo", "3Q2025"): _row(330, 9, "2025-06-28"),
                ("cf_cfo", "FY2025"): _row(460, 12)}
        X.deaccumulate(rows)
        got = [rows[("cf_cfo", f"{q}Q2025")]["value"] for q in (1, 2, 3, 4)]
        self.assertEqual(got, [100, 110, 120, 130])
        self.assertEqual(rows[("cf_cfo", "4Q2025")]["start"], "2025-06-29")

    def test_income_statement_4q_is_derived_but_not_eps(self) -> None:
        rows = {}
        for canon in ("is_ns_total", "is_eps_diluted"):
            for q, v in ((1, 100), (2, 110), (3, 120)):
                rows[(canon, f"{q}Q2025")] = _row(v, 3)
            rows[(canon, "FY2025")] = _row(460, 12)
        X.deaccumulate(rows)
        self.assertEqual(rows[("is_ns_total", "4Q2025")]["value"], 130)
        self.assertNotIn(("is_eps_diluted", "4Q2025"), rows)

    def test_ytd_without_prior_quarter_is_flagged_not_passed_as_quarter(self) -> None:
        rows = {("cf_cfo", "2Q2025"): _row(210, 6),
                ("cf_cfo", "3Q2025"): _row(330, 9),
                ("cf_cfo", "FY2025"): _row(460, 12)}
        notes = X.deaccumulate(rows)
        q2 = rows[("cf_cfo", "2Q2025")]
        self.assertEqual((q2["value"], q2["months"]), (210, 6))
        self.assertTrue(q2["tag"].startswith(X.YTD_UNRESOLVED_TAG))
        self.assertTrue(any(n.startswith("[aviso]") for n in notes))
        # 3Q y 4Q siguen siendo calculables desde los acumulados.
        self.assertEqual(rows[("cf_cfo", "3Q2025")]["value"], 120)
        self.assertEqual(rows[("cf_cfo", "4Q2025")]["value"], 130)

    def test_ytd_after_discrete_quarter_uses_cumulative_not_prior_row(self) -> None:
        rows = {("cf_cfo", "1Q2025"): _row(100, 3),
                ("cf_cfo", "2Q2025"): _row(110, 3),     # ya discreto
                ("cf_cfo", "3Q2025"): _row(330, 9)}     # acumulado 9m
        X.deaccumulate(rows)
        self.assertEqual(rows[("cf_cfo", "3Q2025")]["value"], 120)

    def test_reported_discrete_4q_is_not_overwritten(self) -> None:
        rows = {("is_ns_total", f"{q}Q2025"): _row(100, 3) for q in (1, 2, 3)}
        rows[("is_ns_total", "4Q2025")] = _row(999, 3)
        rows[("is_ns_total", "FY2025")] = _row(1299, 12)
        X.deaccumulate(rows)
        self.assertEqual(rows[("is_ns_total", "4Q2025")]["value"], 999)


def _run_fetch(payload: dict, **kwargs: str) -> tuple[dict[tuple[str, str], dict], str, list[str]]:
    """Corre X.fetch sin red: (filas por (canon, periodo), stdout, encabezado CSV)."""
    orig = (X.resolve_cik, X._get_json)
    X.resolve_cik = lambda ticker, ua: "0000000001"
    X._get_json = lambda url, ua: payload
    buf = io.StringIO()
    try:
        with tempfile.TemporaryDirectory() as tmp, contextlib.redirect_stdout(buf):
            out = X.fetch("TEST", Path(tmp), "Test test@example.com", **kwargs)
            with out.open(encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                rows = {(r["canon"], r["period"]): r for r in reader}
                return rows, buf.getvalue(), list(reader.fieldnames or [])
    finally:
        X.resolve_cik, X._get_json = orig


def _fact(val: float, end: str, accn: str, filed: str, start: str = "",
          form: str = "10-K") -> dict:
    item = {"end": end, "val": val, "accn": accn, "filed": filed, "form": form}
    if start:
        item["start"] = start
    return item


def _facts(*entries: tuple[str, str, dict]) -> dict:
    """(concepto, unidad, hecho)... -> nodo de taxonomia de companyfacts."""
    node: dict = {}
    for concept, unit, item in entries:
        node.setdefault(concept, {"units": {}})["units"].setdefault(unit, []).append(item)
    return node


class FetchTests(unittest.TestCase):

    def _fetch(self, facts: dict) -> dict[tuple[str, str], dict]:
        return _run_fetch({"facts": {"us-gaap": facts}})[0]

    def test_costco_like_years_do_not_collide(self) -> None:
        facts = {"Revenues": {"units": {"USD": [
            {"start": s, "end": e, "val": v, "form": "10-K", "filed": "2025-10-08"}
            for (s, e), v in zip(PeriodLabelTests.COSTCO_FY, (10, 20, 30, 40))]}}}
        got = self._fetch(facts)
        fy = {p: int(r["value"]) for (c, p), r in got.items() if p.startswith("FY")}
        self.assertEqual(fy, {"FY2022": 10, "FY2023": 20, "FY2024": 30, "FY2025": 40})

    def test_legacy_revenue_concept_fills_early_history(self) -> None:
        facts = {
            "SalesRevenueNet": {"units": {"USD": [
                {"start": "2015-09-27", "end": "2016-09-24", "val": 215,
                 "form": "10-K", "filed": "2016-10-26"},
                {"start": "2016-09-25", "end": "2017-09-30", "val": 1,
                 "form": "10-K", "filed": "2017-11-03"}]}},
            "RevenueFromContractWithCustomerExcludingAssessedTax": {"units": {"USD": [
                {"start": "2016-09-25", "end": "2017-09-30", "val": 229,
                 "form": "10-K", "filed": "2019-10-31"}]}},
        }
        got = self._fetch(facts)
        self.assertEqual(int(got[("is_ns_total", "FY2016")]["value"]), 215)
        # Mismo periodo en ambos conceptos: gana el filing mas reciente.
        self.assertEqual(int(got[("is_ns_total", "FY2017")]["value"]), 229)

    def test_maturities_concept_without_debt_suffix_is_captured(self) -> None:
        facts = {"ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities": {
            "units": {"USD": [{"start": "2024-09-29", "end": "2025-09-27", "val": 5,
                               "form": "10-K", "filed": "2025-10-31"}]}}}
        got = self._fetch(facts)
        self.assertIn(("cf_mature_securities", "FY2025"), got)


# --- Issue 36: un solo vintage (accession) por estado y periodo -------------

ORIGINAL, RESTATED, STRAY = "0001-18-000001", "0001-19-000002", "0001-20-000003"
BS_ORIGINAL = {"Assets": 1000, "AssetsCurrent": 400, "LiabilitiesCurrent": 200,
               "Liabilities": 600, "StockholdersEquity": 400,
               "CashAndCashEquivalentsAtCarryingValue": 150}
BS_RESTATED = {"Assets": 1010, "AssetsCurrent": 405, "LiabilitiesCurrent": 203,
               "Liabilities": 606, "StockholdersEquity": 404,
               "CashAndCashEquivalentsAtCarryingValue": 151}
BS_CANONS = ("bs_ta", "bs_ca", "bs_cl", "bs_tl", "bs_equity", "bs_cash")


def _balance(by_accn: dict[str, tuple[str, dict[str, float]]],
             end: str = "2017-12-31") -> dict:
    return _facts(*[(concept, "USD", _fact(val, end, accn, filed))
                    for accn, (filed, values) in by_accn.items()
                    for concept, val in values.items()])


def _only(values: dict[str, float], *concepts: str) -> dict[str, float]:
    return {c: values[c] for c in concepts}


class VintageSelectionTests(unittest.TestCase):

    def test_balance_lines_of_a_period_share_one_accession(self) -> None:
        # Original, re-presentado y un filing posterior que solo repite el
        # capital (estado de cambios en el capital): con dedup por concepto el
        # capital salia del tercero y el balance dejaba de cuadrar.
        facts = _balance({
            ORIGINAL: ("2018-04-19", BS_ORIGINAL),
            RESTATED: ("2019-04-17", BS_RESTATED),
            STRAY: ("2020-04-15", {"StockholdersEquity": 399}),
        })
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        rows = {c: got[(c, "4Q2017")] for c in BS_CANONS}
        self.assertEqual({r["accession"] for r in rows.values()}, {RESTATED})
        self.assertEqual({r["filed"] for r in rows.values()}, {"2019-04-17"})
        self.assertEqual({r["tag"] for r in rows.values()}, {"observado"})
        value = {c: float(r["value"]) for c, r in rows.items()}
        self.assertEqual(value["bs_ta"], value["bs_tl"] + value["bs_equity"])

    def test_later_accession_below_coverage_threshold_does_not_win(self) -> None:
        four = ("Assets", "AssetsCurrent", "LiabilitiesCurrent", "Liabilities")
        five = (*four, "StockholdersEquity")
        for concepts, want in ((four, ORIGINAL), (five, RESTATED)):
            with self.subTest(covered=len(concepts)):
                facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL),
                                  RESTATED: ("2019-04-17", _only(BS_RESTATED, *concepts))})
                concepts_map = list(X.CONCEPT_MAP.items())
                chosen, _ = X.select_by_vintage(
                    X.collect_candidates(facts, concepts_map, "USD", 12))
                self.assertEqual(chosen[("bs_ta", "4Q2017")].accession, want)

    def test_coverage_threshold_is_a_named_constant(self) -> None:
        self.assertEqual(X.VINTAGE_COVERAGE_MIN_PCT, 80)

    def test_fill_in_rows_are_tagged_and_counted(self) -> None:
        five = ("Assets", "AssetsCurrent", "LiabilitiesCurrent", "Liabilities",
                "StockholdersEquity")
        facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL),
                          RESTATED: ("2019-04-17", _only(BS_RESTATED, *five))})
        got, out, _ = _run_fetch({"facts": {"us-gaap": facts}})
        cash = got[("bs_cash", "4Q2017")]
        self.assertEqual(cash["accession"], ORIGINAL)
        self.assertEqual(cash["filed"], "2018-04-19")
        self.assertEqual(cash["tag"], f"observado ({X.VINTAGE_MIX_NOTE}: {ORIGINAL})")
        self.assertEqual(got[("bs_ta", "4Q2017")]["tag"], "observado")
        aviso = [ln for ln in out.splitlines()
                 if ln.startswith("[aviso]") and X.VINTAGE_MIX_NOTE in ln]
        self.assertEqual(len(aviso), 1)
        self.assertIn("1 fila", aviso[0])
        self.assertIn(f"bs_cash 4Q2017 <- {ORIGINAL}", out)

    def test_no_mix_no_warning(self) -> None:
        facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL),
                          RESTATED: ("2019-04-17", BS_RESTATED)})
        _, out, _ = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertNotIn(X.VINTAGE_MIX_NOTE, out)

    def test_same_filing_date_breaks_tie_by_accession_descending(self) -> None:
        facts = _balance({ORIGINAL: ("2019-04-17", BS_ORIGINAL),
                          RESTATED: ("2019-04-17", BS_RESTATED)})
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertEqual(got[("bs_ta", "4Q2017")]["accession"], max(ORIGINAL, RESTATED))

    def test_statements_of_one_period_choose_their_vintage_independently(self) -> None:
        # Un 10-K posterior repite 3 anios de resultados pero solo 2 balances:
        # el IS de FY2017 sale del filing mas nuevo y el balance del anterior.
        newest = "0001-20-000009"
        facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL),
                          RESTATED: ("2019-04-17", BS_RESTATED)})
        facts.update(_facts(*[
            ("Revenues", "USD", _fact(val, "2017-12-31", accn, filed, start="2017-01-01"))
            for accn, filed, val in ((ORIGINAL, "2018-04-19", 500),
                                     (RESTATED, "2019-04-17", 510),
                                     (newest, "2020-04-15", 520))]))
        got, out, _ = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertEqual(got[("is_ns_total_alt", "FY2017")]["accession"], newest)
        self.assertEqual(got[("bs_ta", "4Q2017")]["accession"], RESTATED)
        self.assertNotIn(X.VINTAGE_MIX_NOTE, out)

    def test_shortest_duration_wins_inside_one_accession(self) -> None:
        tenq = "0001-25-000007"
        facts = _facts(
            ("Revenues", "USD", _fact(210, "2025-06-30", tenq, "2025-08-01",
                                      start="2025-01-01", form="10-Q")),
            ("Revenues", "USD", _fact(110, "2025-06-30", tenq, "2025-08-01",
                                      start="2025-04-01", form="10-Q")))
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        row = got[("is_ns_total_alt", "2Q2025")]
        self.assertEqual((float(row["value"]), row["months"]), (110.0, "3"))

    def test_current_concept_beats_legacy_inside_one_accession(self) -> None:
        tenk = "0001-17-000070"
        facts = _facts(
            ("SalesRevenueNet", "USD",
             _fact(1, "2017-09-30", tenk, "2017-11-03", start="2016-09-25")),
            ("RevenueFromContractWithCustomerExcludingAssessedTax", "USD",
             _fact(229, "2017-09-30", tenk, "2017-11-03", start="2016-09-25")))
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        row = got[("is_ns_total", "FY2017")]
        self.assertEqual(float(row["value"]), 229.0)
        self.assertEqual(row["concept"], "RevenueFromContractWithCustomerExcludingAssessedTax")

    def test_deaccumulation_runs_on_selected_rows_and_4q_inherits_fy_accession(self) -> None:
        concept = "NetCashProvidedByUsedInOperatingActivities"
        tenk = "0001-26-000004"
        facts = _facts(
            (concept, "USD", _fact(100, "2025-03-31", "0001-25-000001", "2025-05-01",
                                   start="2025-01-01", form="10-Q")),
            (concept, "USD", _fact(210, "2025-06-30", "0001-25-000002", "2025-08-01",
                                   start="2025-01-01", form="10-Q")),
            (concept, "USD", _fact(330, "2025-09-30", "0001-25-000003", "2025-11-01",
                                   start="2025-01-01", form="10-Q")),
            (concept, "USD", _fact(460, "2025-12-31", tenk, "2026-02-01",
                                   start="2025-01-01")))
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        values = [float(got[("cf_cfo", f"{q}Q2025")]["value"]) for q in (1, 2, 3, 4)]
        self.assertEqual(values, [100, 110, 120, 130])
        q4 = got[("cf_cfo", "4Q2025")]
        self.assertEqual((q4["accession"], q4["filed"], q4["form"]),
                         (tenk, "2026-02-01", "10-K"))
        self.assertEqual(got[("cf_cfo", "2Q2025")]["accession"], "0001-25-000002")

    def test_derived_4q_from_a_fill_in_keeps_the_vintage_mark(self) -> None:
        # El FY del flujo operativo viene de un accession distinto al primario
        # del grupo (cf, FY2025): el 4Q derivado hereda la marca.
        primary, other = "0001-26-000004", "0001-26-000001"
        cfo = "NetCashProvidedByUsedInOperatingActivities"
        flows = ("NetCashProvidedByUsedInInvestingActivities",
                 "NetCashProvidedByUsedInFinancingActivities",
                 "DepreciationDepletionAndAmortization",
                 "ShareBasedCompensation",
                 "PaymentsToAcquirePropertyPlantAndEquipment")
        entries = [(c, "USD", _fact(10, "2025-12-31", primary, "2026-02-01",
                                    start="2025-01-01")) for c in flows]
        entries.append((cfo, "USD", _fact(460, "2025-12-31", other, "2026-01-15",
                                          start="2025-01-01")))
        entries += [(cfo, "USD", _fact(v, end, f"0001-25-00000{q}", f"2025-{2 + 3 * q:02d}-01",
                                       start="2025-01-01", form="10-Q"))
                    for q, end, v in ((1, "2025-03-31", 100), (2, "2025-06-30", 210),
                                      (3, "2025-09-30", 330))]
        got, _, _ = _run_fetch({"facts": {"us-gaap": _facts(*entries)}})
        note = f"{X.VINTAGE_MIX_NOTE}: {other}"
        self.assertIn(note, got[("cf_cfo", "FY2025")]["tag"])
        q4 = got[("cf_cfo", "4Q2025")]
        self.assertTrue(q4["tag"].startswith("derivado"))
        self.assertIn(note, q4["tag"])
        self.assertEqual(float(q4["value"]), 130.0)

    def test_facts_without_accession_fall_back_to_filing_date(self) -> None:
        facts = _facts(
            ("Assets", "USD", {"end": "2017-12-31", "val": 1, "filed": "2018-04-19", "form": "10-K"}),
            ("Assets", "USD", {"end": "2017-12-31", "val": 2, "filed": "2019-04-17", "form": "10-K"}))
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        row = got[("bs_ta", "4Q2017")]
        self.assertEqual((float(row["value"]), row["accession"]), (2.0, ""))


class CsvContractTests(unittest.TestCase):

    LEGACY_FIELDS = ["canon", "period", "value", "months", "start", "end",
                     "form", "filed", "concept", "unit", "source", "tag"]

    def test_new_columns_are_appended_at_the_end(self) -> None:
        facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL)})
        _, _, fieldnames = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertEqual(fieldnames, [*self.LEGACY_FIELDS, "accession", "sign"])

    def test_us_gaap_row_keeps_its_previous_content(self) -> None:
        facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL)})
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        row = got[("bs_ta", "4Q2017")]
        self.assertEqual({k: row[k] for k in self.LEGACY_FIELDS}, {
            "canon": "bs_ta", "period": "4Q2017", "value": "1000", "months": "",
            "start": "", "end": "2017-12-31", "form": "10-K", "filed": "2018-04-19",
            "concept": "Assets", "unit": "USD",
            "source": "XBRL companyfacts CIK0000000001 Assets", "tag": "observado"})


class SignConventionTests(unittest.TestCase):

    def test_outflow_concepts_are_flagged_and_value_is_not_flipped(self) -> None:
        tenk = "0001-26-000004"
        facts = _facts(*[
            (concept, "USD", _fact(val, "2025-12-31", tenk, "2026-02-01", start="2025-01-01"))
            for concept, val in (("PaymentsToAcquirePropertyPlantAndEquipment", 50),
                                 ("PaymentsOfDividends", 20),
                                 ("RepaymentsOfLongTermDebt", 7),
                                 ("NetCashProvidedByUsedInOperatingActivities", 460),
                                 ("NetCashProvidedByUsedInInvestingActivities", -55))])
        got, out, _ = _run_fetch({"facts": {"us-gaap": facts}})
        for canon, val in (("cf_capex", 50.0), ("cf_dividends", 20.0), ("cf_debt_repaid", 7.0)):
            with self.subTest(canon=canon):
                row = got[(canon, "FY2025")]
                self.assertEqual((float(row["value"]), row["sign"]),
                                 (val, X.SIGN_OUTFLOW_POSITIVE))
        self.assertEqual(got[("cf_cfo", "FY2025")]["sign"], "")
        self.assertEqual(float(got[("cf_cfi", "FY2025")]["value"]), -55.0)
        self.assertEqual(got[("cf_cfi", "FY2025")]["sign"], "")
        info = [ln for ln in out.splitlines()
                if ln.startswith("[i]") and X.SIGN_OUTFLOW_POSITIVE in ln]
        self.assertEqual(len(info), 1)

    def test_sign_label_is_the_documented_literal(self) -> None:
        self.assertEqual(X.SIGN_OUTFLOW_POSITIVE, "salida(+)")

    def test_outflow_set_is_explicit_and_only_names_mapped_cash_flow_concepts(self) -> None:
        mapped = {**X.CONCEPT_MAP, **X.LEGACY_CONCEPT_MAP, **X.IFRS_CONCEPT_MAP}
        self.assertIsInstance(X.OUTFLOW_POSITIVE_CONCEPTS, frozenset)
        for concept in X.OUTFLOW_POSITIVE_CONCEPTS:
            with self.subTest(concept=concept):
                self.assertIn(concept, mapped)
                self.assertTrue(mapped[concept].startswith("cf_"))
        for concept in ("PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
                        "DividendsPaidClassifiedAsFinancingActivities",
                        "IncomeTaxesPaidRefundClassifiedAsOperatingActivities",
                        "RepaymentsOfBorrowingsClassifiedAsFinancingActivities",
                        "PaymentsToAcquirePropertyPlantAndEquipment",
                        "PaymentsForRepurchaseOfCommonStock"):
            self.assertIn(concept, X.OUTFLOW_POSITIVE_CONCEPTS)
        for concept in ("NetCashProvidedByUsedInOperatingActivities",
                        "ProceedsFromIssuanceOfLongTermDebt",
                        "ProceedsFromBorrowingsClassifiedAsFinancingActivities"):
            self.assertNotIn(concept, X.OUTFLOW_POSITIVE_CONCEPTS)

    def test_derived_and_deaccumulated_rows_keep_the_sign_flag(self) -> None:
        concept = "PaymentsToAcquirePropertyPlantAndEquipment"
        facts = _facts(*[
            (concept, "USD", _fact(v, end, f"0001-25-00000{i}", filed,
                                   start="2025-01-01", form=form))
            for i, (end, filed, v, form) in enumerate(
                (("2025-03-31", "2025-05-01", 10, "10-Q"),
                 ("2025-06-30", "2025-08-01", 25, "10-Q"),
                 ("2025-09-30", "2025-11-01", 45, "10-Q"),
                 ("2025-12-31", "2026-02-01", 70, "10-K")), start=1)])
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertEqual([float(got[("cf_capex", f"{q}Q2025")]["value"]) for q in (1, 2, 3, 4)],
                         [10, 15, 20, 25])
        self.assertEqual({got[("cf_capex", f"{q}Q2025")]["sign"] for q in (1, 2, 3, 4)},
                         {X.SIGN_OUTFLOW_POSITIVE})


# --- Issue 35: emisoras IFRS (taxonomia ifrs-full, moneda de reporte) --------

TSM_20F = "0001-25-000020"


def _ifrs_year(values: dict[str, float], unit: str = "TWD", accn: str = TSM_20F,
               filed: str = "2025-04-17", year: int = 2024) -> list[tuple[str, str, dict]]:
    return [(concept, unit, _fact(val, f"{year}-12-31", accn, filed,
                                  start=f"{year}-01-01", form="20-F"))
            for concept, val in values.items()]


def _ifrs_balance(values: dict[str, float], unit: str = "TWD", accn: str = TSM_20F,
                  filed: str = "2025-04-17", end: str = "2024-12-31") -> list[tuple[str, str, dict]]:
    return [(concept, unit, _fact(val, end, accn, filed, form="20-F"))
            for concept, val in values.items()]


IFRS_FLOWS = {"Revenue": 2894, "CostOfSales": 1270, "GrossProfit": 1624,
              "ProfitLossFromOperatingActivities": 1322, "ProfitLossBeforeTax": 1406,
              "IncomeTaxExpenseContinuingOperations": 233, "ProfitLoss": 1173,
              "ProfitLossAttributableToOwnersOfParent": 1172,
              "CashFlowsFromUsedInOperatingActivities": 1826,
              "CashFlowsFromUsedInInvestingActivities": -865,
              "CashFlowsFromUsedInFinancingActivities": -346,
              "EffectOfExchangeRateChangesOnCashAndCashEquivalents": 48,
              "IncreaseDecreaseInCashAndCashEquivalents": 663,
              "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities": 956,
              "DividendsPaidClassifiedAsFinancingActivities": 363,
              "IncomeTaxesPaidRefundClassifiedAsOperatingActivities": 198,
              "ProceedsFromBorrowingsClassifiedAsFinancingActivities": 30,
              "RepaymentsOfBorrowingsClassifiedAsFinancingActivities": 12}
IFRS_BALANCE = {"Assets": 6692, "CurrentAssets": 3088, "CurrentLiabilities": 1264,
                "Liabilities": 2368, "Equity": 4324,
                "EquityAttributableToOwnersOfParent": 4286,
                "CashAndCashEquivalents": 2128, "Inventories": 288,
                "RetainedEarnings": 4000}


def _ifrs_payload(extra: list[tuple[str, str, dict]] | None = None) -> dict:
    entries = [*_ifrs_year(IFRS_FLOWS), *_ifrs_balance(IFRS_BALANCE),
               ("DilutedEarningsLossPerShare", "TWD/shares",
                _fact(45.25, "2024-12-31", TSM_20F, "2025-04-17",
                      start="2024-01-01", form="20-F")),
               *(extra or [])]
    dei = {"EntityCommonStockSharesOutstanding": {"units": {"shares": [
        {"end": "2024-12-31", "val": 25930, "accn": TSM_20F, "filed": "2025-04-17",
         "form": "20-F"}]}}}
    return {"facts": {"dei": dei, "ifrs-full": _facts(*entries)}}


class IfrsFetchTests(unittest.TestCase):

    def test_ifrs_payload_in_twd_yields_canon_rows(self) -> None:
        got, out, _ = _run_fetch(_ifrs_payload())
        want = {("is_ns_total", "FY2024"): 2894, ("is_cogs_total", "FY2024"): 1270,
                ("is_gross", "FY2024"): 1624, ("is_ebit", "FY2024"): 1322,
                ("is_ebt", "FY2024"): 1406, ("is_tax", "FY2024"): 233,
                ("is_ni", "FY2024"): 1172, ("is_eps_diluted", "FY2024"): 45.25,
                ("bs_ta", "4Q2024"): 6692, ("bs_ca", "4Q2024"): 3088,
                ("bs_cl", "4Q2024"): 1264, ("bs_tl", "4Q2024"): 2368,
                ("bs_equity", "4Q2024"): 4286, ("bs_cash", "4Q2024"): 2128,
                ("bs_inv", "4Q2024"): 288, ("bs_re", "4Q2024"): 4000,
                ("cf_cfo", "FY2024"): 1826, ("cf_cfi", "FY2024"): -865,
                ("cf_cff", "FY2024"): -346, ("cf_net_change", "FY2024"): 663,
                ("cf_capex", "FY2024"): 956, ("cf_dividends", "FY2024"): 363,
                ("cf_taxes_paid", "FY2024"): 198, ("cf_debt_issued", "FY2024"): 30,
                ("cf_debt_repaid", "FY2024"): 12}
        for key, val in want.items():
            with self.subTest(key=key):
                self.assertEqual(float(got[key]["value"]), float(val))
        self.assertEqual(got[("is_ns_total", "FY2024")]["unit"], "TWD")
        self.assertEqual(got[("is_eps_diluted", "FY2024")]["unit"], "TWD/shares")
        self.assertEqual(got[("is_ns_total", "FY2024")]["accession"], TSM_20F)
        self.assertEqual(got[("is_ns_total", "FY2024")]["source"],
                         "XBRL companyfacts CIK0000000001 ifrs-full:Revenue")
        self.assertIn("ifrs-full", out)
        self.assertIn("TWD", out)

    def test_ifrs_outflows_carry_the_sign_flag(self) -> None:
        got, _, _ = _run_fetch(_ifrs_payload())
        for canon in ("cf_capex", "cf_dividends", "cf_taxes_paid", "cf_debt_repaid"):
            with self.subTest(canon=canon):
                self.assertEqual(got[(canon, "FY2024")]["sign"], X.SIGN_OUTFLOW_POSITIVE)
        self.assertEqual(got[("cf_debt_issued", "FY2024")]["sign"], "")

    def test_alternates_do_not_overwrite_each_other(self) -> None:
        got, _, _ = _run_fetch(_ifrs_payload())
        self.assertEqual(float(got[("is_ni", "FY2024")]["value"]), 1172.0)
        self.assertEqual(float(got[("is_ni_alt", "FY2024")]["value"]), 1173.0)
        self.assertEqual(float(got[("bs_equity", "4Q2024")]["value"]), 4286.0)
        self.assertEqual(float(got[("bs_equity_alt", "4Q2024")]["value"]), 4324.0)
        # Identidad del balance con el capital TOTAL (incluye no controladora).
        self.assertEqual(float(got[("bs_ta", "4Q2024")]["value"]),
                         float(got[("bs_tl", "4Q2024")]["value"])
                         + float(got[("bs_equity_alt", "4Q2024")]["value"]))

    def test_finance_income_is_not_interest_income(self) -> None:
        extra = _ifrs_year({"FinanceIncome": 88, "RevenueFromInterest": 71})
        got, _, _ = _run_fetch(_ifrs_payload(extra))
        self.assertEqual(X.IFRS_CONCEPT_MAP["RevenueFromInterest"], "is_interest_income")
        self.assertNotEqual(X.IFRS_CONCEPT_MAP["FinanceIncome"], "is_interest_income")
        self.assertEqual(X.IFRS_CONCEPT_MAP["FinanceIncome"], "is_other_income_finance")
        self.assertEqual(float(got[("is_interest_income", "FY2024")]["value"]), 71.0)
        self.assertEqual(got[("is_interest_income", "FY2024")]["concept"], "RevenueFromInterest")
        self.assertEqual(float(got[("is_other_income_finance", "FY2024")]["value"]), 88.0)

    def test_additional_paid_in_capital_is_not_mapped(self) -> None:
        self.assertNotIn("AdditionalPaidinCapital", X.IFRS_CONCEPT_MAP)
        extra = _ifrs_balance({"AdditionalPaidinCapital": 690})
        got, _, _ = _run_fetch(_ifrs_payload(extra))
        self.assertNotIn("AdditionalPaidinCapital", {r["concept"] for r in got.values()})

    def test_fx_effect_on_cash_is_captured_and_the_cash_roll_closes(self) -> None:
        got, _, _ = _run_fetch(_ifrs_payload())
        fx = got[("cf_fx_effect", "FY2024")]
        self.assertEqual(fx["concept"], "EffectOfExchangeRateChangesOnCashAndCashEquivalents")
        total = sum(float(got[(c, "FY2024")]["value"])
                    for c in ("cf_cfo", "cf_cfi", "cf_cff", "cf_fx_effect"))
        self.assertEqual(total, float(got[("cf_net_change", "FY2024")]["value"]))

    def test_us_gaap_fx_effect_concepts_feed_the_same_canon(self) -> None:
        current = "EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
        legacy = "EffectOfExchangeRateOnCashAndCashEquivalents"
        self.assertEqual(X.CONCEPT_MAP[current], "cf_fx_effect")
        self.assertEqual({**X.CONCEPT_MAP, **X.LEGACY_CONCEPT_MAP}[legacy], "cf_fx_effect")
        facts = _facts(
            (legacy, "USD", _fact(-3, "2016-12-31", "0001-17-000001", "2017-02-01",
                                  start="2016-01-01")),
            (current, "USD", _fact(-5, "2024-12-31", "0001-25-000001", "2025-02-01",
                                   start="2024-01-01")),
            (legacy, "USD", _fact(-4, "2024-12-31", "0001-25-000001", "2025-02-01",
                                  start="2024-01-01")))
        got, _, _ = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertEqual(float(got[("cf_fx_effect", "FY2016")]["value"]), -3.0)
        self.assertEqual(float(got[("cf_fx_effect", "FY2024")]["value"]), -5.0)

    def test_ifrs_canons_reuse_us_gaap_names_plus_a_declared_list(self) -> None:
        us_canons = set(X.CONCEPT_MAP.values()) | set(X.LEGACY_CONCEPT_MAP.values())
        new = set(X.IFRS_CONCEPT_MAP.values()) - us_canons
        self.assertEqual(new, {"is_ni_alt", "bs_equity_alt", "is_interest_income",
                               "is_other_income_finance", "cf_taxes_paid"})

    def test_fiscal_year_end_is_detected_on_the_chosen_taxonomy(self) -> None:
        # Cierre en marzo: sin detectar sobre ifrs-full el saldo de cierre
        # saldria como 1Q (default diciembre) en lugar de 4Q.
        entries = [("Revenue", "JPY", _fact(v, f"{y}-03-31", f"0001-{y}-000001",
                                            f"{y}-06-20", start=f"{y - 1}-04-01", form="20-F"))
                   for y, v in ((2023, 900), (2024, 950))]
        entries.append(("Assets", "JPY", _fact(7000, "2024-03-31", "0001-2024-000001",
                                               "2024-06-20", form="20-F")))
        got, out, _ = _run_fetch({"facts": {"ifrs-full": _facts(*entries)}})
        self.assertIn(("bs_ta", "4Q2024"), got)
        self.assertIn(("is_ns_total", "FY2024"), got)
        self.assertIn("mes 3", out)


class TaxonomyAndCurrencyTests(unittest.TestCase):

    def _mixed(self) -> dict:
        payload = _ifrs_payload()
        payload["facts"]["us-gaap"] = _facts(
            ("Assets", "USD", _fact(1, "2024-12-31", TSM_20F, "2025-04-17", form="20-F")))
        return payload

    def test_auto_taxonomy_picks_the_one_with_more_mapped_concepts(self) -> None:
        got, out, _ = _run_fetch(self._mixed())
        self.assertEqual(got[("bs_ta", "4Q2024")]["unit"], "TWD")
        line = next(ln for ln in out.splitlines() if "taxonomia" in ln)
        self.assertIn("ifrs-full", line)
        self.assertIn("auto", line)

    def test_choose_taxonomy_is_pure_and_prefers_us_gaap_on_ties(self) -> None:
        all_facts = self._mixed()["facts"]
        name, counts = X.choose_taxonomy(all_facts, "auto")
        self.assertEqual(name, "ifrs-full")
        self.assertEqual(counts["us-gaap"], 1)
        self.assertGreater(counts["ifrs-full"], 20)
        self.assertEqual(X.choose_taxonomy({}, "auto")[0], "us-gaap")
        tie = {"us-gaap": _facts(("Assets", "USD", _fact(1, "2024-12-31", "a", "2025-01-01"))),
               "ifrs-full": _facts(("Assets", "EUR", _fact(1, "2024-12-31", "a", "2025-01-01")))}
        self.assertEqual(X.choose_taxonomy(tie, "auto")[0], "us-gaap")

    def test_explicit_taxonomy_overrides_auto(self) -> None:
        got, out, _ = _run_fetch(self._mixed(), taxonomy="us-gaap")
        self.assertEqual(set(got), {("bs_ta", "4Q2024")})
        self.assertEqual(got[("bs_ta", "4Q2024")]["unit"], "USD")
        self.assertEqual(float(got[("bs_ta", "4Q2024")]["value"]), 1.0)

    def test_unknown_taxonomy_is_rejected(self) -> None:
        with self.assertRaises(SystemExit) as ctx:
            _run_fetch(self._mixed(), taxonomy="mx-ifrs")
        self.assertTrue(str(ctx.exception).startswith("[x]"))

    def test_auto_currency_takes_the_unit_with_most_facts_and_warns(self) -> None:
        extra = _ifrs_year({"Revenue": 90, "ProfitLoss": 36}, unit="USD")
        got, out, _ = _run_fetch(_ifrs_payload(extra))
        self.assertEqual(got[("is_ns_total", "FY2024")]["unit"], "TWD")
        self.assertEqual(float(got[("is_ns_total", "FY2024")]["value"]), 2894.0)
        self.assertEqual({r["unit"] for r in got.values()}, {"TWD", "TWD/shares"})
        aviso = [ln for ln in out.splitlines()
                 if ln.startswith("[aviso]") and "moneda" in ln]
        self.assertEqual(len(aviso), 1)
        self.assertIn("TWD", aviso[0])
        self.assertIn("USD", aviso[0])

    def test_single_currency_does_not_warn(self) -> None:
        _, out, _ = _run_fetch(_ifrs_payload())
        self.assertFalse([ln for ln in out.splitlines()
                          if ln.startswith("[aviso]") and "moneda" in ln])
        self.assertTrue([ln for ln in out.splitlines()
                         if ln.startswith("[ok] moneda") and "TWD" in ln])

    def test_explicit_currency_overrides_auto(self) -> None:
        extra = _ifrs_year({"Revenue": 90, "ProfitLoss": 36}, unit="USD")
        got, out, _ = _run_fetch(_ifrs_payload(extra), currency="usd")
        self.assertEqual({r["unit"] for r in got.values()}, {"USD"})
        self.assertEqual(float(got[("is_ns_total", "FY2024")]["value"]), 90.0)
        self.assertNotIn(("bs_ta", "4Q2024"), got)

    def test_currency_helpers_are_pure(self) -> None:
        facts = _facts(*_ifrs_year({"Revenue": 1, "CostOfSales": 1}),
                       *_ifrs_year({"Revenue": 1}, unit="USD"),
                       ("DilutedEarningsLossPerShare", "USD/shares",
                        _fact(1, "2024-12-31", TSM_20F, "2025-04-17", start="2024-01-01")),
                       ("Revenue", "shares", _fact(1, "2024-12-31", TSM_20F, "2025-04-17")))
        counts = X.currency_counts(facts, list(X.IFRS_CONCEPT_MAP.items()))
        self.assertEqual(counts, {"TWD": 2, "USD": 1})
        self.assertEqual(X.choose_currency(counts, "auto"), "TWD")
        self.assertEqual(X.choose_currency(counts, "usd"), "USD")
        self.assertEqual(X.choose_currency({}, "auto"), "USD")
        self.assertEqual(X.accepted_units("TWD"), ("TWD", "TWD/shares"))

    def test_invalid_currency_code_is_rejected(self) -> None:
        for bad in ("dollars", "US", "U$D"):
            with self.subTest(code=bad), self.assertRaises(SystemExit) as ctx:
                X.choose_currency({"USD": 1}, bad)
            self.assertTrue(str(ctx.exception).startswith("[x]"))

    def test_domestic_us_gaap_issuer_still_resolves_to_usd(self) -> None:
        facts = _balance({ORIGINAL: ("2018-04-19", BS_ORIGINAL)})
        got, out, _ = _run_fetch({"facts": {"us-gaap": facts}})
        self.assertEqual({r["unit"] for r in got.values()}, {"USD"})
        self.assertTrue([ln for ln in out.splitlines()
                         if "taxonomia" in ln and "us-gaap" in ln])

    def test_payload_without_mapped_concepts_warns_instead_of_silence(self) -> None:
        payload = {"facts": {"dei": {"EntityCommonStockSharesOutstanding": {"units": {
            "shares": [{"end": "2024-12-31", "val": 1, "filed": "2025-01-01"}]}}}}}
        got, out, _ = _run_fetch(payload)
        self.assertEqual(got, {})
        aviso = [ln for ln in out.splitlines() if ln.startswith("[aviso]")]
        self.assertTrue(any("dei" in ln for ln in aviso))


class CliTests(unittest.TestCase):

    def _main(self, *args: str) -> dict:
        seen: dict = {}

        def fake_fetch(ticker, dest, ua, taxonomy="auto", currency="auto", unmapped=False):
            seen.update(ticker=ticker, taxonomy=taxonomy, currency=currency)
            return Path("x.csv")

        orig = X.fetch
        X.fetch = fake_fetch
        try:
            with contextlib.redirect_stdout(io.StringIO()), \
                    contextlib.redirect_stderr(io.StringIO()):
                X.main(["xbrl_fetch.py", *args, "--ua", "Test test@example.com"])
        finally:
            X.fetch = orig
        return seen

    def test_defaults_are_auto(self) -> None:
        self.assertEqual(self._main("TSM"),
                         {"ticker": "TSM", "taxonomy": "auto", "currency": "auto"})

    def test_overrides_reach_fetch(self) -> None:
        seen = self._main("TSM", "--taxonomy", "ifrs-full", "--currency", "TWD")
        self.assertEqual((seen["taxonomy"], seen["currency"]), ("ifrs-full", "TWD"))

    def test_unknown_taxonomy_is_an_argparse_error(self) -> None:
        with self.assertRaises(SystemExit):
            self._main("TSM", "--taxonomy", "mx-ifrs")


if __name__ == "__main__":
    unittest.main()
