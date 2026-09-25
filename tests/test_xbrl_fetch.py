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


class FetchTests(unittest.TestCase):

    def _fetch(self, facts: dict) -> dict[tuple[str, str], dict]:
        payload = {"facts": {"us-gaap": facts}}
        orig = (X.resolve_cik, X._get_json)
        X.resolve_cik = lambda ticker, ua: "0000000001"
        X._get_json = lambda url, ua: payload
        try:
            with tempfile.TemporaryDirectory() as tmp, \
                    contextlib.redirect_stdout(io.StringIO()):
                out = X.fetch("TEST", Path(tmp), "Test test@example.com")
                with out.open(encoding="utf-8") as fh:
                    return {(r["canon"], r["period"]): r for r in csv.DictReader(fh)}
        finally:
            X.resolve_cik, X._get_json = orig

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


if __name__ == "__main__":
    unittest.main()
