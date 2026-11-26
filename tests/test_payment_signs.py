"""Pruebas del check C11: signo de los pagos en la capa de captura (issue #49).

Convencion del contrato: en ``canonical_*.csv`` un pago bruto (capex,
dividendos, recompras, amortizacion de deuda...) se guarda como MAGNITUD
POSITIVA, tal como lo reporta XBRL, y la formula del modelo lo resta. Un pago
en negativo es una captura con la convencion contraria: mezclada con el resto,
infla o hunde el FCF sin que el balance lo note. Stdlib unittest.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import model_check as M  # noqa: E402
import xbrl_fetch as X  # noqa: E402
import xlsx_builder as B  # noqa: E402

HEADER = "line_item,period,value,source_doc,source_ref,tag\n"


class PaymentSignTests(unittest.TestCase):

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.inputs = Path(tmp.name) / "inputs"
        self.inputs.mkdir()

    def _write(self, name: str, *rows: str) -> None:
        (self.inputs / name).write_text(HEADER + "".join(r + "\n" for r in rows),
                                        encoding="utf-8")

    def _c11(self, **kwargs: bool) -> B.Finding:
        (finding,) = M.audit_payment_signs(str(self.inputs), **kwargs)
        return finding

    def test_payments_stored_as_positive_magnitudes_pass(self) -> None:
        self._write("canonical_annual.csv",
                    "cf_capex,FY2024,9447,10-K,p.31,observado",
                    "cf_dividends,FY2024,15234,10-K,p.31,observado",
                    "cf_cff,FY2024,-121983,10-K,p.31,observado")
        finding = self._c11()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("2 pagos", finding.detail)

    def test_a_negative_gross_payment_fails_and_names_file_line_and_period(self) -> None:
        self._write("canonical_annual.csv",
                    "cf_capex,FY2023,10959,10-K,p.31,observado",
                    "cf_capex,FY2024,-9447,10-K,p.31,observado")
        self._write("canonical_quarterly.csv",
                    "cf_dividends,1Q2025,-3856,10-Q,p.8,observado")
        finding = self._c11()
        self.assertTrue(finding.blocking)
        self.assertIn("canonical_annual.csv: cf_capex FY2024", finding.detail)
        self.assertIn("canonical_quarterly.csv: cf_dividends 1Q2025", finding.detail)
        self.assertNotIn("FY2023", finding.detail)

    def test_subtotals_and_net_lines_keep_their_natural_sign(self) -> None:
        # CFI, CFF, cambio neto y las lineas NETAS pueden ser negativas: no son pagos brutos.
        self._write("canonical_annual.csv",
                    "cf_cfi,FY2024,-2935,10-K,p.31,observado",
                    "cf_cff,FY2024,-121983,10-K,p.31,observado",
                    "cf_net_change,FY2024,-794,10-K,p.31,observado",
                    "cf_other_investing,FY2024,-1308,10-K,p.31,observado",
                    "cf_commercial_paper_net,FY2024,-3960,10-K,p.31,observado",
                    "cf_taxes_paid,FY2024,-12,20-F,p.F-9,observado")
        finding = self._c11()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("n/a", finding.detail)

    def test_zero_is_a_valid_payment(self) -> None:
        self._write("canonical_annual.csv", "cf_debt_repaid,FY2024,0,10-K,p.31,observado")
        self.assertTrue(self._c11().ok)

    def test_guidance_rows_are_not_capture_and_are_ignored(self) -> None:
        self._write("canonical_annual.csv", "cf_capex,FY2026,-40000,call 3T,p.4,guidance")
        self.assertTrue(self._c11().ok)

    def test_no_capture_layer_is_not_applicable_here_c9_reports_it(self) -> None:
        self.inputs.rmdir()
        finding = self._c11()
        self.assertTrue(finding.ok)
        self.assertIn("n/a", finding.detail)

    def test_unreadable_capture_fails_instead_of_passing_unchecked(self) -> None:
        self._write("canonical_annual.csv", "cf_capex,FY2024,no-es-numero,10-K,p.31,observado")
        self.assertTrue(self._c11().blocking)


class ContractTests(unittest.TestCase):

    def test_every_reserved_payment_line_is_an_outflow_concept_of_xbrl_fetch(self) -> None:
        # La lista de C11 y la marca sign = salida(+) del CSV de XBRL no pueden divergir.
        outflow_canons = {canon for mapping in (X.CONCEPT_MAP, X.LEGACY_CONCEPT_MAP,
                                                X.IFRS_CONCEPT_MAP)
                          for concept, canon in mapping.items()
                          if concept in X.OUTFLOW_POSITIVE_CONCEPTS}
        self.assertLessEqual(M.PAYMENT_LINE_ITEMS, outflow_canons)

    def test_net_lines_are_deliberately_outside_the_reserved_set(self) -> None:
        for net in ("cf_other_investing", "cf_taxes_paid", "cf_acquisitions",
                    "cf_commercial_paper_net", "cf_cfi", "cf_cff"):
            self.assertNotIn(net, M.PAYMENT_LINE_ITEMS)

    def test_model_check_reports_c11_next_to_c9(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            model = Path(tmp) / "demo_2026-01-01_v1.xlsx"
            with contextlib.redirect_stdout(io.StringIO()):
                B._demo(str(model))
            # El demo escribe su propia capa de captura junto al modelo: se le
            # agrega un pago capturado con el signo del estado de flujo.
            capture = Path(tmp) / "inputs" / "canonical_quarterly.csv"
            header = capture.read_text(encoding="utf-8").splitlines()[0].split(",")
            row = dict.fromkeys(header, "")
            row.update(line_item="cf_capex", period="1Q2024", value="-5", tag="observado")
            with capture.open("a", encoding="utf-8", newline="") as fh:
                fh.write(",".join(row[c] for c in header) + "\n")
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                code = M.main(["model_check.py", str(model)])
        self.assertEqual(code, 1)
        self.assertIn("[x] C11 pagos en positivo en la captura", out.getvalue())
        self.assertIn("cf_capex 1Q2024", out.getvalue())


if __name__ == "__main__":
    unittest.main()
