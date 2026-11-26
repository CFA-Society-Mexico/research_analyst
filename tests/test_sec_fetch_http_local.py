"""sec_fetch contra un servidor HTTP REAL en 127.0.0.1 (issue #48).

Las demas pruebas de sec_fetch sustituyen ``urlopen``. Aqui corre la ruta
completa de ``urllib`` -conexion, cabeceras, cuerpo de la respuesta de error,
bytes binarios- contra un ``http.server`` local que imita a EDGAR: indice de
tickers, submissions, ``index.json`` de cada 6-K y sus documentos. Ejercita lo
que la prueba en vivo del 2026-11-25 no toco: reintentos con backoff, corte de
la corrida y exhibits en PDF. Sin red externa y sin dormir (``_sleep`` se
sustituye). Stdlib unittest.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import sec_fetch as S  # noqa: E402

UA = "Prueba Local prueba@example.com"
CIK = 1046179
ANNUAL = "0001046179-25-000010"
SIXK_A = "0001046179-25-000020"
SIXK_B = "0001046179-25-000021"
PDF_BYTES = b"%PDF-1.7\n\x00\x01\x02\xff\xfe binario \r\n%%EOF"

Answer = tuple[int, bytes]


class _Edgar:
    """Servidor local: cada ruta responde su lista de (status, cuerpo) en orden
    y repite la ultima. Guarda cada solicitud con su User-Agent."""

    def __init__(self, routes: dict[str, list[Answer]]) -> None:
        self.routes = {path: list(answers) for path, answers in routes.items()}
        self.requests: list[tuple[str, str]] = []
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:  # noqa: N802 - nombre del protocolo
                owner.requests.append((self.path, self.headers.get("User-Agent", "")))
                answers = owner.routes.get(self.path)
                if answers is None:
                    status, body = 404, b"ruta no servida por la prueba"
                else:
                    status, body = answers.pop(0) if len(answers) > 1 else answers[0]
                self.send_response(status)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, *args: object) -> None:
                return None  # sin ruido en la consola de la suite

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}"
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)

    def __enter__(self) -> "_Edgar":
        self.thread.start()
        return self

    def __exit__(self, *exc: object) -> None:
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=5)

    def hits(self, fragment: str) -> int:
        return sum(1 for path, _ in self.requests if fragment in path)


def _doc(accession: str, name: str) -> str:
    return f"/Archives/edgar/data/{CIK}/{accession.replace('-', '')}/{name}"


def _routes(**overrides: list[Answer]) -> dict[str, list[Answer]]:
    """EDGAR minimo de una emisora extranjera: un 20-F y dos 6-K del mismo periodo."""
    submissions = {"filings": {"files": [], "recent": {
        "form": ["20-F", "6-K", "6-K"],
        "filingDate": ["2025-04-17", "2025-04-17", "2025-05-10"],
        "reportDate": ["2024-12-31", "2025-03-31", "2025-03-31"],
        "accessionNumber": [ANNUAL, SIXK_A, SIXK_B],
        "primaryDocument": ["tsm-20f.htm", "tsm-6k.htm", "tsm-revenue.htm"]}}}
    index_a = {"directory": {"item": [
        {"name": "tsm-6k.htm"}, {"name": "ex99-1.htm"}, {"name": "reporte-1t25.pdf"},
        {"name": "R1.htm"}, {"name": f"{SIXK_A}-index.html"}, {"name": "logo.jpg"}]}}
    index_b = {"directory": {"item": [{"name": "tsm-revenue.htm"}]}}
    routes: dict[str, list[Answer]] = {
        "/files/company_tickers.json": [(200, json.dumps(
            {"0": {"cik_str": CIK, "ticker": "TSM", "title": "TSMC"}}).encode())],
        f"/submissions/CIK{CIK:010d}.json": [(200, json.dumps(submissions).encode())],
        _doc(ANNUAL, "tsm-20f.htm"): [(200, b"<html>20-F FY2024</html>")],
        _doc(SIXK_A, "index.json"): [(200, json.dumps(index_a).encode())],
        _doc(SIXK_A, "tsm-6k.htm"): [(200, b"<html>6-K portada</html>")],
        _doc(SIXK_A, "ex99-1.htm"): [(200, b"<html>press release con guidance</html>")],
        _doc(SIXK_A, "reporte-1t25.pdf"): [(200, PDF_BYTES)],
        _doc(SIXK_B, "index.json"): [(200, json.dumps(index_b).encode())],
        _doc(SIXK_B, "tsm-revenue.htm"): [(200, b"<html>ventas mensuales</html>")],
    }
    routes.update(overrides)
    return routes


class _Case(unittest.TestCase):

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dest = Path(tmp.name) / "filings"
        self.waits: list[float] = []

    def _run(self, routes: dict[str, list[Answer]]) -> tuple[_Edgar, int, str]:
        """Corre el CLI completo contra el servidor local: (servidor, exit, consola)."""
        out = io.StringIO()
        with _Edgar(routes) as edgar, contextlib.ExitStack() as stack:
            for name, url in (
                    ("SEC_TICKERS_URL", "/files/company_tickers.json"),
                    ("SEC_SUBMISSIONS_URL", "/submissions/{name}"),
                    ("SEC_ARCHIVES_URL", "/Archives/edgar/data/{cik}/{accn}/{doc}"),
                    ("SEC_FOLDER_INDEX_URL", "/Archives/edgar/data/{cik}/{accn}/index.json")):
                stack.enter_context(mock.patch.object(S, name, edgar.base + url))
            stack.enter_context(mock.patch.object(S, "_sleep", self.waits.append))
            stack.enter_context(mock.patch.object(S, "_proven_hosts", set()))
            # Un proxy del sistema no debe interceptar la conexion local.
            stack.enter_context(mock.patch.dict(
                os.environ, {"NO_PROXY": "127.0.0.1", "no_proxy": "127.0.0.1"}))
            stack.enter_context(contextlib.redirect_stdout(out))
            try:
                code = S.main(["sec_fetch.py", "TSM", "--dest", str(self.dest),
                               "--ua", UA, "--throttle", "0"])
            except SystemExit as exc:
                out.write(str(exc.code))
                code = 1
        return edgar, code, out.getvalue()

    def _backoff(self) -> list[float]:
        return [w for w in self.waits if w > 0]


class FullRunTests(_Case):

    def test_foreign_issuer_run_downloads_filings_and_pdf_exhibit_byte_for_byte(self) -> None:
        edgar, code, out = self._run(_routes())
        self.assertEqual(code, 0, out)
        self.assertIn("emisora extranjera", out)
        names = sorted(p.name for p in self.dest.iterdir() if p.suffix != ".csv")
        self.assertEqual(names, [
            "TSM_20-F_2024-12-31_2025-04-17_000010.htm",
            "TSM_6-K_2025-03-31_2025-04-17_000020.htm",
            "TSM_6-K_2025-03-31_2025-04-17_000020_ex_ex99-1.htm",
            "TSM_6-K_2025-03-31_2025-04-17_000020_ex_reporte-1t25.pdf",
            "TSM_6-K_2025-03-31_2025-05-10_000021.htm",
        ])
        pdf = self.dest / "TSM_6-K_2025-03-31_2025-04-17_000020_ex_reporte-1t25.pdf"
        self.assertEqual(pdf.read_bytes(), PDF_BYTES)              # binario intacto
        rows = S.read_manifest(self.dest / "TSM_filings_manifest.csv")
        self.assertEqual(len(rows), 5)
        self.assertEqual(len({r["file"] for r in rows}), 5)
        self.assertEqual(sorted(r["form"] for r in rows),
                         ["20-F", "6-K", "6-K", "6-K-EX", "6-K-EX"])
        self.assertEqual(self._backoff(), [])                      # sin errores, sin esperas
        self.assertEqual(edgar.hits("R1.htm") + edgar.hits("logo.jpg"), 0)

    def test_every_request_carries_the_user_agent(self) -> None:
        edgar, _, _ = self._run(_routes())
        self.assertGreaterEqual(len(edgar.requests), 9)
        self.assertEqual({agent for _, agent in edgar.requests}, {UA})

    def test_annual_report_folder_is_not_listed_for_exhibits(self) -> None:
        edgar, _, _ = self._run(_routes())
        self.assertEqual(edgar.hits(_doc(ANNUAL, "index.json")), 0)
        self.assertEqual(edgar.hits(_doc(SIXK_A, "index.json")), 1)


class RetryTests(_Case):

    def test_transient_503_is_retried_and_the_document_arrives(self) -> None:
        path = _doc(SIXK_A, "tsm-6k.htm")
        edgar, code, out = self._run(_routes(**{
            path: [(503, b"Service Unavailable"), (503, b"again"), (200, b"<html>ok</html>")]}))
        self.assertEqual(code, 0, out)
        self.assertEqual(edgar.hits(path), 3)
        self.assertEqual(self._backoff(), [2, 4])
        self.assertEqual((self.dest / "TSM_6-K_2025-03-31_2025-04-17_000020.htm").read_bytes(),
                         b"<html>ok</html>")
        self.assertIn("HTTP 503: reintento 1 en 2 s", out)

    def test_403_after_the_host_already_answered_is_a_rate_limit_not_a_block(self) -> None:
        path = _doc(SIXK_A, "reporte-1t25.pdf")
        edgar, code, out = self._run(_routes(**{
            path: [(403, b"rate limited"), (200, PDF_BYTES)]}))
        self.assertEqual(code, 0, out)
        self.assertEqual(edgar.hits(path), 2)
        self.assertEqual(self._backoff(), [2])
        self.assertNotIn("Sin acceso de red", out)

    def test_403_before_any_success_is_a_blocked_network_after_one_short_retry(self) -> None:
        edgar, code, out = self._run(_routes(**{
            "/files/company_tickers.json": [(403, b"Forbidden")]}))
        self.assertEqual(code, 1)
        self.assertIn("Sin acceso de red a SEC (HTTP 403)", out)
        self.assertEqual(edgar.hits("company_tickers"), 2)          # 1 intento + 1 reintento
        self.assertEqual(self._backoff(), [2])
        self.assertFalse(self.dest.exists() and any(self.dest.iterdir()))

    def test_404_is_final_recorded_and_the_run_continues(self) -> None:
        path = _doc(SIXK_A, "ex99-1.htm")
        edgar, code, out = self._run(_routes(**{path: [(404, b"Not Found")]}))
        self.assertEqual(code, 1)                                   # hubo una falla
        self.assertEqual(edgar.hits(path), 1)                       # un 404 no se reintenta
        self.assertEqual(self._backoff(), [])
        self.assertIn("HTTP 404", out)
        names = {p.name for p in self.dest.iterdir()}
        self.assertNotIn("TSM_6-K_2025-03-31_2025-04-17_000020_ex_ex99-1.htm", names)
        self.assertIn("TSM_6-K_2025-03-31_2025-05-10_000021.htm", names)   # siguio
        self.assertEqual(len(S.read_manifest(self.dest / "TSM_filings_manifest.csv")), 4)

    def test_three_documents_exhausting_retries_stop_the_run_with_the_manifest_saved(self) -> None:
        down = [(503, b"down")]
        edgar, code, out = self._run(_routes(**{
            _doc(SIXK_A, "tsm-6k.htm"): down, _doc(SIXK_A, "ex99-1.htm"): down,
            _doc(SIXK_A, "reporte-1t25.pdf"): down}))
        self.assertEqual(code, 1)
        self.assertIn("corrida detenida", out)
        for name in ("tsm-6k.htm", "ex99-1.htm", "reporte-1t25.pdf"):
            self.assertEqual(edgar.hits(_doc(SIXK_A, name)), 7, name)   # 1 + 6 reintentos
        self.assertEqual(self._backoff(), [2, 4, 8, 16, 32, 64] * 3)
        self.assertEqual(edgar.hits(_doc(SIXK_B, "tsm-revenue.htm")), 0)    # no siguio
        rows = S.read_manifest(self.dest / "TSM_filings_manifest.csv")
        self.assertEqual([r["form"] for r in rows], ["20-F"])       # lo ya bajado, a salvo
        self.assertFalse(list(self.dest.glob("*.part")))            # sin archivos truncados


if __name__ == "__main__":
    unittest.main()
