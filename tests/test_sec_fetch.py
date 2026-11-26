"""Pruebas de tools/sec_fetch.py (stdlib unittest, red simulada).

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import csv
import io
import json
import sys
import tempfile
import unittest
import urllib.error
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import sec_fetch as S  # noqa: E402

UA = "Test test@example.com"
CIK10 = "0000320193"
BACKOFF = [2, 4, 8, 16, 32, 64]
EMPTY_INDEX = b'{"directory": {"item": []}}'


class _Resp:
    def __init__(self, body: bytes) -> None:
        self._body = body

    def read(self) -> bytes:
        return self._body

    def __enter__(self) -> "_Resp":
        return self

    def __exit__(self, *exc: object) -> None:
        return None


def _http_error(url: str, code: int) -> urllib.error.HTTPError:
    return urllib.error.HTTPError(url, code, "err", {}, io.BytesIO(b""))


def _fake_net(routes: dict[str, object]):
    """urlopen falso: la primera ruta cuyo fragmento aparece en la URL decide.
    Valor int = codigo HTTP de error; bytes = cuerpo; excepcion = se lanza;
    lista = respuestas en orden. `.calls` guarda cada URL pedida."""
    def urlopen(req, timeout=0):
        url = req.full_url
        urlopen.calls.append(url)
        for frag, answer in routes.items():
            if frag in url:
                if isinstance(answer, list):
                    answer = answer.pop(0)
                if isinstance(answer, int):
                    raise _http_error(url, answer)
                if isinstance(answer, BaseException):
                    raise answer
                return _Resp(answer)
        raise AssertionError(f"URL no esperada en la prueba: {url}")
    urlopen.calls = []
    return urlopen


def _batch(rows: list[S.FilingRow]) -> dict[str, list[str]]:
    """Lote de submissions de EDGAR equivalente a `rows`."""
    return {"form": [r.form for r in rows],
            "filingDate": [r.filed for r in rows],
            "reportDate": [r.period for r in rows],
            "accessionNumber": [r.accession for r in rows],
            "primaryDocument": [r.primary_doc for r in rows]}


def _index(*names: str) -> bytes:
    return json.dumps({"directory": {"item": [{"name": n} for n in names]}}).encode()


def _url(accession: str, doc: str) -> str:
    return S.SEC_ARCHIVES_URL.format(cik="320193", accn=accession.replace("-", ""), doc=doc)


class _OfflineCase(unittest.TestCase):
    """Sin red, sin esperas reales y con la red 'no probada' al empezar."""

    def setUp(self) -> None:
        self.sleeps: list[float] = []
        mock.patch.object(S, "_sleep", self.sleeps.append).start()
        mock.patch.object(S.time, "sleep",
                          side_effect=AssertionError("espera real en una prueba")).start()
        mock.patch.object(S, "_proven_hosts", set()).start()
        self.addCleanup(mock.patch.stopall)


class HttpClassificationTests(_OfflineCase):

    URL = "https://www.sec.gov/x.json"

    def _get(self, routes: dict[str, object], url: str = URL):
        net = _fake_net(routes)
        buf = io.StringIO()
        with mock.patch("urllib.request.urlopen", net), contextlib.redirect_stdout(buf):
            try:
                return S._get_json(url, UA)
            finally:
                self.calls = net.calls
                self.out = buf.getvalue()

    def _prove_network(self) -> None:
        self._get({"ok.json": b"{}"}, "https://www.sec.gov/ok.json")
        self.sleeps.clear()

    def test_404_is_not_reported_as_blocked_network(self) -> None:
        with self.assertRaises(S.SecHttpError) as ctx:
            self._get({"x.json": 404})
        self.assertNotIsInstance(ctx.exception, S.SecRetriesExhausted)
        self.assertEqual(self.sleeps, [])

    def test_407_is_always_an_immediate_blocked_exit(self) -> None:
        self._prove_network()
        with self.assertRaises(SystemExit) as ctx:
            self._get({"x.json": 407})
        self.assertIn("Sin acceso de red", str(ctx.exception))
        self.assertEqual((len(self.calls), self.sleeps), (1, []))

    def test_403_before_any_success_fails_fast_as_blocked(self) -> None:
        with self.assertRaises(SystemExit) as ctx:
            self._get({"x.json": 403})
        self.assertIn("Sin acceso de red", str(ctx.exception))
        self.assertLessEqual(len(self.calls), 2)
        self.assertLessEqual(sum(self.sleeps), 2)

    def test_urlerror_before_any_success_fails_fast_as_blocked(self) -> None:
        with self.assertRaises(SystemExit) as ctx:
            self._get({"x.json": urllib.error.URLError("proxy")})
        self.assertIn("Sin acceso de red", str(ctx.exception))
        self.assertLessEqual(len(self.calls), 2)
        self.assertLessEqual(sum(self.sleeps), 2)

    def test_403_after_a_success_is_retried_with_backoff(self) -> None:
        self._prove_network()
        self.assertEqual(self._get({"x.json": [403, 403, b'{"ok": 1}']}), {"ok": 1})
        self.assertEqual(self.sleeps, [2, 4])

    def test_urlerror_after_a_success_is_retried_with_backoff(self) -> None:
        self._prove_network()
        down = urllib.error.URLError("reset")
        self.assertEqual(self._get({"x.json": [down, down, down, b'{"ok": 1}']}), {"ok": 1})
        self.assertEqual(self.sleeps, [2, 4, 8])

    def test_localized_os_error_text_reaches_the_console_as_ascii(self) -> None:
        self._prove_network()
        down = urllib.error.URLError("conexión denegada")
        with self.assertRaises(S.SecRetriesExhausted) as ctx:
            self._get({"x.json": down})
        self.assertIn("reintento", self.out)
        self.assertTrue(self.out.isascii())
        self.assertTrue(str(ctx.exception).isascii())

    def test_success_on_one_host_does_not_prove_another(self) -> None:
        self._prove_network()  # www.sec.gov
        with self.assertRaises(SystemExit):
            self._get({"y.json": 403}, "https://data.sec.gov/y.json")
        self.assertLessEqual(len(self.calls), 2)

    def test_429_and_5xx_are_retried_up_to_six_times(self) -> None:
        answers: list[object] = [429, 500, 502, 503, 504, 429, b'{"ok": 1}']
        self.assertEqual(self._get({"x.json": answers}), {"ok": 1})
        self.assertEqual(self.sleeps, BACKOFF)

    def test_timeout_and_connection_errors_are_retried(self) -> None:
        answers: list[object] = [TimeoutError("t"), ConnectionResetError("c"), b'{"ok": 1}']
        self.assertEqual(self._get({"x.json": answers}), {"ok": 1})
        self.assertEqual(self.sleeps, [2, 4])

    def test_exhausted_retries_raise_sec_http_error(self) -> None:
        with self.assertRaises(S.SecRetriesExhausted):
            self._get({"x.json": 503})
        self.assertEqual((len(self.calls), self.sleeps), (7, BACKOFF))

    def test_403_exhausted_after_a_success_is_not_a_system_exit(self) -> None:
        self._prove_network()
        with self.assertRaises(S.SecHttpError):
            self._get({"x.json": 403})
        self.assertEqual(self.sleeps, BACKOFF)


class FetchTests(_OfflineCase):

    def setUp(self) -> None:
        super().setUp()
        mock.patch.object(S, "resolve_cik", lambda t, ua: CIK10).start()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dest = Path(self.tmp.name)
        self.out = ""

    def _run(self, rows: list[S.FilingRow], routes: dict[str, object],
             forms: object = ("10-K", "10-Q", "8-K"), dest: Path | None = None,
             **kw: object) -> tuple[Path, list[str]]:
        self.net = _fake_net(routes)
        buf = io.StringIO()
        with mock.patch.object(S, "load_submissions",
                               lambda cik, ua, throttle=0.0: [_batch(rows)]), \
                mock.patch("urllib.request.urlopen", self.net), \
                contextlib.redirect_stdout(buf):
            try:
                return S.fetch("AAPL", dest or self.dest, forms, "1994-01-01", UA, **kw)
            finally:
                self.out = buf.getvalue()

    def _manifest(self, dest: Path | None = None) -> list[dict[str, str]]:
        path = (dest or self.dest) / "AAPL_filings_manifest.csv"
        with path.open(encoding="utf-8") as fh:
            return list(csv.DictReader(fh))

    def _seed_manifest(self, rows: list[dict[str, str]]) -> None:
        path = self.dest / "AAPL_filings_manifest.csv"
        with path.open("w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=S.MANIFEST_FIELDS)
            writer.writeheader()
            writer.writerows(rows)

    def _downloads(self) -> set[str]:
        return {p.name for p in self.dest.iterdir() if p.suffix != ".csv"}

    # ---- naming (issue 32) -------------------------------------------------

    def test_raw_name_carries_period_filed_and_accession_tail(self) -> None:
        rows = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", "0000320193-24-000123", "k.htm"),
                S.FilingRow("8-K", "2025-05-02", "2025-05-01", "0000320193-25-000010", "a8k.htm")]
        self._run(rows, {"index.json": _index("a8k.htm", "ex99.htm"),
                         "k.htm": b"k", "a8k.htm": b"a", "ex99.htm": b"pr"})
        self.assertEqual(self._downloads(), {
            "AAPL_10-K_2024-09-28_2024-11-01_000123.htm",
            "AAPL_8-K_2025-05-01_2025-05-02_000010.htm",
            "AAPL_8-K_2025-05-01_2025-05-02_000010_ex_ex99.htm"})

    def test_two_6k_of_the_same_period_get_two_files_and_two_rows(self) -> None:
        rows = [S.FilingRow("6-K", "2025-01-16", "2025-01-31", "0001046179-25-000007", "a6k.htm"),
                S.FilingRow("6-K", "2025-01-16", "2025-01-31", "0001046179-25-000008", "b6k.htm")]
        routes = {"index.json": EMPTY_INDEX, "a6k.htm": b"primero", "b6k.htm": b"segundo"}
        self._run(rows, dict(routes), forms=("20-F", "6-K"))
        files = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual(files, {
            "0001046179-25-000007": "AAPL_6-K_2025-01-31_2025-01-16_000007.htm",
            "0001046179-25-000008": "AAPL_6-K_2025-01-31_2025-01-16_000008.htm"})
        self.assertEqual((self.dest / files["0001046179-25-000008"]).read_bytes(), b"segundo")
        # Independiente del orden en que EDGAR liste los filings.
        with tempfile.TemporaryDirectory() as other:
            self._run(rows[::-1], dict(routes), forms=("20-F", "6-K"), dest=Path(other))
            reversed_files = {r["accession"]: r["file"] for r in self._manifest(Path(other))}
        self.assertEqual(reversed_files, files)

    def test_same_tail_from_two_accessions_falls_back_to_full_accession(self) -> None:
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0001193125-25-000010", "b.htm"),
                S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a.htm")]
        self._run(rows, {"index.json": EMPTY_INDEX, "a.htm": b"primero", "b.htm": b"segundo"})
        files = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual(files, {
            "0000320193-25-000010": "AAPL_8-K_2025-05-01_2025-05-01_000010.htm",
            "0001193125-25-000010": "AAPL_8-K_2025-05-01_2025-05-01_0001193125-25-000010.htm"})
        self.assertEqual((self.dest / files["0001193125-25-000010"]).read_bytes(), b"segundo")

    def test_short_name_stays_with_the_accession_that_already_owns_it(self) -> None:
        late = S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0001193125-25-000010", "b.htm")
        early = S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a.htm")
        self._run([late], {"index.json": EMPTY_INDEX, "b.htm": b"segundo"})
        self._run([late, early], {"index.json": EMPTY_INDEX, "a.htm": b"primero"})
        files = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual(files["0001193125-25-000010"],
                         "AAPL_8-K_2025-05-01_2025-05-01_000010.htm")
        self.assertEqual((self.dest / files["0000320193-25-000010"]).read_bytes(), b"primero")
        self.assertEqual(len(set(files.values())), 2)

    def test_two_documents_never_share_one_file(self) -> None:
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a8k.htm")]
        _, failures = self._run(rows, {"index.json": _index("a8k.htm", "ex99.htm", "EX99.htm"),
                                       "a8k.htm": b"a", "ex99.htm": b"uno", "EX99.htm": b"dos"})
        self.assertEqual(len(failures), 1)
        self.assertEqual(len({r["file"].lower() for r in self._manifest()}), len(self._manifest()))

    def test_lost_manifest_does_not_misattribute_existing_files(self) -> None:
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a.htm"),
                S.FilingRow("8-K", "2025-05-02", "2025-05-01", "0000320193-25-000011", "b.htm")]
        self._run(rows, {"index.json": EMPTY_INDEX, "a.htm": b"primero", "b.htm": b"segundo"})
        (self.dest / "AAPL_filings_manifest.csv").unlink()
        self._run(rows[::-1], {"index.json": EMPTY_INDEX})  # cualquier descarga falla la prueba
        files = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual((self.dest / files["0000320193-25-000010"]).read_bytes(), b"primero")
        self.assertEqual((self.dest / files["0000320193-25-000011"]).read_bytes(), b"segundo")

    def test_amendment_form_does_not_create_a_folder(self) -> None:
        rows = [S.FilingRow("10-K/A", "2020-01-10", "2019-09-28", "0000320193-20-000001", "ka.htm")]
        self._run(rows, {"ka.htm": b"x"}, forms=("10-K/A",))
        self.assertTrue((self.dest / "AAPL_10-K-A_2019-09-28_2020-01-10_000001.htm").exists())
        self.assertEqual(self._manifest()[0]["form"], "10-K/A")

    # ---- legacy workspaces -------------------------------------------------

    def test_unambiguous_legacy_name_is_reused_without_download(self) -> None:
        accn = "0000320193-24-000123"
        (self.dest / "AAPL_10-K_2024-09-28.htm").write_bytes(b"viejo")
        self._seed_manifest([{"file": "AAPL_10-K_2024-09-28.htm", "form": "10-K",
                              "filed": "2024-11-01", "period": "2024-09-28",
                              "accession": accn, "source": _url(accn, "k.htm")}])
        rows = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", accn, "k.htm")]
        _, failures = self._run(rows, {})  # cualquier descarga falla la prueba
        self.assertEqual(failures, [])
        self.assertIn("[ok] ya existe: AAPL_10-K_2024-09-28.htm", self.out)
        self.assertEqual([r["file"] for r in self._manifest()], ["AAPL_10-K_2024-09-28.htm"])
        self.assertEqual(self._downloads(), {"AAPL_10-K_2024-09-28.htm"})
        self.assertNotIn("[aviso]", self.out)

    def test_legacy_row_whose_file_is_gone_downloads_the_new_name(self) -> None:
        accn = "0000320193-24-000123"
        self._seed_manifest([{"file": "AAPL_10-K_2024-09-28.htm", "form": "10-K",
                              "filed": "2024-11-01", "period": "2024-09-28",
                              "accession": accn, "source": _url(accn, "k.htm")}])
        rows = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", accn, "k.htm")]
        self._run(rows, {"k.htm": b"k"})
        self.assertEqual([r["file"] for r in self._manifest()],
                         ["AAPL_10-K_2024-09-28_2024-11-01_000123.htm"])

    def test_ambiguous_legacy_name_is_refetched_under_the_new_name(self) -> None:
        # Manifest de la version vieja: dos accessions -> el mismo archivo.
        first, second = "0000320193-25-000010", "0000320193-25-000011"
        legacy = "AAPL_8-K_2025-05-01.htm"
        (self.dest / legacy).write_bytes(b"primero")
        self._seed_manifest([
            {"file": legacy, "form": "8-K", "filed": "2025-05-01", "period": "2025-05-01",
             "accession": first, "source": _url(first, "a.htm")},
            {"file": legacy, "form": "8-K", "filed": "2025-05-02", "period": "2025-05-01",
             "accession": second, "source": _url(second, "b.htm")}])
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", first, "a.htm"),
                S.FilingRow("8-K", "2025-05-02", "2025-05-01", second, "b.htm")]
        self._run(rows, {"index.json": EMPTY_INDEX, "a.htm": b"primero", "b.htm": b"segundo"})
        by_acc = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual(by_acc, {first: "AAPL_8-K_2025-05-01_2025-05-01_000010.htm",
                                  second: "AAPL_8-K_2025-05-01_2025-05-02_000011.htm"})
        self.assertEqual((self.dest / by_acc[second]).read_bytes(), b"segundo")
        self.assertEqual((self.dest / legacy).read_bytes(), b"primero")  # nada se borra
        avisos = [line for line in self.out.splitlines() if line.startswith("[aviso]")]
        self.assertEqual(len(avisos), 1)
        self.assertIn("2 ", avisos[0])
        # La segunda corrida ya no tiene nada ambiguo ni nada que bajar.
        self._run(rows, {"index.json": EMPTY_INDEX})
        self.assertNotIn("[aviso]", self.out)

    # ---- manifest ----------------------------------------------------------

    def test_manifest_is_merged_across_runs(self) -> None:
        first = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", "0000320193-24-000123", "k.htm")]
        second = [S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm")]
        self._run(first, {"k.htm": b"k"})
        self._run(second, {"q.htm": b"q"})
        self.assertEqual({r["form"] for r in self._manifest()}, {"10-K", "10-Q"})

    def _quarters(self, count: int) -> list[S.FilingRow]:
        return [S.FilingRow("10-Q", f"2025-0{i}-15", f"2025-0{i}-01",
                            f"0000320193-25-00000{i}", f"d{i}.htm")
                for i in range(1, count + 1)]

    def test_manifest_is_persisted_while_the_run_is_still_going(self) -> None:
        real_download = S._download
        rows_on_disk: list[int] = []

        def killed_on_third(url: str, dest: Path, user_agent: str) -> None:
            if "d3.htm" in url:
                rows_on_disk.append(len(self._manifest()))
                raise KeyboardInterrupt
            real_download(url, dest, user_agent)

        with mock.patch.object(S, "_download", killed_on_third):
            with self.assertRaises(KeyboardInterrupt):
                self._run(self._quarters(4), {"d1.htm": b"1", "d2.htm": b"2"})
        self.assertEqual(rows_on_disk, [2])
        self.assertEqual(len(self._manifest()), 2)

    def test_manifest_is_on_disk_when_the_run_exits_midway(self) -> None:
        with self.assertRaises(SystemExit):
            self._run(self._quarters(4), {"d1.htm": b"1", "d2.htm": b"2", "d3.htm": 407})
        self.assertEqual([r["period"] for r in self._manifest()], ["2025-01-01", "2025-02-01"])

    # ---- failures and retries (issue 34) -----------------------------------

    def test_exhibit_index_404_does_not_abort_the_run(self) -> None:
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a8k.htm"),
                S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm")]
        _, failures = self._run(rows, {"index.json": 404, "a8k.htm": b"k", "q.htm": b"q"})
        self.assertEqual(failures, [])
        self.assertEqual(len(self._manifest()), 2)

    def test_failed_download_is_reported_and_others_kept(self) -> None:
        rows = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", "0000320193-24-000123", "k.htm"),
                S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm")]
        _, failures = self._run(rows, {"k.htm": 404, "q.htm": b"q"})
        self.assertEqual(failures, ["AAPL_10-K_2024-09-28_2024-11-01_000123.htm"])
        self.assertEqual([r["form"] for r in self._manifest()], ["10-Q"])
        self.assertFalse(any(p.suffix == ".part" for p in self.dest.iterdir()))

    def test_exhausted_retries_record_the_failure_and_the_run_continues(self) -> None:
        _, failures = self._run(self._quarters(3),
                                {"d1.htm": b"1", "d2.htm": 403, "d3.htm": b"3"})
        self.assertEqual(failures, ["AAPL_10-Q_2025-02-01_2025-02-15_000002.htm"])
        self.assertEqual([r["period"] for r in self._manifest()], ["2025-01-01", "2025-03-01"])
        self.assertEqual([s for s in self.sleeps if s >= 2], BACKOFF)

    def test_three_exhausted_documents_in_a_row_stop_the_run(self) -> None:
        routes = {"d1.htm": b"1", "d2.htm": 503, "d3.htm": 503, "d4.htm": 503, "d5.htm": b"5"}
        _, failures = self._run(self._quarters(5), routes)
        self.assertEqual(len(failures), 3)
        self.assertFalse(any("d5.htm" in url for url in self.net.calls))
        self.assertEqual([r["period"] for r in self._manifest()], ["2025-01-01"])
        stop = [line for line in self.out.splitlines()
                if line.startswith("[x]") and "detenida" in line]
        self.assertEqual(len(stop), 1)

    def test_a_success_resets_the_circuit_breaker(self) -> None:
        routes = {"d1.htm": 503, "d2.htm": 503, "d3.htm": b"3", "d4.htm": 503, "d5.htm": 503,
                  "d6.htm": b"6"}
        _, failures = self._run(self._quarters(6), routes)
        self.assertEqual(len(failures), 4)
        self.assertEqual([r["period"] for r in self._manifest()], ["2025-03-01", "2025-06-01"])

    # ---- foreign private issuers (issue 35) --------------------------------

    def _foreign_rows(self) -> list[S.FilingRow]:
        return [S.FilingRow("6-K", "2025-01-16", "2025-01-31", "0001046179-25-000001", "tsm6k.htm"),
                S.FilingRow("20-F", "2025-04-15", "2024-12-31", "0001046179-25-000004", "tsm20f.htm"),
                S.FilingRow("SC 13G", "2025-02-10", "2024-12-31", "0001046179-25-000002", "sc.htm")]

    def test_foreign_issuer_is_detected_when_forms_are_not_given(self) -> None:
        self._run(self._foreign_rows(), {"index.json": EMPTY_INDEX, "tsm6k.htm": b"6",
                                         "tsm20f.htm": b"20"}, forms=None)
        self.assertEqual([r["form"] for r in self._manifest()], ["6-K", "20-F"])
        detected = [line for line in self.out.splitlines() if "detectadas" in line]
        self.assertEqual(len(detected), 1, self.out)
        self.assertTrue(detected[0].startswith("[ok]"))
        self.assertIn("extranjera", detected[0])
        self.assertIn("20-F,6-K", detected[0])

    def test_domestic_issuer_detection_keeps_the_default_forms(self) -> None:
        rows = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", "0000320193-24-000123", "k.htm"),
                S.FilingRow("6-K", "2025-01-16", "2025-01-31", "0000320193-25-000001", "x.htm")]
        self._run(rows, {"k.htm": b"k"}, forms=None)
        self.assertEqual([r["form"] for r in self._manifest()], ["10-K"])
        detected = [line for line in self.out.splitlines() if "detectadas" in line]
        self.assertEqual(len(detected), 1, self.out)
        self.assertIn("10-K,10-Q,8-K", detected[0])

    def test_explicit_forms_win_over_detection(self) -> None:
        self._run(self._foreign_rows(), {"tsm20f.htm": b"20"}, forms=("20-F",))
        self.assertEqual([r["form"] for r in self._manifest()], ["20-F"])
        self.assertNotIn("detectadas", self.out)

    def test_6k_exhibits_are_fetched_including_pdf(self) -> None:
        accn = "0001046179-25-000001"
        index = _index("tsm6k.htm", "ex99-1.htm", "ex99-2.pdf", "logo.jpg", "R1.htm",
                       f"{accn}-index.htm", f"{accn}.txt", "Financial_Report.xlsx")
        self._run(self._foreign_rows()[:1], {"index.json": index, "tsm6k.htm": b"6",
                                             "ex99-1.htm": b"pr", "ex99-2.pdf": b"%PDF"},
                  forms=("6-K",))
        prefix = "AAPL_6-K_2025-01-31_2025-01-16_000001"
        self.assertEqual(self._downloads(), {f"{prefix}.htm", f"{prefix}_ex_ex99-1.htm",
                                             f"{prefix}_ex_ex99-2.pdf"})
        self.assertEqual([r["form"] for r in self._manifest()], ["6-K", "6-K-EX", "6-K-EX"])

    def test_submissions_index_is_fetched_once(self) -> None:
        body = json.dumps({"filings": {"recent": _batch(self._foreign_rows()),
                                       "files": []}}).encode()
        net = _fake_net({"submissions/CIK": body, "index.json": EMPTY_INDEX,
                         "tsm6k.htm": b"6", "tsm20f.htm": b"20"})
        with mock.patch("urllib.request.urlopen", net), \
                contextlib.redirect_stdout(io.StringIO()):
            S.fetch("TSM", self.dest, None, "1994-01-01", UA)
        self.assertEqual(sum("submissions" in url for url in net.calls), 1)
        self.assertTrue((self.dest / "TSM_20-F_2024-12-31_2025-04-15_000004.htm").exists())


class DetectFormsTests(unittest.TestCase):

    def _forms(self, *filed: str) -> tuple[str, ...]:
        forms, _label = S.detect_forms([{"form": list(filed)}])
        return forms

    def test_20f_filer_without_10k_uses_20f_and_6k(self) -> None:
        self.assertEqual(self._forms("6-K", "20-F", "6-K", "20-F/A"), ("20-F", "6-K"))

    def test_40f_filer_without_10k_uses_40f_and_6k(self) -> None:
        self.assertEqual(self._forms("40-F", "6-K"), ("40-F", "6-K"))

    def test_domestic_filer_keeps_the_default(self) -> None:
        self.assertEqual(self._forms("10-K", "10-Q", "8-K"), S.DEFAULT_FORMS)
        self.assertEqual(self._forms("S-1"), S.DEFAULT_FORMS)

    def test_filer_with_10k_and_20f_keeps_the_domestic_default(self) -> None:
        self.assertEqual(self._forms("20-F", "10-K", "8-K"), S.DEFAULT_FORMS)

    def test_forms_are_detected_across_every_batch(self) -> None:
        forms, _label = S.detect_forms([{"form": ["6-K"]}, {"form": ["20-F"]}])
        self.assertEqual(forms, ("20-F", "6-K"))


class ListFilingsTests(unittest.TestCase):

    def _submissions(self, forms: list[str]) -> dict:
        count = len(forms)
        return {"filings": {"recent": {
            "form": forms,
            "filingDate": [f"2025-01-{i + 10}" for i in range(count)],
            "reportDate": ["2024-12-31"] * count,
            "accessionNumber": [f"acc-{i}" for i in range(count)],
            "primaryDocument": [f"d{i}.htm" for i in range(count)]}, "files": []}}

    def test_amendments_are_counted_unless_requested(self) -> None:
        sub = self._submissions(["10-K", "10-K/A", "8-K"])
        with mock.patch.object(S, "_get_json", lambda url, ua: json.loads(json.dumps(sub))):
            rows, omitted = S.list_filings(CIK10, ["10-K", "8-K"], "1994-01-01", UA)
            rows_a, omitted_a = S.list_filings(CIK10, ["10-K", "8-K"], "1994-01-01",
                                               UA, amendments=True)
        self.assertEqual((len(rows), omitted), (2, 1))
        self.assertEqual((len(rows_a), omitted_a), (3, 0))

    def test_foreign_amendments_follow_the_same_rule(self) -> None:
        batches = [self._submissions(["20-F", "20-F/A", "6-K", "6-K/A"])["filings"]["recent"]]
        rows, omitted = S.select_filings(batches, ["20-F", "6-K"], "1994-01-01")
        rows_a, omitted_a = S.select_filings(batches, ["20-F", "6-K"], "1994-01-01",
                                             amendments=True)
        self.assertEqual(([r.form for r in rows], omitted), (["20-F", "6-K"], 2))
        self.assertEqual(([r.form for r in rows_a], omitted_a),
                         (["20-F", "20-F/A", "6-K", "6-K/A"], 0))


class ExhibitListingTests(unittest.TestCase):

    def test_8k_name_is_kept_as_an_alias(self) -> None:
        self.assertIs(S.list_8k_exhibits, S.list_exhibits)

    def test_amended_8k_and_6k_also_carry_exhibits(self) -> None:
        self.assertEqual([S.has_exhibits(f) for f in ("8-K", "8-K/A", "6-K", "6-K/A",
                                                      "10-K", "20-F")],
                         [True, True, True, True, False, False])


class CliTests(_OfflineCase):

    def setUp(self) -> None:
        super().setUp()
        mock.patch.object(S, "resolve_cik", lambda t, ua: CIK10).start()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dest = Path(self.tmp.name) / "sec"
        self.rows = [S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm"),
                     S.FilingRow("10-Q", "2025-05-01", "2025-03-29", "0000320193-25-000015", "p.htm")]

    def _main(self, *extra: str, routes: dict[str, object] | None = None) -> tuple[int, str]:
        net = _fake_net({"q.htm": b"q", "p.htm": b"p"} if routes is None else routes)
        buf = io.StringIO()
        with mock.patch.object(S, "load_submissions",
                               lambda cik, ua, throttle=0.0: [_batch(self.rows)]), \
                mock.patch("urllib.request.urlopen", net), \
                contextlib.redirect_stdout(buf), contextlib.redirect_stderr(io.StringIO()):
            code = S.main(["sec_fetch", "AAPL", "--dest", str(self.dest), "--ua", UA, *extra])
        return code, buf.getvalue()

    def test_default_throttle_is_a_quarter_second(self) -> None:
        code, _ = self._main()
        self.assertEqual(code, 0)
        self.assertEqual(set(self.sleeps), {0.25})

    def test_throttle_option_sets_the_pause_between_requests(self) -> None:
        code, _ = self._main("--throttle", "1.5")
        self.assertEqual(code, 0)
        self.assertEqual(set(self.sleeps), {1.5})

    def test_negative_throttle_is_rejected(self) -> None:
        with self.assertRaises(SystemExit) as ctx:
            self._main("--throttle", "-1")
        self.assertEqual(ctx.exception.code, 2)

    def test_log_option_appends_every_console_line(self) -> None:
        log = Path(self.tmp.name) / "logs" / "run.log"
        _, first = self._main("--log", str(log))
        _, second = self._main("--log", str(log))
        self.assertIn("[..] bajando", first)
        self.assertIn("[ok] ya existe", second)
        logged = log.read_text(encoding="utf-8").splitlines()
        for line in (first + second).splitlines():
            self.assertIn(line, logged)
        self.assertGreater(len(logged), len((first + second).splitlines()))  # + encabezados

    def test_log_records_the_blocked_network_message(self) -> None:
        log = Path(self.tmp.name) / "run.log"
        with self.assertRaises(SystemExit):
            self._main("--log", str(log), routes={"q.htm": 407, "p.htm": b"p"})
        self.assertIn("Sin acceso de red", log.read_text(encoding="utf-8"))

    def test_circuit_breaker_exits_with_code_1(self) -> None:
        self.rows = [S.FilingRow("10-Q", f"2025-0{i}-15", f"2025-0{i}-01",
                                 f"0000320193-25-00000{i}", f"d{i}.htm") for i in range(1, 5)]
        code, out = self._main(routes={"d1.htm": 503, "d2.htm": 503, "d3.htm": 503})
        self.assertEqual(code, 1)
        self.assertIn("detenida", out)


if __name__ == "__main__":
    unittest.main()
