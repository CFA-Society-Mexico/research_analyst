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
    Valor int = codigo HTTP de error; bytes = cuerpo; lista = respuestas en orden."""
    def urlopen(req, timeout=0):
        url = req.full_url
        for frag, answer in routes.items():
            if frag in url:
                if isinstance(answer, list):
                    answer = answer.pop(0)
                if isinstance(answer, int):
                    raise _http_error(url, answer)
                return _Resp(answer)
        raise AssertionError(f"URL no esperada en la prueba: {url}")
    return urlopen


class HttpClassificationTests(unittest.TestCase):

    def setUp(self) -> None:
        self.sleep = mock.patch.object(S.time, "sleep").start()
        self.addCleanup(mock.patch.stopall)

    def test_404_is_not_reported_as_blocked_network(self) -> None:
        with mock.patch("urllib.request.urlopen", _fake_net({"x.json": 404})):
            with self.assertRaises(S.SecHttpError):
                S._get_json("https://www.sec.gov/x.json", UA)

    def test_403_and_urlerror_are_blocked_network(self) -> None:
        with mock.patch("urllib.request.urlopen", _fake_net({"x.json": 403})):
            with self.assertRaises(SystemExit) as ctx:
                S._get_json("https://www.sec.gov/x.json", UA)
        self.assertIn("Sin acceso de red", str(ctx.exception))

        def down(req, timeout=0):
            raise urllib.error.URLError("proxy")
        with mock.patch("urllib.request.urlopen", down):
            with self.assertRaises(SystemExit):
                S._get_json("https://www.sec.gov/x.json", UA)

    def test_429_is_retried(self) -> None:
        net = _fake_net({"x.json": [429, b'{"ok": 1}']})
        with mock.patch("urllib.request.urlopen", net):
            self.assertEqual(S._get_json("https://www.sec.gov/x.json", UA), {"ok": 1})
        self.sleep.assert_called()


class FetchTests(unittest.TestCase):

    def setUp(self) -> None:
        mock.patch.object(S.time, "sleep").start()
        mock.patch.object(S, "resolve_cik", lambda t, ua: "0000320193").start()
        self.addCleanup(mock.patch.stopall)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dest = Path(self.tmp.name)

    def _run(self, rows: list[S.FilingRow], routes: dict[str, object],
             **kw: object) -> tuple[Path, list[str]]:
        mock.patch.object(S, "list_filings",
                          lambda cik, forms, since, ua, am=False: (rows, 0)).start()
        with mock.patch("urllib.request.urlopen", _fake_net(routes)), \
                contextlib.redirect_stdout(io.StringIO()):
            return S.fetch("AAPL", self.dest, ["10-K", "10-Q", "8-K"], "1994-01-01",
                           UA, **kw)

    def _manifest(self) -> list[dict[str, str]]:
        with (self.dest / "AAPL_filings_manifest.csv").open(encoding="utf-8") as fh:
            return list(csv.DictReader(fh))

    def test_exhibit_index_404_does_not_abort_the_run(self) -> None:
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a8k.htm"),
                S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm")]
        _, failures = self._run(rows, {"index.json": 404, "a8k.htm": b"k", "q.htm": b"q"})
        self.assertEqual(failures, [])
        self.assertEqual(len(self._manifest()), 2)

    def test_same_event_date_8ks_get_distinct_files(self) -> None:
        rows = [S.FilingRow("8-K", "2025-05-01", "2025-05-01", "0000320193-25-000010", "a.htm"),
                S.FilingRow("8-K", "2025-05-02", "2025-05-01", "0000320193-25-000011", "b.htm")]
        self._run(rows, {"index.json": b'{"directory": {"item": []}}',
                         "a.htm": b"primero", "b.htm": b"segundo"})
        files = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual(len(set(files.values())), 2)
        self.assertEqual((self.dest / files["0000320193-25-000011"]).read_bytes(), b"segundo")

    def test_amendment_form_does_not_create_a_folder(self) -> None:
        rows = [S.FilingRow("10-K/A", "2020-01-10", "2019-09-28", "0000320193-20-000001", "ka.htm")]
        self._run(rows, {"ka.htm": b"x"})
        self.assertTrue((self.dest / "AAPL_10-K-A_2019-09-28.htm").exists())
        self.assertEqual(self._manifest()[0]["form"], "10-K/A")

    def test_manifest_is_merged_across_runs(self) -> None:
        first = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", "0000320193-24-000123", "k.htm")]
        second = [S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm")]
        self._run(first, {"k.htm": b"k"})
        self._run(second, {"q.htm": b"q"})
        self.assertEqual({r["form"] for r in self._manifest()}, {"10-K", "10-Q"})

    def test_failed_download_is_reported_and_others_kept(self) -> None:
        rows = [S.FilingRow("10-K", "2024-11-01", "2024-09-28", "0000320193-24-000123", "k.htm"),
                S.FilingRow("10-Q", "2025-08-01", "2025-06-28", "0000320193-25-000020", "q.htm")]
        _, failures = self._run(rows, {"k.htm": 404, "q.htm": b"q"})
        self.assertEqual(failures, ["AAPL_10-K_2024-09-28.htm"])
        self.assertEqual([r["form"] for r in self._manifest()], ["10-Q"])
        self.assertFalse(any(p.suffix == ".part" for p in self.dest.iterdir()))

    def test_row_from_old_colliding_manifest_is_replaced(self) -> None:
        # Manifest escrito por la version anterior: dos accessions -> mismo archivo.
        (self.dest / "AAPL_8-K_2025-05-01.htm").write_bytes(b"primero")
        src_b = S.SEC_ARCHIVES_URL.format(cik="320193", accn="000032019325000011", doc="b.htm")
        with (self.dest / "AAPL_filings_manifest.csv").open("w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=S.MANIFEST_FIELDS)
            w.writeheader()
            w.writerow({"file": "AAPL_8-K_2025-05-01.htm", "form": "8-K", "filed": "2025-05-01",
                        "period": "2025-05-01", "accession": "0000320193-25-000010",
                        "source": src_b.replace("11/b.htm", "10/a.htm")})
            w.writerow({"file": "AAPL_8-K_2025-05-01.htm", "form": "8-K", "filed": "2025-05-02",
                        "period": "2025-05-01", "accession": "0000320193-25-000011",
                        "source": src_b})
        rows = [S.FilingRow("8-K", "2025-05-02", "2025-05-01", "0000320193-25-000011", "b.htm")]
        self._run(rows, {"index.json": b'{"directory": {"item": []}}', "b.htm": b"segundo"})
        by_acc = {r["accession"]: r["file"] for r in self._manifest()}
        self.assertEqual(len(self._manifest()), 2)
        self.assertNotEqual(by_acc["0000320193-25-000011"], "AAPL_8-K_2025-05-01.htm")


class ListFilingsTests(unittest.TestCase):

    def test_amendments_are_counted_unless_requested(self) -> None:
        sub = {"filings": {"recent": {
            "form": ["10-K", "10-K/A", "8-K"],
            "filingDate": ["2024-11-01", "2024-12-01", "2025-01-02"],
            "reportDate": ["2024-09-28", "2024-09-28", "2025-01-02"],
            "accessionNumber": ["a", "b", "c"],
            "primaryDocument": ["a.htm", "b.htm", "c.htm"]}, "files": []}}
        with mock.patch.object(S, "_get_json", lambda url, ua: json.loads(json.dumps(sub))):
            rows, omitted = S.list_filings("0000320193", ["10-K", "8-K"], "1994-01-01", UA)
            rows_a, omitted_a = S.list_filings("0000320193", ["10-K", "8-K"], "1994-01-01",
                                               UA, amendments=True)
        self.assertEqual((len(rows), omitted), (2, 1))
        self.assertEqual((len(rows_a), omitted_a), (3, 0))


if __name__ == "__main__":
    unittest.main()
