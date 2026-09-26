"""Pruebas de tools/fred_fetch.py (stdlib unittest, red simulada).

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

import fred_fetch as F  # noqa: E402

KEY = "abcdef0123456789abcdef0123456789"


class _Resp:
    def __init__(self, payload: dict) -> None:
        self._body = json.dumps(payload).encode("utf-8")

    def read(self) -> bytes:
        return self._body

    def __enter__(self) -> "_Resp":
        return self

    def __exit__(self, *exc: object) -> None:
        return None


def _fake_fred(bad_series: str = "TYPO", bad_message: str = "The series does not exist."):
    def urlopen(req, timeout=0):
        url = req.full_url
        if f"series_id={bad_series}&" in url:
            body = json.dumps({"error_code": 400, "error_message": f"Bad Request.  {bad_message}"})
            raise urllib.error.HTTPError(url, 400, "Bad Request", {}, io.BytesIO(body.encode()))
        if "/observations" in url:
            return _Resp({"observations": [{"date": "2025-01-02", "value": "4.5"},
                                           {"date": "2025-01-03", "value": "."}]})
        return _Resp({"seriess": [{"title": "t", "units": "u", "frequency": "d"}]})
    return urlopen


class FredFetchTests(unittest.TestCase):

    def setUp(self) -> None:
        mock.patch.object(F.time, "sleep").start()
        self.addCleanup(mock.patch.stopall)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.series = Path(self.tmp.name) / "macro" / "series"

    def _fetch(self, wanted: dict[str, str], net=None) -> tuple[list[str], str]:
        out = io.StringIO()
        with mock.patch("urllib.request.urlopen", net or _fake_fred()), \
                contextlib.redirect_stdout(out):
            failed = F.fetch(self.series, KEY, wanted)
        return failed, out.getvalue()

    def _manifest(self) -> list[str]:
        with (self.series / "manifest.csv").open(encoding="utf-8") as fh:
            return [r["series"] for r in csv.DictReader(fh)]

    def test_bad_series_id_does_not_abort_the_batch(self) -> None:
        wanted = {"DGS10": "a", "TYPO": "b", "FEDFUNDS": "c"}
        failed, out = self._fetch(wanted)
        self.assertEqual(failed, ["TYPO"])
        self.assertIn("[x] TYPO", out)
        self.assertEqual(self._manifest(), ["DGS10", "FEDFUNDS"])

    def test_invalid_key_stops_everything_without_printing_it(self) -> None:
        net = _fake_fred(bad_series="DGS10",
                         bad_message=f"The value for variable api_key is not registered ({KEY}).")
        with self.assertRaises(SystemExit) as ctx:
            self._fetch({"DGS10": "a"}, net)
        self.assertIn("API key", str(ctx.exception))
        self.assertNotIn(KEY, str(ctx.exception))

    def test_series_error_message_is_redacted(self) -> None:
        net = _fake_fred(bad_series="DGS10", bad_message=f"echo {KEY}")
        failed, out = self._fetch({"DGS10": "a"}, net)
        self.assertEqual(failed, ["DGS10"])
        self.assertNotIn(KEY, out)

    def test_manifest_keeps_series_from_previous_runs(self) -> None:
        self._fetch({"DGS10": "a", "FEDFUNDS": "b", "DEXMXUS": "c"})
        self._fetch({"DGS10": "a"})                      # como --series DGS10
        self.assertEqual(self._manifest(), ["DEXMXUS", "DGS10", "FEDFUNDS"])

    def test_main_exit_code_reflects_failures(self) -> None:
        with mock.patch("urllib.request.urlopen", _fake_fred()), \
                contextlib.redirect_stdout(io.StringIO()):
            code = F.main(["fred_fetch.py", "--dest", str(self.series),
                           "--api-key", KEY, "--series", "DGS10,TYPO"])
        self.assertEqual(code, 1)


if __name__ == "__main__":
    unittest.main()
