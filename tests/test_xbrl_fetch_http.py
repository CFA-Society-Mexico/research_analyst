"""Pruebas del manejo HTTP de tools/xbrl_fetch.py (stdlib unittest, red simulada).

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import io
import sys
import unittest
import urllib.error
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import xbrl_fetch as X  # noqa: E402

UA = "Test test@example.com"
FACTS = X.FACTS_URL.format(cik="0000000001")


class _Resp:
    def read(self) -> bytes:
        return b'{"facts": {}}'

    def __enter__(self) -> "_Resp":
        return self

    def __exit__(self, *exc: object) -> None:
        return None


def _raising(*codes: int):
    queue = list(codes)

    def urlopen(req, timeout=0):
        if queue:
            code = queue.pop(0)
            raise urllib.error.HTTPError(req.full_url, code, "err", {}, io.BytesIO(b""))
        return _Resp()
    return urlopen


class XbrlHttpTests(unittest.TestCase):

    def setUp(self) -> None:
        mock.patch.object(X.time, "sleep").start()
        self.addCleanup(mock.patch.stopall)

    def test_missing_companyfacts_is_not_blocked_network(self) -> None:
        with mock.patch("urllib.request.urlopen", _raising(404)):
            with self.assertRaises(SystemExit) as ctx:
                X._get_json(FACTS, UA)
        self.assertIn("companyfacts", str(ctx.exception))
        self.assertNotIn("Sin acceso de red", str(ctx.exception))

    def test_403_is_blocked_network(self) -> None:
        with mock.patch("urllib.request.urlopen", _raising(403)):
            with self.assertRaises(SystemExit) as ctx:
                X._get_json(FACTS, UA)
        self.assertIn("Sin acceso de red", str(ctx.exception))

    def test_transient_503_is_retried(self) -> None:
        with mock.patch("urllib.request.urlopen", _raising(503, 503)):
            self.assertEqual(X._get_json(FACTS, UA), {"facts": {}})


if __name__ == "__main__":
    unittest.main()
