"""Consola ASCII-only: el contrato de CONTRIBUTING 4.3 hecho prueba.

La consola de Windows es cp1252 y un caracter fuera de rango revienta con
UnicodeEncodeError a media corrida. Dos defensas: (1) escaneo estatico de todo
literal que llega a print()/argparse/stderr en tools/, (2) corrida real de los
CLIs con el stdout forzado a ASCII estricto. Stdlib unittest, sin red.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import ast
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = sorted((ROOT / "tools").glob("*.py"))
FETCH_TOOLS = ("fred_fetch.py", "sec_fetch.py", "xbrl_fetch.py")

# Llamadas cuyo texto termina en la consola.
_CONSOLE_FUNCS = frozenset({"print"})
_CONSOLE_METHODS = frozenset({"add_argument", "ArgumentParser", "add_parser", "error", "exit", "write"})


def _is_console_call(node: ast.Call) -> bool:
    func = node.func
    if isinstance(func, ast.Name):
        return func.id in _CONSOLE_FUNCS or func.id in _CONSOLE_METHODS
    if isinstance(func, ast.Attribute):
        if func.attr != "write":
            return func.attr in _CONSOLE_METHODS
        # .write solo cuenta sobre sys.stdout / sys.stderr (no sobre archivos).
        owner = func.value
        return isinstance(owner, ast.Attribute) and owner.attr in ("stdout", "stderr")
    return False


def _non_ascii_console_literals(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    hits: list[str] = []
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Call) and _is_console_call(node)):
            continue
        for sub in ast.walk(node):
            if isinstance(sub, ast.Constant) and isinstance(sub.value, str) and not sub.value.isascii():
                hits.append(f"{path.name}:{sub.lineno}: {ascii(sub.value[:60])}")
    return hits


def _run_ascii(args: list[str]) -> subprocess.CompletedProcess[bytes]:
    env = dict(os.environ, PYTHONIOENCODING="ascii:strict")
    return subprocess.run([sys.executable, *args], cwd=ROOT, env=env, capture_output=True, timeout=120)


class StaticAsciiTests(unittest.TestCase):

    def test_console_literals_in_tools_are_ascii(self) -> None:
        self.assertTrue(TOOLS, "no se encontraron tools/*.py")
        hits = [hit for tool in TOOLS for hit in _non_ascii_console_literals(tool)]
        self.assertEqual(hits, [], "literal no-ASCII rumbo a la consola (rompe cp1252)")

    def test_scan_covers_every_module_of_the_builder(self) -> None:
        # xlsx_builder.py es fachada: el reporte a consola (_print_report) vive
        # en xlsx_audit.py. Si los modulos salen de tools/ (p. ej. a un
        # subpaquete), el glob de arriba dejaria de verlos.
        sys.path.insert(0, str(ROOT / "tools"))
        import xlsx_builder
        parts = {Path(m.__file__).name for m in xlsx_builder._SUBMODULES} | {"xlsx_builder.py"}
        self.assertLessEqual(parts, {tool.name for tool in TOOLS})


class RuntimeAsciiTests(unittest.TestCase):

    def test_fetch_tools_help_survives_ascii_console(self) -> None:
        for name in FETCH_TOOLS:
            with self.subTest(tool=name):
                done = _run_ascii([f"tools/{name}", "--help"])
                self.assertEqual(done.returncode, 0, done.stderr.decode("ascii", "replace"))
                self.assertTrue(done.stdout.isascii())

    def test_builder_demo_and_audit_survive_ascii_console(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book = str(Path(tmp) / "demo.xlsx")
            for args in (["tools/xlsx_builder.py", "demo", book],
                         ["tools/xlsx_builder.py", "audit", book]):
                with self.subTest(cmd=args[1]):
                    done = _run_ascii(args)
                    self.assertEqual(done.returncode, 0, done.stderr.decode("ascii", "replace"))
                    self.assertNotIn(b"UnicodeEncodeError", done.stderr)
                    self.assertTrue(done.stdout.isascii())

    def test_model_check_report_survives_ascii_console(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            book = str(Path(tmp) / "demo_2026-01-01_v1.xlsx")
            self.assertEqual(_run_ascii(["tools/xlsx_builder.py", "demo", book]).returncode, 0)
            done = _run_ascii(["tools/model_check.py", book])
            # 3 = pendiente: el demo sale de openpyxl sin valores calculados (S10).
            self.assertEqual(done.returncode, 3, done.stderr.decode("ascii", "replace"))
            self.assertTrue(done.stdout.isascii())
            self.assertEqual(_run_ascii(["tools/model_check.py", "--help"]).returncode, 0)

    def test_deck_builder_demo_survives_ascii_console(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            done = _run_ascii(["tools/deck_builder.py", "demo", tmp])
            self.assertEqual(done.returncode, 0, done.stderr.decode("ascii", "replace"))
            self.assertNotIn(b"UnicodeEncodeError", done.stderr)
            self.assertTrue(done.stdout.isascii())
            self.assertEqual(_run_ascii(["tools/deck_builder.py", "--help"]).returncode, 0)

    def test_builder_usage_message_is_ascii(self) -> None:
        done = _run_ascii(["tools/xlsx_builder.py"])
        self.assertEqual(done.returncode, 2)
        self.assertNotIn(b"UnicodeEncodeError", done.stderr)
        self.assertTrue(done.stdout.isascii())


if __name__ == "__main__":
    unittest.main()
