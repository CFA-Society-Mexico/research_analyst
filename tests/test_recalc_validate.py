"""Pruebas de tools/recalc_validate.py y de ``model_check.py --recalc``.

Issue #33: nada del plugin recalculaba el libro, asi que un check de formula
viva nunca evaluado contaba como verde. El motor se simula (sin Excel en CI);
``ExcelEngineTests`` lo corre de verdad donde hay Excel.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Callable
from unittest import mock

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.workbook.workbook import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import model_check as M  # noqa: E402
import recalc_validate as R  # noqa: E402
import xlsx_builder as B  # noqa: E402

VERSIONED = "demo_2026-01-01_v1.xlsx"
LIVE_OK = '=IF(ABS(Operating!C27-Operating!C27)>0.0001,"Error","OK")'
LIVE_BROKEN_ARITY = '=IF(MAX(ABS(Operating!C27,Operating!D27,Operating!E27))>0,"Error","OK")'
LIVE_FAILING = '=IF(Operating!C27>0,"Error","OK")'


def _demo_with_checks(path: Path, checks: dict[str, str]) -> None:
    """Demo del builder + formulas vivas en la tab Checks (escritas crudas: el
    caso que se prueba es un libro que NO paso por el lint del builder)."""
    with contextlib.redirect_stdout(io.StringIO()):
        B._demo(str(path))
    wb = load_workbook(path)
    ws = wb["Checks"]
    for row, (label, formula) in enumerate(checks.items(), start=7):
        ws.cell(row=row, column=2, value=label).font = Font(name=B.FONT_NAME, size=11)
        ws.cell(row=row, column=3, value=formula).font = Font(name=B.FONT_NAME, size=11)
    wb.save(path)


def _fake_engine(mutate: Callable[[Workbook], None]) -> Callable[..., Path]:
    """Sustituto de ``recalculate``: copia el libro y le aplica ``mutate``."""
    def recalculate(source: Path, workdir: Path, engine: R.Engine) -> Path:
        out = workdir / "recalculado" / source.name
        out.parent.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(source)
        mutate(wb)
        wb.save(out)
        return out
    return recalculate


class _Case(unittest.TestCase):

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)
        self.model = self.dir / VERSIONED
        self.work = self.dir / "wd"

    def _validate(self, mutate: Callable[[Workbook], None]) -> dict[str, B.Finding]:
        with mock.patch.object(R, "find_engine", return_value=R.Engine.EXCEL), \
                mock.patch.object(R, "recalculate", side_effect=_fake_engine(mutate)):
            findings, _ = R.validate(self.model, self.work)
        return {f.check.split()[0]: f for f in findings}


class ValidateTests(_Case):

    def test_without_an_engine_the_verdict_is_pending_and_says_unevaluated(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_OK})
        with mock.patch.object(R, "find_engine", return_value=None):
            findings, recalculated = R.validate(self.model, self.work)
        self.assertIsNone(recalculated)
        self.assertTrue(findings[0].pending)
        self.assertIn("SIN evaluar", findings[0].detail)

    def test_formula_dropped_by_the_engine_fails_r1_and_names_the_cell(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_OK, "C2": LIVE_OK})

        def repair(wb: Workbook) -> None:       # lo que hace Excel al reparar
            wb["Checks"]["C8"] = None

        r1 = self._validate(repair)["R1"]
        self.assertFalse(r1.ok)
        self.assertIn("Checks!C8", r1.detail)
        self.assertIn("descarto 1 formulas", r1.detail)

    def test_engine_that_cannot_open_the_book_fails_r1_with_the_f21_suspects(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_BROKEN_ARITY})
        with mock.patch.object(R, "find_engine", return_value=R.Engine.EXCEL), \
                mock.patch.object(R, "recalculate",
                                  side_effect=R.RecalcError("excel no pudo abrir")):
            findings, recalculated = R.validate(self.model, self.work)
        self.assertIsNone(recalculated)
        self.assertTrue(findings[0].blocking)
        self.assertIn("Checks!C7", findings[0].detail)
        self.assertIn("ABS", findings[0].detail)
        self.assertIn("SIN evaluar", findings[0].detail)

    def test_calculated_error_value_fails_r2(self) -> None:
        _demo_with_checks(self.model, {})

        def poison(wb: Workbook) -> None:
            wb["Operating"]["T70"] = "#NAME?"

        found = self._validate(poison)
        self.assertTrue(found["R1"].ok, found["R1"].detail)
        self.assertFalse(found["R2"].ok)
        self.assertIn("Operating!T70", found["R2"].detail)

    def test_clean_book_is_green_on_r1_and_r2(self) -> None:
        _demo_with_checks(self.model, {})
        found = self._validate(lambda wb: None)
        self.assertTrue(found["R1"].ok and found["R2"].ok)


class CliTests(_Case):

    def _run(self, *args: str, engine: object = R.Engine.EXCEL,
             mutate: Callable[[Workbook], None] = lambda wb: None) -> tuple[int, str]:
        out = io.StringIO()
        with mock.patch.object(R, "find_engine", return_value=engine), \
                mock.patch.object(R, "recalculate", side_effect=_fake_engine(mutate)), \
                contextlib.redirect_stdout(out):
            code = R.main(["recalc_validate.py", str(self.model), *args])
        return code, out.getvalue()

    def test_missing_model_is_a_usage_error(self) -> None:
        code, text = self._run()
        self.assertEqual(code, 2)
        self.assertIn("[x] no existe el modelo", text)

    def test_no_engine_exits_pending(self) -> None:
        _demo_with_checks(self.model, {})
        code, text = self._run(engine=None)
        self.assertEqual(code, 3)
        self.assertIn("[pendiente] R1", text)

    def test_in_place_is_refused_when_the_engine_dropped_formulas(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_OK})
        before = self.model.read_bytes()

        def repair(wb: Workbook) -> None:
            wb["Checks"]["C7"] = None

        code, text = self._run("--in-place", mutate=repair)
        self.assertEqual(code, 1)
        self.assertIn("--in-place omitido", text)
        self.assertEqual(self.model.read_bytes(), before)       # el modelo no se toca

    def test_in_place_saves_the_recalculated_book_when_r1_is_green(self) -> None:
        _demo_with_checks(self.model, {})

        def mark(wb: Workbook) -> None:
            wb["Cover"]["H40"] = "recalculado"

        _, text = self._run("--in-place", mutate=mark)
        self.assertIn("valores calculados guardados", text)
        self.assertEqual(load_workbook(self.model)["Cover"]["H40"].value, "recalculado")
        self.assertFalse(self.model.with_name(self.model.name + ".part").exists())


class ModelCheckRecalcTests(_Case):

    def _model_check(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = M.main(["model_check.py", str(self.model), *args])
        return code, out.getvalue()

    def test_unevaluated_live_checks_keep_the_report_pending_without_recalc(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_OK})
        code, text = self._model_check()
        self.assertEqual(code, 3)
        self.assertIn("[pendiente] Checks vivos", text)
        self.assertIn("1 de 1 checks vivos SIN evaluar", text)

    def test_recalc_audits_the_recalculated_copy_and_reports_group_r(self) -> None:
        _demo_with_checks(self.model, {})
        seen: list[str] = []
        real_structure = M.audit_structure

        def spy(path: str, schedules: object = None) -> list[B.Finding]:
            seen.append(path)
            return real_structure(path, schedules)  # type: ignore[arg-type]

        with mock.patch.object(R, "find_engine", return_value=R.Engine.EXCEL), \
                mock.patch.object(R, "recalculate", side_effect=_fake_engine(lambda wb: None)), \
                mock.patch.object(M, "audit_structure", side_effect=spy):
            _, text = self._model_check("--recalc")
        self.assertIn("-- Recalculo (R)", text)
        self.assertIn("[ok] R1", text)
        self.assertEqual(len(seen), 1)
        self.assertNotEqual(Path(seen[0]), self.model)           # audita la COPIA
        self.assertEqual(Path(seen[0]).name, VERSIONED)          # con el mismo nombre (S8)
        self.assertEqual(text.count("Checks vivos"), 1)          # no se reporta dos veces

    def test_recalc_without_engine_falls_back_to_the_model_and_stays_pending(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_OK})
        with mock.patch.object(R, "find_engine", return_value=None):
            code, text = self._model_check("--recalc")
        self.assertEqual(code, 3)
        self.assertIn("[pendiente] R1", text)
        self.assertIn("[pendiente] Checks vivos", text)


@unittest.skipUnless(R.find_engine("excel") is R.Engine.EXCEL, "requiere Excel (COM)")
class ExcelEngineTests(_Case):
    """El caso real del issue #33, con Excel de verdad."""

    def test_excel_rejects_abs_with_three_arguments_and_r1_names_it(self) -> None:
        _demo_with_checks(self.model, {"C1": LIVE_BROKEN_ARITY, "C2": LIVE_OK})
        findings, recalculated = R.validate(self.model, self.work, "excel")
        self.assertIsNone(recalculated)
        self.assertTrue(findings[0].blocking, findings[0].detail)
        self.assertIn("Checks!C7", findings[0].detail)

    def test_excel_evaluates_live_checks_and_reports_the_failing_one(self) -> None:
        _demo_with_checks(self.model, {"C1 sano": LIVE_OK, "C2 roto": LIVE_FAILING})
        self.assertTrue(B.live_checks_finding(str(self.model)).pending)     # antes: sin evaluar
        findings, recalculated = R.validate(self.model, self.work, "excel")
        found = {f.check.split()[0]: f for f in findings}
        self.assertTrue(found["R1"].ok, found["R1"].detail)
        self.assertTrue(found["R2"].ok, found["R2"].detail)
        self.assertFalse(found["Checks"].ok)
        self.assertIn("C2 roto", found["Checks"].detail)
        self.assertNotIn("C1 sano", found["Checks"].detail)
        assert recalculated is not None
        self.assertEqual(recalculated.name, VERSIONED)
        self.assertTrue(shutil.which("powershell"))


if __name__ == "__main__":
    unittest.main()
