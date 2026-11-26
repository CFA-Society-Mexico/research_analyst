"""Pruebas del lint de formulas (check F21) y de la carga del builder por ruta.

F21 nace del issue #33: el audit F daba verde con formulas que Excel descarta
al abrir (``ABS`` con tres argumentos) o evalua mal (``MEDIAN(IF(rango=1,...))``
escrita como formula normal). Stdlib unittest, sin red y sin Excel.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

TOOLS = Path(__file__).resolve().parents[1] / "tools"
sys.path.insert(0, str(TOOLS))

import xlsx_builder as B  # noqa: E402

# (formula, texto que el problema debe nombrar)
REJECTED = (
    ("=MAX(ABS(C1,D1,E1))", "ABS"),                 # el caso de TSMC: 7 checks borrados
    ("=ROUND(C1)", "ROUND"),
    ("=IF(C1)", "IF"),
    ("=IF(C1,1,2,3)", "IF"),
    ("=PI(1)", "PI"),
    ("=IFERROR(C1)", "IFERROR"),
    ("=MEDIAN(IF(C1:C9=1,D1:D9))", "arreglo"),      # mediana de comps "n/a"
    ("=C1:C9*D1", "arreglo"),
    ("=-C1:C9", "arreglo"),
    ("=XLOOKUP(A1,B1:B5,C1:C5)", "_xlfn"),
    ("=IFS(A1>0,1,A1<0,2)", "_xlfn"),
    ("=SUM(C1", "parentesis"),
    ("=A1)", "parentesis"),
)

ACCEPTED = (
    "=ABS(C1)",
    "=MAX(ABS(C1),ABS(D1),ABS(E1))",
    "=ROUND(C1,2)",
    "=IF(C1>0,1)",
    '=IF(C1>0,"a,b(",0)',                           # coma y parentesis dentro de texto
    "=IF(A1,,B1)",                                  # argumento vacio: son 3
    "=NA()",
    "=SUM(C1:C9)/COUNT(C1:C9)",
    "=SUM(Operating!C12:F12)",
    "=AVERAGE(C1,D1,E1)",
    "=SUMPRODUCT(C1:C9/(1+D1)^E1:E9)",              # SUMPRODUCT evalua arreglos nativo
    "=_xlfn.XLOOKUP(A1,B1:B5,C1:C5)",
    '=+IF(ABS(C1-D1)>0.0001,"Error","OK")',
    '=IF(C1<=WACC,"n/s",C2/(WACC-C1))',
    "=IF(OR(C1=0,D1=0),\"\",C1/D1*DAYS_YEAR)",
    "=C1*(1+D1)*(1-D1/E1)",
    "=COUNTIF(Operating!C3:Z3,\"*Q2026A\")",
)


class LintFormulaTests(unittest.TestCase):

    def test_rejected_formulas_name_the_problem(self) -> None:
        for formula, culprit in REJECTED:
            with self.subTest(formula=formula):
                problems = B.lint_formula(formula)
                self.assertTrue(problems, "debio rechazarse")
                self.assertIn(culprit, " | ".join(problems))

    def test_accepted_formulas_are_clean(self) -> None:
        for formula in ACCEPTED:
            with self.subTest(formula=formula):
                self.assertEqual(B.lint_formula(formula), [])

    def test_non_formula_values_are_ignored(self) -> None:
        for value in ("Ventas", "n/s", "", "x"):
            self.assertEqual(B.lint_formula(value), [])


class BuilderPreventionTests(unittest.TestCase):
    """El builder rechaza la formula al construir, no en el audit."""

    def setUp(self) -> None:
        self.styler = B.ModelStyler()
        self.ws = self.styler.new_sheet("Operating")

    def test_set_cell_rejects_a_formula_excel_would_discard(self) -> None:
        with self.assertRaises(ValueError) as caught:
            self.styler.set_cell(self.ws, "C5", "=MAX(ABS(C1,D1,E1))", B.CellRole.FORMULA)
        self.assertIn("ABS", str(caught.exception))
        self.assertIn("Operating!C5", str(caught.exception))
        self.assertIsNone(self.ws["C5"].value)

    def test_set_cell_rejects_array_operation_in_a_normal_formula(self) -> None:
        with self.assertRaises(ValueError):
            self.styler.set_cell(self.ws, "C5", "=MEDIAN(IF(C1:C9=1,D1:D9))",
                                 B.CellRole.FORMULA)

    def test_array_formula_is_the_sanctioned_way_and_audits_clean(self) -> None:
        self.styler.set_array_formula(self.ws, "C5", "=MEDIAN(IF(C1:C4=1,D1:D4))",
                                      B.CellRole.FORMULA, B.NumFmt.MULT)
        self.assertIsInstance(self.ws["C5"].value, ArrayFormula)
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "arr.xlsx")
            self.styler.save(path)
            self.assertIsInstance(load_workbook(path)["Operating"]["C5"].value, ArrayFormula)
            f21 = next(f for f in B.audit_format(path) if f.check.startswith("F21 "))
        self.assertTrue(f21.ok, f21.detail)

    def test_array_formula_still_lints_arity(self) -> None:
        with self.assertRaises(ValueError):
            self.styler.set_array_formula(self.ws, "C5", "=MEDIAN(IF(C1:C4=1))",
                                          B.CellRole.FORMULA)


class LoadByPathTests(unittest.TestCase):
    """Issue #40: cargar el builder por ruta SIN registrarlo en sys.modules
    rompia las dataclasses (anotaciones como texto + modulo no registrado)."""

    def test_builder_loads_by_path_without_sys_modules_registration(self) -> None:
        name = "xlsx_builder_cargado_por_ruta"
        self.assertNotIn(name, sys.modules)
        spec = importlib.util.spec_from_file_location(name, TOOLS / "xlsx_builder.py")
        assert spec is not None and spec.loader is not None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)          # sin sys.modules[name] = module
        self.assertNotIn(name, sys.modules)
        self.assertEqual(module.Finding("F1 x", True, "ok").mark, "[ok]")
        self.assertEqual(module.PeriodHeader(2020, 2030, 2025).last_actual_year, 2025)


if __name__ == "__main__":
    unittest.main()
