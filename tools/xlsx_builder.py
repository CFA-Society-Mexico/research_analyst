"""Deterministic xlsx style/scaffold engine for the research_analyst plugin.

Doctrine: format is CODE, not model judgement. model-standards builds every
workbook THROUGH this module (never raw openpyxl styling), and /model-check
runs ``audit_format`` (checks F) against any workbook, plugin-built or not.

Empirical basis: CFI workbook corpus (3-Statement Model Complete, AMZN
Advanced case, Valuation Model, template library) extracted 2026-08-30.

Console output policy: ASCII only ([ok]/[x], no unicode symbols).

Usage as CLI:
    python tools/xlsx_builder.py audit <path.xlsx>     -> run checks F
        [brand/DESIGN.md]  colores de marca aceptados por F4
        [--external]       modelo que NO construyo el plugin: F10 es aviso
    python tools/xlsx_builder.py demo  <path.xlsx>     -> build a v3 skeleton (self-test)

Exit codes: 0 = verde; 1 = alguna falla; 3 = sin fallas pero con checks
[pendiente] (p. ej. libro sin valores calculados: recalcular y re-auditar).
Un [aviso] se lista siempre y nunca cambia el exit code.

Este archivo es la FACHADA: ``import xlsx_builder as B`` y
``from xlsx_builder import ...`` ven todos los nombres de siempre. El codigo
vive por responsabilidad en modulos hermanos de tools/:

    xlsx_contract.py  contrato de formato: Color, NumFmt, CellRole, sellos,
                      contratos de hojas y Ratios, marca, limites de escaneo
    xlsx_formulas.py  texto de formulas: lint_formula (F21) y formulas de
                      valuacion (DCF, comps, banderas)
    xlsx_styler.py    ModelStyler, el builder (build_ratios incluido)
    xlsx_scan.py      lectores del libro para los checks (periodos, ciclos,
                      huecos, series, roll de caja)
    xlsx_audit.py     Finding, audit_format (F1-F21), checks de formula viva
                      y _print_report
    xlsx_builder.py   fachada + _demo + CLI

Los modulos nuevos se importan a traves de la fachada, nunca entre scripts
de cobertura; un modulo hermano jamas importa ``xlsx_builder``.
"""

# SIN ``from __future__ import annotations`` a proposito: con anotaciones
# como texto, ``@dataclass`` consulta ``sys.modules[cls.__module__]`` y truena
# si el modulo se cargo por ruta (``importlib.util.spec_from_file_location`` +
# ``exec_module``) sin registrarlo antes. Con anotaciones evaluadas el builder
# carga igual por ``import`` que por ruta (tests/test_formula_lint.py).

import csv
import sys
import types
from pathlib import Path

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# --- Fachada ------------------------------------------------------------------
# Los submodulos viven junto a este archivo. Al cargar la fachada por ruta sin
# tools/ en sys.path (importlib), los imports de abajo los encuentran igual.
_TOOLS_DIR = str(Path(__file__).resolve().parent)
if _TOOLS_DIR not in sys.path:
    sys.path.append(_TOOLS_DIR)

import xlsx_contract  # noqa: E402
import xlsx_formulas  # noqa: E402
import xlsx_scan  # noqa: E402
import xlsx_styler  # noqa: E402
import xlsx_audit  # noqa: E402

_SUBMODULES = (xlsx_contract, xlsx_formulas, xlsx_scan, xlsx_styler, xlsx_audit)

# Lo que usan el demo y el CLI de abajo (la re-exportacion lo trae igual; el
# import explicito es para quien lee el archivo).
from xlsx_audit import _print_report, audit_format, live_checks_finding  # noqa: E402
from xlsx_contract import (  # noqa: E402
    FONT_NAME,
    FROZEN_SHEET_PREFIXES,
    CellRole,
    NumFmt,
    PeriodHeader,
    load_brand,
)
from xlsx_styler import ModelStyler  # noqa: E402

# Re-exportacion TOTAL: cada nombre de cada submodulo, privados incluidos. Los
# scripts de build de coberturas viven fuera del repo y usan
# ``from xlsx_builder import ...`` / ``B._nombre``: ninguno puede desaparecer
# (tests/test_xlsx_builder_split.py congela la lista).
for _module in _SUBMODULES:
    for _name, _value in vars(_module).items():
        if not _name.startswith("__"):
            globals()[_name] = _value
del _module, _name, _value

_MISSING = object()


class _Facade(types.ModuleType):
    """Modulo fachada: una asignacion sobre ``xlsx_builder`` llega al
    submodulo que usa el nombre.

    ``mock.patch.object(B, "_MAX_SCAN_ROWS", 50)`` hace ``setattr`` sobre la
    fachada, pero el codigo del audit lee el nombre en SU modulo. Cada
    asignacion se replica en todo submodulo cuyo nombre apunta al MISMO objeto
    que tenia la fachada; al restaurar el parche aplica la misma regla. Una
    copia cargada por ruta sin registrarse en ``sys.modules`` no propaga.
    """

    def __setattr__(self, name: str, value: object) -> None:
        old = self.__dict__.get(name, _MISSING)
        if old is not _MISSING:
            for module in _SUBMODULES:
                if module.__dict__.get(name, _MISSING) is old:
                    setattr(module, name, value)
        super().__setattr__(name, value)


if __name__ in sys.modules:
    sys.modules[__name__].__class__ = _Facade


def _write_demo_capture(folder: Path, rows: list[tuple[str, str, float]]) -> None:
    """Capa de captura minima del demo (line_item, periodo, valor): el fixture
    contra el que corre C9. No es un artefacto de cobertura."""
    folder.mkdir(parents=True, exist_ok=True)
    with (folder / "canonical_quarterly.csv").open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(("line_item", "period", "value", "source_doc", "source_ref", "tag"))
        for line_item, period, value in rows:
            writer.writerow((line_item, period, value, "demo", "xlsx_builder._demo", "observado"))


def _demo(path: str) -> None:
    """Self-test: libro v3 minimo construido SOLO con la API del builder.

    Estructura del model-spec vigente (modo quarterly): Operating trimestral
    puro con Assumptions -> IS -> BS -> CF -> Ratios (build_ratios) ->
    Schedules, y Annual como agregado por formula. Debe pasar su propio audit
    en verde; un check nuevo debe fallar en un libro malo y pasar aqui.
    """
    from pathlib import Path as _Path

    _Path(path).parent.mkdir(parents=True, exist_ok=True)
    styler = ModelStyler()
    styler.set_periodicity("quarterly")
    wb = styler.wb

    def label(ws: Worksheet, row: int, text: str) -> None:
        ws.cell(row=row, column=2, value=text).font = Font(name=FONT_NAME, size=11)

    for name in ("Cover", "Checks"):
        ws = styler.new_sheet(name, freeze=None)
        styler.brand_bar(ws, name)
        styler.label_col_width(ws)
    label(wb["Checks"], 5, "D3 Staleness de comps (ejemplo de escaneo)")
    styler.check_result(wb["Checks"], 5, 3, "OK", note="demo: escaneo de ejemplo",
                        computed_at="2026-01-01")

    # --- Operating: trimestral puro, 8 trimestres A + 8 E ---------------------
    op = styler.new_sheet("Operating")
    styler.brand_bar(op, "Operating model (trimestral)")
    styler.label_col_width(op)
    quarters = [f"{q}Q{y}{'A' if y <= 2024 else 'E'}"
                for y in range(2023, 2027) for q in (1, 2, 3, 4)]
    first, n_cols, n_act = 3, len(quarters), 8
    styler.quarter_header(op, 3, first, quarters)
    cols = [get_column_letter(first + i) for i in range(n_cols)]

    rows = {"days_q": 6, "days_y": 7, "wacc": 8, "growth": 9,
            "rev": 12, "cogs": 13, "gross": 14, "da": 15, "ebit": 16,
            "interest": 17, "ebt": 18, "tax": 19, "ni": 20,
            "cash": 23, "ar": 24, "inv": 25, "ca": 26, "ta": 27, "ap": 28,
            "cl": 29, "debt": 30, "equity": 31, "re": 32, "cfo": 35}
    # Linea de captura de cada fila OBSERVADA del demo (contrato del C9).
    capture_lines = {"rev": "is_ns_total", "cogs": "is_cogs_total", "da": "cf_da",
                     "interest": "is_interest", "tax": "is_tax", "cash": "bs_cash",
                     "ar": "bs_ar", "inv": "bs_inv", "ca": "bs_ca", "ta": "bs_ta",
                     "ap": "bs_ap", "cl": "bs_cl", "debt": "bs_debt",
                     "equity": "bs_equity", "re": "bs_re", "cfo": "cf_cfo"}
    captured: list[tuple[str, str, float]] = []

    styler.section_header(op, 5, "Assumptions")
    for key, text, value, fmt in (("days_q", "Dias por trimestre", 91.25, NumFmt.DEC2),
                                  ("days_y", "Dias por anio", 365, NumFmt.NUM),
                                  ("wacc", "WACC", 0.09, NumFmt.PCT2)):
        label(op, rows[key], text)
        styler.set_cell(op, f"C{rows[key]}", value, CellRole.INPUT, fmt)
    styler.define_constant("DAYS_QUARTER", "Operating", f"C{rows['days_q']}")
    styler.define_constant("DAYS_YEAR", "Operating", f"C{rows['days_y']}")
    styler.define_constant("WACC", "Operating", f"C{rows['wacc']}")
    r_rev, r_g = rows["rev"], rows["growth"]
    styler.series_row(
        op, r_g, "Crecimiento de ventas yoy (%)", first,
        hist_values=[None] * 4 + [f"={c}{r_rev}/{cols[i - 4]}{r_rev}-1"
                                  for i, c in enumerate(cols[4:n_act], start=4)],
        forecast_values=[0.05] * (n_cols - n_act), numfmt=NumFmt.PCT1)
    styler.group_rows(op, 6, 9)

    def line(key: str, text: str, hist: object, fc: object) -> None:
        """Una serie = una fila: observado en A, formula en E (o formula en todo)."""
        r = rows[key]
        label(op, r, text)
        if not callable(hist):
            # Fila con historico observado: se vincula y su serie va al CSV.
            styler.link_capture(op, r, capture_lines[key])
            captured.extend((capture_lines[key], quarters[i][:-1], float(hist[i]))
                            for i in range(n_act))
        for i, c in enumerate(cols):
            prev = cols[i - 1] if i else c
            if i < n_act and not callable(hist):
                styler.set_cell(op, f"{c}{r}", hist[i], CellRole.OBSERVED)
                continue
            make = hist if (i < n_act or fc is None) else fc
            styler.set_cell(op, f"{c}{r}", make(c, prev, i), CellRole.FORMULA)

    rv = [100 + 3 * i for i in range(n_act)]
    styler.section_header(op, 11, "Estado de resultados")
    line("rev", "Ventas", rv, lambda c, p, i: f"={cols[i - 4]}{r_rev}*(1+{c}{r_g})")
    line("cogs", "Costo de ventas", [round(0.6 * v) for v in rv],
         lambda c, p, i: f"={c}{r_rev}*{p}{rows['cogs']}/{p}{r_rev}")
    line("gross", "Utilidad bruta", lambda c, p, i: f"={c}{r_rev}-{c}{rows['cogs']}", None)
    line("da", "D&A", [5] * n_act, lambda c, p, i: f"={p}{rows['da']}")
    line("ebit", "EBIT", lambda c, p, i: f"={c}{rows['gross']}-{c}{rows['da']}", None)
    line("interest", "Gasto por intereses", [2] * n_act,
         lambda c, p, i: f"={p}{rows['interest']}")
    line("ebt", "Utilidad antes de impuestos",
         lambda c, p, i: f"={c}{rows['ebit']}-{c}{rows['interest']}", None)
    line("tax", "Impuestos", [8] * n_act,
         lambda c, p, i: f"={c}{rows['ebt']}*{p}{rows['tax']}/{p}{rows['ebt']}")
    line("ni", "Utilidad neta", lambda c, p, i: f"={c}{rows['ebt']}-{c}{rows['tax']}", None)
    styler.group_rows(op, 12, 20)

    styler.section_header(op, 22, "Balance")
    carry = lambda key: (lambda c, p, i: f"={p}{rows[key]}")  # noqa: E731
    for key, text, base in (("cash", "Efectivo", 40), ("ar", "Cuentas por cobrar", 30),
                            ("inv", "Inventario", 25), ("ca", "Activo circulante", 95),
                            ("ta", "Activo total", 300), ("ap", "Cuentas por pagar", 20),
                            ("cl", "Pasivo circulante", 45), ("debt", "Deuda total", 80),
                            ("equity", "Capital contable", 150),
                            ("re", "Utilidades retenidas", 60)):
        line(key, text, [base + i for i in range(n_act)], carry(key))
    styler.group_rows(op, 23, 32)

    styler.section_header(op, 34, "Flujo de efectivo")
    line("cfo", "Flujo de operacion", [22] * n_act,
         lambda c, p, i: f"={c}{rows['ni']}+{c}{rows['da']}")
    styler.group_rows(op, 35, 35)

    ref = {k: f"Operating!{{c}}{rows[k]}" for k in (
        "rev", "cogs", "gross", "ebit", "ebt", "ni", "interest", "tax", "ta",
        "equity", "cash", "ar", "inv", "ap", "ca", "cl", "debt", "re", "cfo", "da")}
    # window=4 explicito: ref apunta a flujos de UN trimestre.
    end_row, skipped = styler.build_ratios(op, 37, first, n_cols, ref, wacc_ref="WACC",
                                           window=4)
    assert not skipped, skipped
    heads = [r for r in range(37, end_row) if op.cell(row=r, column=1).value == "x"]
    for i, hr in enumerate(heads):
        styler.group_rows(op, hr + 1, (heads[i + 1] if i + 1 < len(heads) else end_row) - 1)

    r = end_row + 1
    styler.section_header(op, r, "Schedules")
    for block, text, key in (("PPE", "D&A (link a IS)", "da"),
                             ("Debt", "Deuda (link a BS)", "debt"),
                             ("WC", "Cuentas por cobrar (link a BS)", "ar")):
        r += 1
        styler.schedule_block_header(op, r, block)
        label(op, r + 1, text)
        for c in cols:
            styler.set_cell(op, f"{c}{r + 1}", f"={c}{rows[key]}", CellRole.FORMULA)
        styler.group_rows(op, r + 1, r + 1)
        r += 2

    # --- Annual: agregados por formula, cero inputs ---------------------------
    an = styler.new_sheet("Annual")
    styler.brand_bar(an, "Annual (agregado de Operating)")
    styler.label_col_width(an)
    styler.period_header(an, 3, 3, PeriodHeader(2023, 2026, 2024))
    styler.section_header(an, 5, "Estado de resultados (suma de 4 trimestres)")
    for r_an, (key, text) in enumerate((("rev", "Ventas"), ("ni", "Utilidad neta")), start=6):
        label(an, r_an, text)
        for y in range(4):
            a, b = cols[4 * y], cols[4 * y + 3]
            styler.set_cell(an, f"{get_column_letter(3 + y)}{r_an}",
                            f"=SUM(Operating!{a}{rows[key]}:{b}{rows[key]})", CellRole.LINK)
    styler.group_rows(an, 6, 7)
    styler.section_header(an, 9, "Balance (saldo del 4Q)")
    label(an, 10, "Efectivo")
    for y in range(4):
        styler.set_cell(an, f"{get_column_letter(3 + y)}10",
                        f"=Operating!{cols[4 * y + 3]}{rows['cash']}", CellRole.LINK)
    styler.group_rows(an, 10, 10)

    for name in ("Macro", "Rev_Reconcile", "Val_Comps", "Sensitivity", "Summary"):
        freeze = "C4" if name.startswith(FROZEN_SHEET_PREFIXES) else None
        ws = styler.new_sheet(name, freeze=freeze)
        styler.brand_bar(ws, name)
        styler.label_col_width(ws)
    styler.save(path)
    _write_demo_capture(Path(path).parent / "inputs", captured)
    print(f"[ok] demo escrito: {path}")


def main(argv: list[str]) -> int:
    external = "--external" in argv[1:]
    args = [a for a in argv if a != "--external"]
    if args[1:2] == ["audit"] and len(args) in (3, 4):
        brand = load_brand(args[3]) if len(args) == 4 else None
        return _print_report(audit_format(args[2], brand=brand, external=external),
                             live=live_checks_finding(args[2]))
    if len(args) == 3 and args[1] == "demo" and not external:
        _demo(args[2])
        return _print_report(audit_format(args[2]))
    print("uso: python xlsx_builder.py audit <path.xlsx> [brand/DESIGN.md] [--external]"
          " | demo <path.xlsx>")
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
