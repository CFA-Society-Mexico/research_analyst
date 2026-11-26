"""Checks estructurales (S) y de escaneo contable (C10) del modelo, por codigo.

Doctrina: un check es determinista o no es un check. Los F ya vivian en
``xlsx_builder.py audit``; los S y el C10 se declaraban "escaneo por codigo"
en integrity-checks.md pero cada cobertura los reimplementaba a mano (el
``run_checks.py`` del dogfood de AAPL). Este modulo es la implementacion unica:
mismo libro, mismo veredicto, sin interpretacion del agente.

Solo lectura: jamas modifica el libro.

Uso:
    python tools/model_check.py <modelo.xlsx> [--driver-map driver-map.md]
                                [--brand brand/DESIGN.md] [--external] [--recalc]

Corre S (estructurales) -> C por escaneo (C10) -> F (formato, via
``xlsx_builder.audit_format``) -> estado de los checks de formula viva (tabs
Checks y Cover). Los D (doctrina) no son escaneo de codigo y no se reportan
aqui.

``--recalc``: antes de auditar, recalcula una COPIA del modelo con Excel o
LibreOffice (``recalc_validate``) y audita esa copia ya calculada: R1 (el
motor acepta el libro), R2 (sin celdas en error), S10/F19 y los checks vivos
dejan de estar pendientes. Sin ``--recalc`` y sin valores calculados, los
checks vivos se reportan [pendiente]: SIN evaluar, nunca verdes.

Exit codes (los mismos del audit F): 0 = verde; 1 = alguna falla; 3 = sin
fallas pero con checks [pendiente] (libro sin valores calculados: recalcular
con Excel y re-auditar). Un pendiente NUNCA es verde. Un [aviso] se lista
siempre y nunca cambia el exit code.

Consola ASCII-only.
"""

from __future__ import annotations

import argparse
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Sequence

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_builder as B  # noqa: E402
from xlsx_builder import Finding  # noqa: E402

# ---------------------------------------------------------------------------
# Contrato (model-spec.md + integrity-checks.md)
# ---------------------------------------------------------------------------

BASE_TABS = ("Cover", "Checks", "Macro")
CORE_TABS_BY_MODE = {"quarterly": ("Operating", "Annual"), "annual": ("Model",)}
CORE_SCHEDULES = ("PPE", "Debt", "WC")
# S5: los grids de periodo del modelo. Las hojas auxiliares con header de anios
# (Val_Comps, Rev_Reconcile) mezclan tablas de corte transversal -columnas que
# son metricas, no periodos- y ahi "misma formula por columna" no aplica.
PERIOD_GRID_SHEETS = ("Operating", "Annual", "Model")
ASSUMPTIONS_TITLE = "assumptions"

# S9: lo que debe vivir como seccion/bloque y no como hoja suelta.
LOOSE_SHEET_NAMES = frozenset({"IS", "BS", "CF", "Val_DCF"})
LOOSE_SHEET_PREFIXES = ("Sch_", "Sch ")

VOLATILE_FUNCTIONS = ("OFFSET", "INDIRECT", "NOW", "TODAY")

# S4: hojas sin calculo propio (leyenda, umbrales de checks, datos observados).
NON_CALC_SHEETS = frozenset({"Cover", "Checks", "Macro"})
# S4: literales estructurales (x*(1+g), x/y-1, IF(...,0)). Todo lo demas es un
# supuesto escondido y va a Assumptions o a una constante con nombre.
STRUCTURAL_LITERALS = frozenset({0.0, 1.0})
# Fila reservada por el builder para el check por columna (ModelStyler.check_row).
CHECK_ROW = 3
CHECK_MARKERS = ('"Error"', '"OK"')

VERSIONED_NAME = re.compile(r"_\d{4}-\d{2}-\d{2}_v\d+\.xlsx$")

_STRING_LITERAL = re.compile(r'"[^"]*"')
_SHEET_PREFIX = re.compile(r"(?:'[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!")
_FUNCTION_NAME = re.compile(r"[A-Za-z][A-Za-z0-9_.]*\(")
_PLAIN_REF = re.compile(r"(?<![A-Za-z0-9_.])\$?[A-Z]{1,3}\$?\d+(?![A-Za-z0-9_(])")
_IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_NUMBER = re.compile(r"\d+(?:\.\d+)?(?:[eE][+-]?\d+)?")
_SCHEDULE_BLOCK = re.compile(r"Sch:\s*([^`|<>\n]+?)\s*[`|\n]")

# Referencia A1 con hoja y rango opcionales; los grupos de "$" distinguen una
# columna absoluta (constante) de una relativa (periodo).
_REF = re.compile(
    r"(?<![A-Za-z0-9_.$'!])"
    r"(?:(?:'(?P<qsheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"(?P<abs1>\$?)(?P<col1>[A-Z]{1,3})\$?(?P<row1>\d+)"
    r"(?::(?P<abs2>\$?)(?P<col2>[A-Z]{1,3})\$?(?P<row2>\d+))?"
    r"(?![A-Za-z0-9_(])"
)


@dataclass(frozen=True)
class PeriodColumn:
    """Una columna de periodo del header: anio fiscal y, si aplica, trimestre."""

    column: int
    actual: bool
    year: int
    quarter: int | None


@dataclass(frozen=True)
class Book:
    """El libro cargado una sola vez, con lo que todos los checks consultan."""

    path: Path
    formulas: Workbook
    mode: str | None  # "quarterly" | "annual" | None (no determinable)

    @property
    def core_sheet(self) -> Worksheet | None:
        """Hoja donde se construye el modelo: Operating (trimestral) o Model."""
        name = {"quarterly": "Operating", "annual": "Model"}.get(self.mode or "")
        return self.formulas[name] if name and name in self.formulas.sheetnames else None


# ---------------------------------------------------------------------------
# Lectura del libro
# ---------------------------------------------------------------------------


def _detect_mode(wb: Workbook) -> str | None:
    """Periodicidad: sello del builder; sin sello, se infiere de las hojas."""
    stamp = re.search(r"periodicity=([a-z_]+)", wb.properties.keywords or "")
    if stamp:
        return "annual" if stamp.group(1) == "annual" else "quarterly"
    if "Operating" in wb.sheetnames:
        return "quarterly"
    if "Model" in wb.sheetnames:
        return "annual"
    return None


def load_book(path: str) -> Book:
    wb = load_workbook(path, data_only=False)
    return Book(path=Path(path), formulas=wb, mode=_detect_mode(wb))


def _cells(ws: Worksheet) -> Iterator[Cell]:
    for row in ws.iter_rows(max_row=min(ws.max_row, B._MAX_SCAN_ROWS),
                            max_col=min(ws.max_column, B._MAX_SCAN_COLS)):
        yield from row


def _formula_cells(ws: Worksheet) -> Iterator[Cell]:
    for cell in _cells(ws):
        if isinstance(cell.value, str) and cell.value.startswith("="):
            yield cell


def _is_input(cell: Cell) -> bool:
    fill = cell.fill
    return (cell.value is not None and fill is not None and fill.patternType == "solid"
            and getattr(fill.fgColor, "rgb", None) == B.Color.INPUT_FILL.value)


def _section_rows(ws: Worksheet, title: str) -> tuple[int, int] | None:
    """Filas (primera, ultima) de una seccion: de su banda a la siguiente banda.

    Banda de seccion = label en columna B, bold y de 16 pt o mas (``section_header``).
    """
    first: int | None = None
    last_row = min(ws.max_row, B._MAX_SCAN_ROWS)
    for r in range(1, last_row + 1):
        cell = ws.cell(row=r, column=2)
        font = cell.font
        if not (isinstance(cell.value, str) and font is not None and font.bold
                and (font.size or 0) >= 16):
            continue
        if first is not None:
            return first, r - 1
        if cell.value.strip().lower() == title:
            first = r
    return (first, last_row) if first is not None else None


def _period_axis(ws: Worksheet) -> list[PeriodColumn]:
    """Columnas de periodo con su anio/trimestre (mismas reglas que el audit F:
    anios con formato 0"A"/0"E" y trimestres de texto '1Q2026E', filas 1-8)."""
    axis: list[PeriodColumn] = []
    for row in ws.iter_rows(min_row=1, max_row=8, max_col=min(ws.max_column, B._MAX_SCAN_COLS)):
        for cell in row:
            value = cell.value
            if isinstance(value, bool) or value is None:
                continue
            if cell.number_format in (B.NumFmt.YEAR_A.value, B.NumFmt.YEAR_E.value):
                if isinstance(value, (int, float)):
                    axis.append(PeriodColumn(cell.column, cell.number_format == B.NumFmt.YEAR_A.value,
                                             int(value), None))
            elif isinstance(value, str):
                match = re.fullmatch(r"([1-4])Q(20\d\d)([AE])", value.strip())
                if match:
                    axis.append(PeriodColumn(cell.column, match.group(3) == "A",
                                             int(match.group(2)), int(match.group(1))))
        if len(axis) >= 4:
            break
    return axis if len(axis) >= 4 else []


def _examples(hits: Sequence[str], limit: int = 6) -> str:
    extra = f" (+{len(hits) - limit} mas)" if len(hits) > limit else ""
    return "; ".join(hits[:limit]) + extra


def schedules_from_driver_map(path: str) -> list[str]:
    """Bloques ``Sch: <nombre>`` que el driver-map exige (columna Bloque)."""
    text = Path(path).read_text(encoding="utf-8")
    names = [m.group(1).strip() for m in _SCHEDULE_BLOCK.finditer(text + "\n")]
    return sorted({n for n in names if n})


# ---------------------------------------------------------------------------
# Checks S
# ---------------------------------------------------------------------------


def _check_s1(book: Book, driver_schedules: Sequence[str] | None) -> Finding:
    name = "S1 tabs y bloques del contrato"
    sheets = book.formulas.sheetnames
    if book.mode is None:
        return Finding(name, False, "sin hoja Operating ni Model: no hay modelo que auditar")
    required = BASE_TABS + CORE_TABS_BY_MODE[book.mode]
    missing_tabs = [t for t in required if t not in sheets]
    blocks: set[str] = set()
    for tab in CORE_TABS_BY_MODE[book.mode]:
        if tab not in sheets:
            continue
        ws = book.formulas[tab]
        for r in range(1, min(ws.max_row, B._MAX_SCAN_ROWS) + 1):
            label = ws.cell(row=r, column=2).value
            if isinstance(label, str) and label.startswith("Sch: "):
                blocks.add(label[5:].strip())
    expected = list(CORE_SCHEDULES) + list(driver_schedules or ())
    missing_blocks = sorted({s for s in expected if s not in blocks})
    problems: list[str] = []
    if missing_tabs:
        problems.append("faltan tabs: " + ", ".join(missing_tabs))
    if missing_blocks:
        problems.append("faltan bloques Sch: " + ", ".join(missing_blocks))
    if problems:
        return Finding(name, False, "; ".join(problems))
    scope = "cores + driver-map" if driver_schedules else "cores (sin --driver-map: bloques de drivers no evaluados)"
    return Finding(name, True, f"modo {book.mode}; {len(blocks)} bloques Sch, {scope}")


def _check_s2_s3(book: Book) -> list[Finding]:
    external: list[str] = []
    volatile: list[str] = []
    for ws in book.formulas.worksheets:
        for cell in _formula_cells(ws):
            body = _STRING_LITERAL.sub('""', str(cell.value))
            where = f"{ws.title}!{cell.coordinate}"
            if "[" in body:
                external.append(where)
            upper = body.upper()
            for fn in VOLATILE_FUNCTIONS:
                if re.search(rf"(?<![A-Z0-9_.]){fn}\(", upper):
                    volatile.append(f"{where} {fn}")
    return [
        Finding("S2 cero links a otros libros", not external, _examples(external) or "ok"),
        Finding("S3 sin funciones volatiles", not volatile, _examples(volatile) or "ok"),
    ]


def hardcoded_literals(formula: str) -> list[str]:
    """Numeros tecleados dentro de una formula (S4), fuera de los estructurales.

    Se descartan antes: textos entre comillas, prefijos de hoja, nombres de
    funcion (LOG10), referencias de celda y nombres definidos (DAYS_YEAR).
    """
    body = _STRING_LITERAL.sub('""', formula)
    body = _SHEET_PREFIX.sub("", body)
    body = _FUNCTION_NAME.sub("(", body)
    body = _PLAIN_REF.sub("#", body)
    body = _IDENTIFIER.sub("#", body)
    return [m.group(0) for m in _NUMBER.finditer(body)
            if float(m.group(0)) not in STRUCTURAL_LITERALS]


def _check_s4(book: Book) -> Finding:
    name = "S4 sin numeros hard-coded en formulas"
    core = book.core_sheet
    assumptions = _section_rows(core, ASSUMPTIONS_TITLE) if core is not None else None
    hits: list[str] = []
    for ws in book.formulas.worksheets:
        if ws.title in NON_CALC_SHEETS:
            continue
        for cell in _formula_cells(ws):
            formula = str(cell.value)
            if ws is core and assumptions and assumptions[0] <= cell.row <= assumptions[1]:
                continue  # Assumptions: ahi VIVEN los numeros
            if cell.row == CHECK_ROW or any(mark in formula for mark in CHECK_MARKERS):
                continue  # tolerancia de un check, no un supuesto
            literals = hardcoded_literals(formula)
            if literals:
                hits.append(f"{ws.title}!{cell.coordinate} -> {', '.join(literals[:3])}")
    return Finding(name, not hits,
                   _examples(hits) or "solo 0 y 1 estructurales (exentas: Assumptions y tolerancias de checks)")


def relative_form(formula: str, column: int, year: int | None,
                  quarter_axes: dict[str, dict[int, PeriodColumn]],
                  defined_names: frozenset[str]) -> str:
    """Forma relativa de una formula para compararla entre columnas (S5).

    Columna relativa -> desplazamiento respecto a la columna de la celda.
    Referencia desde una hoja anual a una trimestral -> 'trimestre N del anio
    relativo': la fila anual avanza UNA columna por anio y sus referencias a
    la hoja trimestral avanzan CUATRO, asi que en A1 nunca coincidirian.
    Columnas absolutas ($C) se conservan: son constantes, no periodos.
    """

    def endpoint(absolute: str, letters: str, row: str, target: str | None) -> str:
        if absolute:
            return f"${letters}{row}"
        index = column_index_from_string(letters)
        quarters = quarter_axes.get(target or "")
        if year is not None and quarters and index in quarters:
            period = quarters[index]
            return f"<q{period.quarter}y{period.year - year:+d}>r{row}"
        return f"<c{index - column:+d}>r{row}"

    def replace(match: re.Match[str]) -> str:
        token = match.group(0)
        if token in defined_names:
            return token
        target = match.group("qsheet") or match.group("sheet")
        out = (f"{target}!" if target else "") + endpoint(
            match.group("abs1"), match.group("col1"), match.group("row1"), target)
        if match.group("col2"):
            out += ":" + endpoint(match.group("abs2"), match.group("col2"),
                                  match.group("row2"), target)
        return out

    return _REF.sub(replace, _STRING_LITERAL.sub('""', formula))


def _segments(axis: Sequence[PeriodColumn], first_quarterly_year: int | None,
              transition_year: int | None = None) -> list[list[PeriodColumn]]:
    """Tramos del model-spec: historico, estimado y, en hoja anual con modelo
    trimestral, los anios pre-corte (observados anuales, formula distinta).

    ``transition_year`` (sello ``set_transition_year`` del builder): en una
    hoja ANUAL, la columna estimada de ese anio es un tramo propio - se ancla
    en los trimestres ya observados y su formula difiere a proposito del resto
    del forecast. Sin sello no hay excepcion: una formula distinta en la
    primera columna E sigue siendo falla. En hoja trimestral no aplica (cada
    trimestre es observado o estimado, no hay anclaje parcial).
    """
    is_quarterly = any(p.quarter is not None for p in axis)
    actual = [p for p in axis if p.actual]
    estimate = [p for p in axis if not p.actual]
    if is_quarterly:
        return [actual, estimate]
    transition = [p for p in estimate if p.year == transition_year]
    estimate = [p for p in estimate if p.year != transition_year]
    if first_quarterly_year is None:
        return [actual, transition, estimate]
    return [[p for p in actual if p.year < first_quarterly_year],
            [p for p in actual if p.year >= first_quarterly_year], transition, estimate]


def _check_s5(book: Book) -> Finding:
    name = "S5 una formula por fila y por tramo"
    wb = book.formulas
    axes = {ws.title: _period_axis(ws) for ws in wb.worksheets}
    quarter_axes = {title: {p.column: p for p in axis}
                    for title, axis in axes.items() if any(p.quarter is not None for p in axis)}
    quarterly_years = [p.year for cols in quarter_axes.values() for p in cols.values()]
    first_quarterly_year = min(quarterly_years) if quarterly_years else None
    defined = frozenset(wb.defined_names)
    transition_year = B.transition_year_of(wb.properties.keywords)
    hits: list[str] = []
    scanned: list[str] = []
    for ws in wb.worksheets:
        axis = axes[ws.title]
        if ws.title not in PERIOD_GRID_SHEETS or not axis:
            continue
        scanned.append(ws.title)
        is_annual = ws.title not in quarter_axes
        for segment in _segments(axis, first_quarterly_year, transition_year):
            if len(segment) < 2:
                continue
            for r in range(1, min(ws.max_row, B._MAX_SCAN_ROWS) + 1):
                forms: set[str] = set()
                for period in segment:
                    value = ws.cell(row=r, column=period.column).value
                    if isinstance(value, str) and value.startswith("="):
                        forms.add(relative_form(value, period.column,
                                                period.year if is_annual else None,
                                                quarter_axes, defined))
                if len(forms) > 1:
                    label = str(ws.cell(row=r, column=2).value or "")[:30]
                    span = (f"{get_column_letter(segment[0].column)}-"
                            f"{get_column_letter(segment[-1].column)}")
                    hits.append(f"{ws.title}!fila {r} ({label}): {len(forms)} formulas en {span}")
    if not scanned:
        return Finding(name, False, "sin hoja Operating/Annual/Model con header de periodos")
    scope = ", ".join(scanned) + (f"; anio de transicion {transition_year} como tramo propio"
                                  if transition_year else "")
    return Finding(name, not hits,
                   _examples(hits) or f"formula identica dentro de cada tramo ({scope})")


def _check_s6(book: Book) -> Finding:
    name = "S6 inputs solo en Assumptions"
    core = book.core_sheet
    if core is None:
        return Finding(name, False, "sin hoja Operating/Model donde ubicar Assumptions")
    assumptions = _section_rows(core, ASSUMPTIONS_TITLE)
    if assumptions is None:
        return Finding(name, False, f"{core.title} no tiene seccion Assumptions")
    first, last = assumptions
    outside: list[str] = []
    for ws in book.formulas.worksheets:
        if ws.title == "Cover":
            continue  # la leyenda de colores muestra la muestra de input
        for cell in _cells(ws):
            if _is_input(cell) and not (ws is core and first <= cell.row <= last):
                outside.append(f"{ws.title}!{cell.coordinate}")
    return Finding(name, not outside,
                   _examples(outside) or f"todos en {core.title} filas {first}-{last} (exenta: leyenda de Cover)")


def _check_s7(book: Book) -> Finding:
    calc = book.formulas.calculation
    mode = getattr(calc, "calcMode", None) if calc is not None else None
    # calcMode ausente en el XML = automatico (default de Excel).
    return Finding("S7 calculo en automatico", mode in ("auto", None), str(mode or "auto (default)"))


def _check_s8(book: Book) -> Finding:
    ok = bool(VERSIONED_NAME.search(book.path.name))
    return Finding("S8 versionado del archivo", ok,
                   book.path.name if ok else f"{book.path.name}: se espera <nombre>_YYYY-MM-DD_v#.xlsx")


def _check_s9(book: Book) -> Finding:
    loose = [n for n in book.formulas.sheetnames
             if n in LOOSE_SHEET_NAMES or n.startswith(LOOSE_SHEET_PREFIXES)]
    return Finding("S9 schedules y estados como bloques, no hojas", not loose,
                   ", ".join(loose) or "ok")


def uncached_verdict(total: int, uncached: Sequence[str]) -> Finding:
    """S10 a partir del conteo: formulas del libro vs formulas sin valor cacheado.

    Ninguna con valor = el libro no se ha recalculado (openpyxl borra la cache
    al guardar): [pendiente], jamas [ok]. Algunas sin valor tras un recalculo =
    ciclo o cadena de dependencias rota: falla.
    """
    name = "S10 sin ciclos ni cadenas rotas"
    if total == 0:
        return Finding(name, True, "n/a (libro sin formulas)")
    if len(uncached) == total:
        return Finding(name, False,
                       f"{total} formulas sin valor calculado: recalcula el libro "
                       "(Excel COM: CalculateFullRebuild + Save) y re-audita", pending=True)
    if uncached:
        return Finding(name, False,
                       f"{len(uncached)} de {total} formulas sin valor tras recalcular: "
                       + _examples(uncached))
    return Finding(name, True, f"{total} formulas con valor calculado")


def _check_s10(book: Book) -> Finding:
    values = load_workbook(str(book.path), data_only=True)
    total = 0
    uncached: list[str] = []
    for ws in book.formulas.worksheets:
        cached = values[ws.title]
        for cell in _formula_cells(ws):
            total += 1
            if cached.cell(row=cell.row, column=cell.column).value is None:
                uncached.append(f"{ws.title}!{cell.coordinate}")
    return uncached_verdict(total, uncached)


def audit_structure(path: str, driver_schedules: Sequence[str] | None = None) -> list[Finding]:
    """Checks S1-S10 sobre un libro. Lectura pura."""
    book = load_book(path)
    return [
        _check_s1(book, driver_schedules),
        *_check_s2_s3(book),
        _check_s4(book),
        _check_s5(book),
        _check_s6(book),
        _check_s7(book),
        _check_s8(book),
        _check_s9(book),
        _check_s10(book),
    ]


# ---------------------------------------------------------------------------
# C10 - desfase del roll de caja (escaneo de formulas)
# ---------------------------------------------------------------------------

_CASH_WORDS = ("efectivo", "caja")
_OPEN_WORDS = ("al inicio", "inicial", "apertura")
_CLOSE_WORDS = ("al cierre", "final")


def _cash_rows(ws: Worksheet) -> tuple[int, int, int | None] | None:
    """Filas (inicio, cierre, caja del balance) del roll de caja, por label."""
    opening = closing = balance = None
    for r in range(1, min(ws.max_row, B._MAX_SCAN_ROWS) + 1):
        label = ws.cell(row=r, column=2).value
        if not isinstance(label, str):
            continue
        low = label.strip().lower()
        if not any(word in low for word in _CASH_WORDS):
            continue
        if opening is None and any(word in low for word in _OPEN_WORDS):
            opening = r
        elif closing is None and any(word in low for word in _CLOSE_WORDS):
            closing = r
        elif balance is None and low.startswith(_CASH_WORDS) and "cambio" not in low:
            balance = r
    return (opening, closing, balance) if opening and closing else None


def _references(formula: str, sheet: str, row: int, column: int) -> bool:
    """True si la formula referencia (sheet, row, column)."""
    return (sheet, row, column) in B._formula_deps(formula, sheet, {sheet})


def audit_cash_roll_formulas(path: str) -> list[Finding]:
    """C10 por escaneo de formulas (no requiere valores calculados).

    (i) cada 'Efectivo al inicio' con formula referencia el CIERRE de la
    columna previa; (ii) cada caja del balance con formula referencia el
    cierre del flujo de su misma columna. El punto (iii) del check
    (rendimientos sobre saldo de apertura) lo cubre F18: violarlo es un ciclo.
    """
    name = "C10 desfase del roll de caja"
    wb = load_workbook(path, data_only=False)
    hits: list[str] = []
    evaluated = 0
    for ws in wb.worksheets:
        columns = sorted(p.column for p in _period_axis(ws))
        rows = _cash_rows(ws)
        if not columns or rows is None:
            continue
        opening, closing, balance = rows
        for previous, column in zip(columns, columns[1:]):
            formula = ws.cell(row=opening, column=column).value
            if isinstance(formula, str) and formula.startswith("="):
                evaluated += 1
                if not _references(formula, ws.title, closing, previous):
                    hits.append(f"{ws.title}!{get_column_letter(column)}{opening}: "
                                "inicio no referencia el cierre de la columna previa")
        if balance is None:
            continue
        for column in columns:
            formula = ws.cell(row=balance, column=column).value
            if isinstance(formula, str) and formula.startswith("="):
                evaluated += 1
                if not _references(formula, ws.title, closing, column):
                    hits.append(f"{ws.title}!{get_column_letter(column)}{balance}: "
                                "caja del balance no referencia el cierre del flujo")
    if hits:
        return [Finding(name, False, _examples(hits))]
    if not evaluated:
        return [Finding(name, True, "n/a (sin roll de caja con formulas identificable)")]
    return [Finding(name, True, f"{evaluated} celdas del roll bien encadenadas")]


# ---------------------------------------------------------------------------
# Reporte y CLI
# ---------------------------------------------------------------------------


def print_report(groups: Sequence[tuple[str, Sequence[Finding]]]) -> int:
    """Imprime el reporte por familia. Exit: 1 fallas; 3 solo pendientes; 0 verde.
    Los avisos se listan y se cuentan, pero no cambian el exit code."""
    for title, findings in groups:
        print(f"-- {title}")
        for finding in findings:
            print(f"{finding.mark} {finding.check}: {finding.detail}")
    everything = [finding for _, findings in groups for finding in findings]
    failures = sum(1 for f in everything if f.blocking)
    pending = sum(1 for f in everything if f.pending)
    warnings = sum(1 for f in everything if f.warning and not f.pending)
    ok = sum(1 for f in everything if f.ok and not f.pending)
    print(f"Resumen: {ok} ok, {failures} fallas, {pending} pendientes, {warnings} avisos")
    if failures:
        print("Veredicto: FALLA")
        return 1
    if pending:
        print("Veredicto: PENDIENTE - no es verde; resuelve los pendientes y re-audita")
        return 3
    print("Veredicto: VERDE (checks por codigo y checks vivos; los D no se evaluan aqui)")
    return 0


def audit_groups(model: str, schedules: Sequence[str] | None,
                 brand: dict[str, str] | None, external: bool,
                 live: bool = True) -> list[tuple[str, Sequence[Finding]]]:
    """Familias S, C por escaneo y F; ``live`` agrega los checks de formula
    viva (con ``--recalc`` ya vienen en el grupo R, sobre el libro calculado)."""
    groups: list[tuple[str, Sequence[Finding]]] = [
        ("Estructurales (S)", audit_structure(model, schedules)),
        ("Contables por escaneo (C)", audit_cash_roll_formulas(model)),
        ("Formato (F)", B.audit_format(model, brand=brand, external=external)),
    ]
    if live:
        groups.append(("Contables de formula viva (C)", [B.live_checks_finding(model)]))
    return groups


def main(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Checks S, C10 y F del modelo por codigo (solo lectura).")
    parser.add_argument("model", help="ruta del modelo .xlsx")
    parser.add_argument("--driver-map", help="driver-map.md: exige un bloque Sch por schedule (S1)")
    parser.add_argument("--brand", help="brand/DESIGN.md: colores de marca aceptados por F4")
    parser.add_argument("--external", action="store_true",
                        help="modelo que no construyo el plugin: la falta de sello (F10) es aviso")
    parser.add_argument("--recalc", action="store_true",
                        help="recalcular una copia con Excel/LibreOffice y auditarla "
                             "(R1, R2 y checks vivos evaluados)")
    args = parser.parse_args(list(argv[1:]))
    if not Path(args.model).is_file():
        print(f"[x] no existe el modelo: {args.model}")
        return 2
    schedules = schedules_from_driver_map(args.driver_map) if args.driver_map else None
    brand = B.load_brand(args.brand) if args.brand else None
    if not args.recalc:
        return print_report(audit_groups(args.model, schedules, brand, args.external))
    import recalc_validate  # import tardio: solo --recalc lanza un motor externo
    with tempfile.TemporaryDirectory(prefix="ra_recalc_") as tmp:
        recalc, recalculated = recalc_validate.validate(Path(args.model), Path(tmp))
        # La copia recalculada conserva el nombre (S8) y trae valores (S10, F19);
        # sus checks vivos ya vienen dentro del grupo R.
        target = str(recalculated) if recalculated is not None else args.model
        groups = audit_groups(target, schedules, brand, args.external,
                              live=recalculated is None)
        return print_report([("Recalculo (R)", recalc), *groups])


if __name__ == "__main__":
    sys.exit(main(sys.argv))
