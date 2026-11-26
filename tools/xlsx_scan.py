"""Lectores del libro para los checks: lo que el audit mide, sin veredicto.

Escaneo de fuentes, rellenos y formatos; columnas de periodo A/E y header
trimestral; grafo de dependencias entre formulas y ciclos (F18); huecos (F15),
continuidad (F11) y series partidas (F12); roll de caja (F19, y
``cash_roll_bounds`` para model_check). Devuelven hechos (listas de
violaciones, columnas, ciclos); los ``Finding`` los arma ``xlsx_audit``.
Todo escaneo respeta ``_MAX_SCAN_ROWS`` x ``_MAX_SCAN_COLS``. Se importa a
traves de la fachada ``xlsx_builder``.
"""

# SIN ``from __future__ import annotations`` a proposito: ver el encabezado
# de ``xlsx_builder.py`` (el builder tambien se carga por ruta).

from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from xlsx_contract import _MAX_SCAN_COLS, _MAX_SCAN_ROWS, Color, NumFmt


def _scan_fonts_fills_formats(ws: Worksheet) -> tuple[set[str], set[str], set[str], set[str]]:
    font_names: set[str] = set()
    font_colors: set[str] = set()
    fills: set[str] = set()
    numfmts: set[str] = set()
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, _MAX_SCAN_ROWS),
                            max_col=min(ws.max_column, _MAX_SCAN_COLS)):
        for cell in row:
            if cell.value is None:
                continue
            font = cell.font
            if font is not None and font.name:
                font_names.add(str(font.name))
                rgb = getattr(font.color, "rgb", None) if font.color else None
                if isinstance(rgb, str):
                    font_colors.add(rgb)
            if cell.fill is not None and cell.fill.patternType == "solid":
                rgb = getattr(cell.fill.fgColor, "rgb", None)
                if isinstance(rgb, str) and rgb != "00000000":
                    fills.add(rgb)
            numfmts.add(cell.number_format)
    return font_names, font_colors, fills, numfmts


import re as _re_mod

_QUARTER_HDR = _re_mod.compile(r"[1-4]Q20\d\d([AE])")


def _period_columns(ws: Worksheet) -> tuple[list[int], list[int]]:
    """Columnas de periodo (histórico A, estimado E) del header de la hoja.

    Reconoce AMBAS formas: años con numfmt 0"A"/0"E" y trimestres como TEXTO
    ('1Q2026E') — sin esto, las hojas trimestrales (Operating) quedan
    invisibles para los checks de series (el punto ciego del smoke #4).
    """
    cols_a: list[int] = []
    cols_e: list[int] = []
    for row in ws.iter_rows(min_row=1, max_row=8,
                            max_col=min(ws.max_column, _MAX_SCAN_COLS)):
        for c in row:
            if c.value is None:
                continue
            if c.number_format == NumFmt.YEAR_A.value:
                cols_a.append(c.column)
            elif c.number_format == NumFmt.YEAR_E.value:
                cols_e.append(c.column)
            elif isinstance(c.value, str):
                m = _QUARTER_HDR.fullmatch(c.value.strip())
                if m:
                    (cols_a if m.group(1) == "A" else cols_e).append(c.column)
        if len(cols_a) + len(cols_e) >= 4:
            break
    return cols_a, cols_e


def _has_quarter_header(ws: Worksheet) -> bool:
    """True si el header (filas 1-8) trae trimestres de texto ('1Q2026E')."""
    for row in ws.iter_rows(min_row=1, max_row=8,
                            max_col=min(ws.max_column, _MAX_SCAN_COLS)):
        for c in row:
            if isinstance(c.value, str) and _QUARTER_HDR.fullmatch(c.value.strip()):
                return True
    return False


def _is_header_row(ws: Worksheet, r: int) -> bool:
    for col in (1, 2):
        f = ws.cell(row=r, column=col).font
        if f is not None and f.bold and (f.size or 0) >= 13:
            return True
    return False


# --- Deteccion de referencias circulares (check F18, sin Excel) ------------

_SHEET_REF = _re_mod.compile(
    r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!"          # hoja opcional
    r"\$?([A-Z]{1,3})\$?(\d+)"
    r"(?::\$?([A-Z]{1,3})\$?(\d+))?"                      # rango opcional
)
_LOCAL_REF = _re_mod.compile(
    r"(?<![A-Z0-9_!$:])\$?([A-Z]{1,3})\$?(\d+)"
    r"(?::\$?([A-Z]{1,3})\$?(\d+))?"
)
_FN_NAME = _re_mod.compile(r"[A-Z][A-Z0-9_.]*\(")


def _formula_deps(formula: str, sheet: str,
                  known_sheets: set[str]) -> set[tuple[str, int, int]]:
    """Celdas (hoja, fila, col) de las que depende una fórmula. Rangos se
    expanden acotados (<= 400 celdas) para no explotar en SUM grandes."""
    deps: set[tuple[str, int, int]] = set()
    body = formula[1:] if formula.startswith("=") else formula
    consumed: list[tuple[int, int]] = []
    for m in _SHEET_REF.finditer(body):
        tgt = (m.group(1) or m.group(2))
        if tgt not in known_sheets:
            continue
        consumed.append((m.start(), m.end()))
        c1 = column_index_from_string(m.group(3))
        r1 = int(m.group(4))
        if m.group(5):
            c2 = column_index_from_string(m.group(5))
            r2 = int(m.group(6))
        else:
            c2, r2 = c1, r1
        if (abs(c2 - c1) + 1) * (abs(r2 - r1) + 1) > 400:
            continue
        for rr in range(min(r1, r2), max(r1, r2) + 1):
            for cc in range(min(c1, c2), max(c1, c2) + 1):
                deps.add((tgt, rr, cc))
    masked = list(body)
    for s, e in consumed:
        for i in range(s, e):
            masked[i] = " "
    rest = "".join(masked)
    rest = _FN_NAME.sub(lambda m: " " * (m.end() - m.start()), rest)
    for m in _LOCAL_REF.finditer(rest):
        c1 = column_index_from_string(m.group(1))
        r1 = int(m.group(2))
        if m.group(3):
            c2 = column_index_from_string(m.group(3))
            r2 = int(m.group(4))
        else:
            c2, r2 = c1, r1
        if (abs(c2 - c1) + 1) * (abs(r2 - r1) + 1) > 400:
            continue
        for rr in range(min(r1, r2), max(r1, r2) + 1):
            for cc in range(min(c1, c2), max(c1, c2) + 1):
                deps.add((sheet, rr, cc))
    return deps


def _find_cycles(wb) -> list[str]:
    """F18: ciclos reales en el grafo de fórmulas (DFS con pila).

    Detecta la circularidad estructural — rendimiento sobre saldo de la misma
    columna, interés sobre deuda de cierre — SIN necesitar Excel: es la causa
    del 'forecast que no calcula'. Ratios legítimos (EBT/EBIT de la misma
    columna) no son ciclos y no aparecen aquí.
    """
    known = set(wb.sheetnames)
    graph: dict[tuple[str, int, int], set[tuple[str, int, int]]] = {}
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(max_row=min(ws.max_row, _MAX_SCAN_ROWS),
                                max_col=min(ws.max_column, _MAX_SCAN_COLS)):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    graph[(name, cell.row, cell.column)] = _formula_deps(
                        v, name, known)
    WHITE, GREY, BLACK = 0, 1, 2
    color: dict[tuple[str, int, int], int] = {}
    cycles: list[str] = []

    def fmt(node) -> str:
        s, r, c = node
        return f"{s}!{get_column_letter(c)}{r}"

    for start in list(graph):
        if color.get(start, WHITE) != WHITE:
            continue
        stack = [(start, iter(graph.get(start, ())))]
        color[start] = GREY
        path = [start]
        while stack:
            node, it = stack[-1]
            advanced = False
            for dep in it:
                if dep not in graph:
                    continue
                st = color.get(dep, WHITE)
                if st == GREY:
                    i = path.index(dep) if dep in path else 0
                    loop = path[i:] + [dep]
                    cycles.append(" -> ".join(fmt(n) for n in loop[:6])
                                  + (" ..." if len(loop) > 6 else ""))
                    if len(cycles) >= 5:
                        return cycles
                elif st == WHITE:
                    color[dep] = GREY
                    path.append(dep)
                    stack.append((dep, iter(graph.get(dep, ()))))
                    advanced = True
                    break
            if not advanced:
                color[node] = BLACK
                stack.pop()
                if path and path[-1] == node:
                    path.pop()
    return cycles


def _gap_violations(ws: Worksheet) -> list[str]:
    """F15: dos venenos de serie.

    (a) HUECOS: fila con contenido en >= mitad de las columnas de periodo pero
        con celdas vacias = formula faltante en un tramo (la NOPAT vacia del
        smoke #4 que convirtio el FCFF en chatarra).
    (b) TEXTO LITERAL ('n/d', 'n/a'...) en columnas de periodo de una fila de
        serie: rompe la cadena de calculo (#VALUE! aguas abajo). El hueco de
        dato es DECISION del analista con gate — 0 explicito con comentario,
        carry-forward del ultimo disponible (formula marcada supuesto), o
        exclusion documentada — jamas texto. Formulas (empiezan con '=') no
        cuentan como texto.
    Bloques de valor unico (una columna) no disparan."""
    ca, ce = _period_columns(ws)
    pcols = sorted(set(ca + ce))
    if len(pcols) < 4:
        return []
    hits: list[str] = []
    for r in range(6, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
        if _is_header_row(ws, r):
            continue
        vals = [ws.cell(row=r, column=c).value for c in pcols]
        filled = sum(1 for v in vals if v is not None)
        label = ws.cell(row=r, column=2).value or ws.cell(row=r, column=1).value
        # Huecos INICIALES permitidos (hasta 4): growth yoy / UDM no tienen
        # ventana previa en los primeros periodos — vacio (no texto) es la
        # convencion. Huecos DESPUES del primer dato = serie rota.
        first_filled = next((i for i, v in enumerate(vals) if v is not None), None)
        if first_filled is not None and first_filled <= 4:
            interior = vals[first_filled:]
            interior_empty = sum(1 for v in interior if v is None)
            if filled >= max(4, len(pcols) // 2) and interior_empty > 0:
                hits.append(f"{ws.title}!fila {r} ({str(label)[:25]}): "
                            f"{interior_empty} huecos")
        elif filled >= max(4, len(pcols) // 2) and filled < len(pcols):
            hits.append(f"{ws.title}!fila {r} ({str(label)[:25]}): "
                        f"{len(pcols) - filled} huecos")
        text_cells = sum(1 for v in vals
                         if isinstance(v, str) and not v.startswith("=")
                         and not _QUARTER_HDR.fullmatch(v.strip()))
        if filled >= 4 and text_cells > 0:
            hits.append(f"{ws.title}!fila {r} ({str(label)[:25]}): "
                        f"{text_cells} celdas de TEXTO en serie (rompe calculo)")
    return hits


def _series_continuity_violations(ws: Worksheet) -> list[str]:
    """F11: rows that look like input series must span the WHOLE horizon.

    Period columns = columns whose header cell uses the 0"A"/0"E" year formats.
    Any row with >= 3 input-filled cells among period columns is a series row;
    a series row with empty period cells means history/forecast got split or
    history was left unpopulated (the AAPL smoke-test failure pattern).
    """
    ca, ce = _period_columns(ws)   # anios A/E numfmt Y trimestres de texto
    period_cols = sorted(set(ca + ce))
    if len(period_cols) < 4:
        return []
    violations: list[str] = []
    for r in range(1, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
        if _is_header_row(ws, r):
            continue
        filled_inputs = 0
        vals = []
        for col in period_cols:
            cell = ws.cell(row=r, column=col)
            is_input_fill = (cell.fill is not None
                             and cell.fill.patternType == "solid"
                             and getattr(cell.fill.fgColor, "rgb", None)
                             == Color.INPUT_FILL.value)
            if cell.value is not None and is_input_fill:
                filled_inputs += 1
            vals.append(cell.value)
        if filled_inputs < 3:
            continue
        # Mismo carve-out que F15: hasta 4 huecos INICIALES son legitimos
        # (growth yoy / UDM sin ventana previa van VACIOS, no con texto);
        # un hueco DESPUES del primer dato si es serie rota.
        first = next((i for i, v in enumerate(vals) if v is not None), None)
        if first is None:
            continue
        interior_empty = sum(1 for v in vals[first:] if v is None)
        if first <= 4:
            if interior_empty > 0:
                violations.append(
                    f"{ws.title}!fila {r} ({interior_empty} celdas vacias)")
        elif interior_empty > 0 or first > 0:
            empties = sum(1 for v in vals if v is None)
            violations.append(f"{ws.title}!fila {r} ({empties} celdas vacias)")
    return violations


# El roll de caja se reconoce por label. Inicio y cierre deben EMPEZAR con
# estas palabras: un supuesto como "Rendimiento sobre la caja de apertura" o
# "Tasa de D&A (% del PP&E neto de apertura)" menciona apertura y no es el roll.
CASH_LABEL_PREFIXES = ("efectivo", "caja")
_CASH_OPEN_WORDS = ("al inicio", "inicial", "apertura")
_CASH_CLOSE_WORDS = ("al cierre", "final")


def cash_roll_bounds(ws: Worksheet) -> Optional[tuple[int, int]]:
    """Filas (inicio, cierre) del roll de caja de la hoja, o None.

    Regla unica para F19 (valores) y C10 (formulas, en model_check.py).
    """
    r_open = r_close = None
    for r in range(1, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
        lab = str(ws.cell(row=r, column=2).value or "").strip().lower()
        if not lab.startswith(CASH_LABEL_PREFIXES):
            continue
        if r_open is None and any(word in lab for word in _CASH_OPEN_WORDS):
            r_open = r
        elif r_close is None and any(word in lab for word in _CASH_CLOSE_WORDS):
            r_close = r
    return (r_open, r_close) if r_open and r_close else None


def _cash_roll_violations(path: str) -> Optional[tuple[list[str], int, int]]:
    """F19: el roll de caja debe CERRAR en todas las columnas.

    (i) inicio(t) = cierre(t-1)  — el desfase temporal;
    (ii) inicio(t) + cambio neto(t) = cierre(t) — la identidad del roll.

    (ii) es el que caza un flujo de efectivo INCOMPLETO: si el CF omite una
    seccion (p. ej. el movimiento de valores negociables, el mayor flujo
    despues del operativo en emisoras con tesoreria grande), el balance puede
    seguir cuadrando con caja observada y el tie-out BS<->CF tambien — ambos
    leen el MISMO observado — mientras el roll no cierra en silencio. Es el
    bug del smoke #5: 38 trimestres historicos con el roll roto.
    """
    wbf = load_workbook(path, data_only=False)
    wbv = load_workbook(path, data_only=True)
    hits: list[str] = []
    found_any = False
    evaluated = uncalculated = 0
    for name in wbf.sheetnames:
        wsf, wsv = wbf[name], wbv[name]
        ca, ce = _period_columns(wsf)
        pcols = sorted(set(ca + ce))
        if len(pcols) < 4:
            continue
        bounds = cash_roll_bounds(wsf)
        r_open, r_close = bounds if bounds else (None, None)
        r_change = None
        for r in range(1, min(wsf.max_row, _MAX_SCAN_ROWS) + 1):
            lab = str(wsf.cell(row=r, column=2).value or "").lower()
            if ("cambio neto" in lab or "aumento (disminucion)" in lab
                    or "variacion neta" in lab):
                r_change = r
                break
        if not (r_open and r_close and r_change):
            continue
        found_any = True
        for i, c in enumerate(pcols):
            col = get_column_letter(c)
            op = wsv.cell(row=r_open, column=c).value
            cl = wsv.cell(row=r_close, column=c).value
            ch = wsv.cell(row=r_change, column=c).value
            if not all(isinstance(x, (int, float)) for x in (op, ch, cl)):
                if any(str(wsf.cell(row=rr, column=c).value or "").startswith("=")
                       for rr in (r_open, r_close, r_change)):
                    uncalculated += 1        # formula sin valor: libro sin recalcular
            else:
                evaluated += 1
                tol = max(1.0, abs(cl) * 1e-6)
                if abs(op + ch - cl) > tol:
                    hdr = wsf.cell(row=4, column=c).value
                    hits.append(f"{name}!{col} ({hdr}): inicio+cambio != cierre "
                                f"(brecha {op + ch - cl:,.0f})")
            if i > 0:
                prev_cl = wsv.cell(row=r_close, column=pcols[i - 1]).value
                if (isinstance(op, (int, float))
                        and isinstance(prev_cl, (int, float))
                        and abs(op - prev_cl) > max(1.0, abs(op) * 1e-6)):
                    hdr = wsf.cell(row=4, column=c).value
                    hits.append(f"{name}!{col} ({hdr}): inicio != cierre previo")
    return (hits, evaluated, uncalculated) if found_any else None


def _split_series_violations(ws: Worksheet) -> list[str]:
    """F12: label contiene 'forecast' u 'historico' y solo su mitad esta llena."""
    period_cols_a, period_cols_e = _period_columns(ws)
    if not period_cols_a or not period_cols_e:
        return []
    hits: list[str] = []
    for r in range(1, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
        label = ws.cell(row=r, column=1).value
        if not isinstance(label, str) or label.strip().lower() == "x":
            label = ws.cell(row=r, column=2).value
        if not isinstance(label, str):
            continue
        low = label.lower()
        is_fc = "forecast" in low
        is_hist = "historico" in low or "histórico" in low
        if not (is_fc or is_hist):
            continue
        a_vals = sum(1 for c in period_cols_a if ws.cell(row=r, column=c).value is not None)
        e_vals = sum(1 for c in period_cols_e if ws.cell(row=r, column=c).value is not None)
        if is_fc and e_vals >= 2 and a_vals == 0:
            hits.append(f"{ws.title}!fila {r} (solo forecast)")
        if is_hist and a_vals >= 2 and e_vals == 0:
            hits.append(f"{ws.title}!fila {r} (solo historico)")
    return hits
