"""Audit de formato del xlsx: checks F1-F21 y checks de formula viva.

``audit_format`` corre los checks F contra cualquier libro, lo haya
construido el plugin o no; ``live_checks_finding`` lee el estado de las tabs
Checks y Cover en un libro ya recalculado; ``_print_report`` es la unica regla
de reporte a consola (y de exit code) que comparten xlsx_builder,
model_check y recalc_validate. Los hechos del libro salen de ``xlsx_scan``.
Se importa a traves de la fachada ``xlsx_builder``.
"""

# SIN ``from __future__ import annotations`` a proposito: ver el encabezado
# de ``xlsx_builder.py`` (el builder tambien se carga por ruta).

from dataclasses import dataclass
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from xlsx_contract import (
    _MAX_SCAN_COLS,
    _MAX_SCAN_ROWS,
    BUILDER_STAMP_KEY,
    FONT_NAME,
    FROZEN_SHEET_PREFIXES,
    JUNK_SHEET_NAMES,
    REQUIRED_RATIO_LABELS,
    Color,
    NumFmt,
)
from xlsx_formulas import lint_formula
from xlsx_scan import (
    _cash_roll_violations,
    _find_cycles,
    _gap_violations,
    _is_header_row,
    _period_columns,
    _scan_fonts_fills_formats,
    _series_continuity_violations,
    _split_series_violations,
)


# ---------------------------------------------------------------------------
# Checks F — format audit (deterministic, any workbook)
# ---------------------------------------------------------------------------

_ALLOWED_FONT_COLORS = {c.value for c in (
    Color.INPUT_BLUE, Color.LINK_GREEN, Color.WARN_RED, Color.WHITE, Color.BLACK)}
_ALLOWED_FONT_COLORS.add("FF333333")  # near-black tolerated
_ALLOWED_FILLS = {c.value for c in (
    Color.DARK_BLUE, Color.LIGHT_BLUE, Color.INPUT_FILL, Color.SCENARIO_FILL)}
_ALLOWED_NUMFMTS = {f.value for f in NumFmt}


@dataclass(frozen=True)
class Finding:
    """Resultado de un check F. ``pending`` = no se pudo evaluar (p. ej. libro
    sin valores calculados): NO es verde ni falla; se reporta [pendiente].
    ``warning`` = el check NO se cumple pero la doctrina lo declara "aviso,
    no bloqueo": se reporta [aviso], se lista siempre y no cambia el exit
    code. ``ok`` sigue diciendo el hecho (False); ``blocking`` dice si tumba
    el veredicto."""

    check: str
    ok: bool
    detail: str
    pending: bool = False
    warning: bool = False

    @property
    def blocking(self) -> bool:
        """True si el hallazgo es una falla que tumba el veredicto."""
        return not self.ok and not self.pending and not self.warning

    @property
    def mark(self) -> str:
        """Marca de consola: una sola regla para todo reporte."""
        if self.pending:
            return "[pendiente]"
        if self.ok:
            return "[ok]"
        return "[aviso]" if self.warning else "[x]"


def audit_format(path: str, brand: Optional[dict[str, str]] = None,
                 external: bool = False) -> list[Finding]:
    """Run checks F (F1-F21) on a workbook. Pure read; returns findings.

    ``brand``: output of load_brand(brand/DESIGN.md) — its values extend the
    fill whitelist (F4) so a branded model audits green with its own DESIGN.md.
    ``external``: quien audita declara que el modelo NO lo construyo el plugin
    (modo auditoria de /model-check sobre un modelo propio del analista).
    Entonces la falta de sello (F10) es [aviso]; en el pipeline del plugin,
    donde el sello es obligatorio, sigue siendo falla.
    """
    allowed_fills = _ALLOWED_FILLS | set((brand or {}).values())
    wb = load_workbook(path, data_only=False)
    findings: list[Finding] = []
    visible = [wb[n] for n in wb.sheetnames if wb[n].sheet_state == "visible"]

    # F1 gridlines off everywhere
    bad = [ws.title for ws in visible if ws.sheet_view.showGridLines in (True, None)]
    findings.append(Finding("F1 gridlines off", not bad, ", ".join(bad) or "todas ok"))

    # F20 cobertura: los checks escanean hasta _MAX_SCAN_ROWS x _MAX_SCAN_COLS.
    # Antes el tope era 400 filas x 40 columnas y lo que quedaba abajo (los
    # schedules de Operating) pasaba sin auditar y en verde.
    too_big = [f"{ws.title} ({ws.max_row}x{ws.max_column})" for ws in visible
               if ws.max_row > _MAX_SCAN_ROWS or ws.max_column > _MAX_SCAN_COLS]
    findings.append(Finding(
        "F20 cobertura del audit", not too_big,
        ("hojas mas grandes que el limite, auditadas a medias: " + ", ".join(too_big))
        if too_big else f"todo dentro de {_MAX_SCAN_ROWS} filas x {_MAX_SCAN_COLS} columnas"))

    all_fonts: set[str] = set()
    all_colors: set[str] = set()
    all_fills: set[str] = set()
    all_fmts: set[str] = set()
    for ws in visible:
        fn, fc, fl, nf = _scan_fonts_fills_formats(ws)
        all_fonts |= fn
        all_colors |= fc
        all_fills |= fl
        all_fmts |= nf

    # F2 single standard font family
    alien_fonts = sorted(all_fonts - {FONT_NAME})
    findings.append(Finding("F2 fuente estandar", not alien_fonts,
                            ", ".join(alien_fonts) or FONT_NAME))
    # F3 font colors within whitelist
    alien_colors = sorted(all_colors - _ALLOWED_FONT_COLORS)
    findings.append(Finding("F3 colores de fuente", not alien_colors,
                            ", ".join(alien_colors) or "whitelist ok"))
    # F4 fills within palette (+ brand slots if DESIGN.md provided)
    alien_fills = sorted(all_fills - allowed_fills)
    findings.append(Finding("F4 paleta de fills", not alien_fills,
                            ", ".join(alien_fills) or "whitelist ok"))
    # F5 number formats within whitelist
    alien_fmts = sorted(all_fmts - _ALLOWED_NUMFMTS)
    findings.append(Finding("F5 formatos numericos", not alien_fmts,
                            "; ".join(alien_fmts[:6]) or "whitelist ok"))
    # F6 frozen panes on data sheets
    no_freeze = [ws.title for ws in visible
                 if ws.title.startswith(FROZEN_SHEET_PREFIXES) and not ws.freeze_panes]
    findings.append(Finding("F6 freeze panes", not no_freeze,
                            ", ".join(no_freeze) or "ok"))
    # F7 outline grouping: en la tab Model (secciones apiladas) o en Schedules
    host = next((ws for ws in visible if ws.title in ("Model", "Schedules")), None)
    hosts = [ws for ws in visible
             if ws.title in ("Operating", "Annual", "Model", "Schedules")]
    if not hosts:
        findings.append(Finding("F7 outline por seccion", False,
                                "sin hojas Operating/Annual/Model/Schedules"))
    else:
        # Por SECCION (marcador 'x' en col A + bold >=13 en col B): cada
        # seccion con contenido debe tener filas agrupadas — agrupar solo
        # algunas secciones (bug del smoke 2026-08-31) FALLA.
        ungrouped: list[str] = []
        total_grouped = 0
        for host in hosts:
            total_grouped += sum(1 for d in host.row_dimensions.values()
                                 if d.outlineLevel)
            headers: list[int] = []
            for r in range(1, min(host.max_row, _MAX_SCAN_ROWS) + 1):
                a = host.cell(row=r, column=1).value
                b = host.cell(row=r, column=2)
                if (isinstance(a, str) and a.strip().lower() == "x"
                        and b.font is not None and b.font.bold
                        and (b.font.size or 0) >= 13):
                    headers.append(r)
            for i, hr in enumerate(headers):
                end = (headers[i + 1] - 1 if i + 1 < len(headers)
                       else min(host.max_row, _MAX_SCAN_ROWS))
                content = [r for r in range(hr + 1, end + 1)
                           if any(host.cell(row=r, column=c).value is not None
                                  for c in range(2, min(host.max_column, _MAX_SCAN_COLS) + 1))]
                if not content:
                    continue
                grouped = sum(1 for r in content
                              if r in host.row_dimensions
                              and host.row_dimensions[r].outlineLevel)
                if grouped == 0:
                    label = host.cell(row=hr, column=2).value
                    ungrouped.append(f"{host.title}!fila {hr} ({str(label)[:25]})")
        ok = total_grouped > 0 and not ungrouped
        detail = (f"{total_grouped} filas agrupadas"
                  + (f"; secciones SIN outline: {', '.join(ungrouped[:5])}"
                     if ungrouped else ""))
        findings.append(Finding("F7 outline por seccion", ok, detail))
    # F8 no junk sheets, no Sch_* sheets
    junk = [n for n in wb.sheetnames if n in JUNK_SHEET_NAMES or n.startswith("Sch_")]
    findings.append(Finding("F8 sin hojas basura/Sch_*", not junk,
                            ", ".join(junk) or "ok"))
    # F9 A/E period formats present somewhere
    has_ae = any(f in all_fmts for f in (NumFmt.YEAR_A.value, NumFmt.YEAR_E.value))
    findings.append(Finding("F9 headers de periodo A/E", has_ae,
                            "presentes" if has_ae else "sin formato 0\"A\"/0\"E\""))
    # F10 builder stamp
    kw = wb.properties.keywords or ""
    stamped = BUILDER_STAMP_KEY in kw
    findings.append(Finding("F10 sello del builder", stamped,
                            kw if stamped else "sin sello (no construido por xlsx_builder)",
                            warning=external and not stamped))
    # F11 series continuity: input-series rows span the whole horizon
    all_violations: list[str] = []
    for ws in visible:
        all_violations.extend(_series_continuity_violations(ws))
    findings.append(Finding("F11 continuidad de series", not all_violations,
                            "; ".join(all_violations[:8]) or
                            "filas de input completas en todo el horizonte"))
    # F12 series partidas: fila etiquetada "forecast"/"historico" con solo una
    # mitad del horizonte poblada = la serie se partio en dos filas (violacion
    # de "una serie = una fila"; el patron exacto del smoke AAPL)
    split_hits: list[str] = []
    for ws in visible:
        split_hits.extend(_split_series_violations(ws))
    findings.append(Finding("F12 sin series partidas", not split_hits,
                            "; ".join(split_hits[:8]) or
                            "ninguna fila hist/forecast partida"))
    # F15 series sin huecos: fila con contenido en >= mitad de las columnas
    # de periodo pero con celdas vacias = formula faltante en un tramo (la
    # NOPAT vacia en forecast del smoke #4 — S5 hecho codigo).
    gap_hits: list[str] = []
    for ws in visible:
        if ws.title in ("Operating", "Annual", "Model", "Schedules", "IS",
                        "BS", "CF", "Ratios"):
            gap_hits.extend(_gap_violations(ws))
    findings.append(Finding("F15 series sin huecos", not gap_hits,
                            "; ".join(gap_hits[:8]) or
                            "sin huecos en filas de serie"))
    # F16 respiro tipografico: fila EN BLANCO antes de cada header/sub-header
    # (salvo headers consecutivos y el tope de la hoja) — elegancia del modelo.
    breath_hits: list[str] = []
    for ws in visible:
        if ws.title not in ("Operating", "Annual", "Model", "Schedules",
                            "Summary"):
            continue
        for r in range(6, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
            if not _is_header_row(ws, r):
                continue
            prev = r - 1
            if _is_header_row(ws, prev):
                continue  # headers consecutivos: sin respiro
            prev_has_content = any(
                ws.cell(row=prev, column=c).value is not None
                for c in range(1, min(ws.max_column, _MAX_SCAN_COLS) + 1))
            label = ws.cell(row=r, column=2).value
            if prev_has_content:
                breath_hits.append(f"{ws.title}!fila {r} ({str(label)[:22]})")
            elif (prev in ws.row_dimensions
                  and ws.row_dimensions[prev].outlineLevel):
                # El respiro existe pero vive DENTRO del grupo: al colapsar
                # desaparece y los headers quedan pegados.
                breath_hits.append(
                    f"{ws.title}!fila {prev} (respiro agrupado, se pierde al colapsar)")
    findings.append(Finding("F16 respiro antes de headers", not breath_hits,
                            "; ".join(breath_hits[:6]) or
                            "headers con fila en blanco previa"))
    # F17 FORECAST COMPLETO (mandato duro): toda fila con historico (>=3
    # celdas en columnas A) debe tener el tramo E COMPLETO — cero vacios.
    # "Los forecast deben tener todas las formulas completas" (Alan).
    fc_hits: list[str] = []
    for ws in visible:
        if ws.title not in ("Operating", "Annual", "Model", "Schedules", "IS",
                            "BS", "CF", "Ratios"):
            continue
        ca_, ce_ = _period_columns(ws)
        if len(ca_) < 2 or len(ce_) < 2:
            continue
        for r in range(6, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
            if _is_header_row(ws, r):
                continue
            # Exencion explicita: serie que la EMISORA descontinuo (deja de
            # reportar el dato) o linea declarada como deuda de driver. No se
            # inventa un forecast de un dato que la fuente ya no publica; la
            # fila queda como referencia historica documentada en su label.
            lab_txt = str(ws.cell(row=r, column=2).value or "").lower()
            if ("descontinuad" in lab_txt or "deuda de driver" in lab_txt
                    or "sin driver" in lab_txt):
                continue
            a_filled = sum(1 for c in ca_
                           if ws.cell(row=r, column=c).value is not None)
            # Serie TEMPORAL = poblada en al menos la mitad del historico.
            # Un bloque de valor unico horizontal (beta levered / D/E / beta
            # unlevered de cada comp en el bloque Hamada) ocupa 2-3 columnas
            # y NO es una serie: exigirle forecast seria absurdo.
            if a_filled < max(3, (len(ca_) + 1) // 2):
                continue
            e_empty = sum(1 for c in ce_
                          if ws.cell(row=r, column=c).value is None)
            if e_empty > 0:
                label = (ws.cell(row=r, column=2).value
                         or ws.cell(row=r, column=1).value)
                fc_hits.append(f"{ws.title}!fila {r} ({str(label)[:25]}): "
                               f"{e_empty}/{len(ce_)} E vacias")
    findings.append(Finding("F17 forecast completo", not fc_hits,
                            "; ".join(fc_hits[:8]) or
                            "todo el tramo forecast con formula"))
    # F13 completitud + UNICIDAD de Ratios: el set completo presente, y cada
    # razon UNA sola vez por hoja — un label duplicado delata secciones
    # "Ratios historico" / "Ratios forecast" partidas (bug del smoke #3);
    # la serie completa vive en UNA fila.
    counts: dict[tuple[str, str], int] = {}
    with_formula: set[str] = set()
    for ws in visible:
        if ws.title not in ("Operating", "Annual", "Model", "Ratios"):
            continue
        for r in range(1, min(ws.max_row, _MAX_SCAN_ROWS) + 1):
            a_cell = ws.cell(row=r, column=1)
            b_cell = ws.cell(row=r, column=2)
            # Saltar encabezados de seccion/sub-seccion (mismo criterio que
            # F7: bold >=13) — sus titulos contienen nombres de razones
            # ("ROIC y economic profit") y NO son filas de razon.
            def _is_header(c) -> bool:
                return (c.font is not None and c.font.bold
                        and (c.font.size or 0) >= 13)
            if _is_header(a_cell) or _is_header(b_cell):
                continue
            label = b_cell.value
            if not isinstance(label, str) or not label.strip():
                label = a_cell.value
            if not isinstance(label, str):
                continue
            low = label.strip().lower()
            if low == "x":
                continue
            # Comparacion por PREFIJO: "CCC (DSO+DIO-DPO)" cuenta solo para
            # CCC, no para DSO/DIO/DPO (que aparecen como substring).
            for req in REQUIRED_RATIO_LABELS:
                if low.startswith(req.lower()):
                    key_ = (ws.title, req)
                    counts[key_] = counts.get(key_, 0) + 1
                    # Una etiqueta sin formulas no es una razon: el demo
                    # anterior pasaba F13 escribiendo solo los 25 nombres.
                    if any(isinstance(v, str) and v.startswith("=")
                           for v in (ws.cell(row=r, column=c).value for c in
                                     range(3, min(ws.max_column, _MAX_SCAN_COLS) + 1))):
                        with_formula.add(req)
    labels_found = {label for (_, label) in counts}
    missing_ratios = [x for x in REQUIRED_RATIO_LABELS if x not in labels_found]
    no_formula = [x for x in REQUIRED_RATIO_LABELS
                  if x in labels_found and x not in with_formula]
    dup_ratios = [f"{sheet}:{label}" for (sheet, label), n in counts.items()
                  if n >= 2]
    ok13 = not missing_ratios and not dup_ratios and not no_formula
    parts13: list[str] = []
    if missing_ratios:
        parts13.append("faltan: " + ", ".join(missing_ratios[:8])
                       + (" ..." if len(missing_ratios) > 8 else ""))
    if no_formula:
        parts13.append("sin formulas (solo etiqueta): " + ", ".join(no_formula[:8])
                       + (" ..." if len(no_formula) > 8 else ""))
    if dup_ratios:
        parts13.append("secciones de Ratios PARTIDAS (label duplicado): "
                       + ", ".join(dup_ratios[:5]))
    detail13 = ("; ".join(parts13) if parts13 else
                f"{len(REQUIRED_RATIO_LABELS)} razones con formula, sin duplicados")
    findings.append(Finding("F13 Ratios completa y unica", ok13, detail13))
    # F14 columnas trimestrales estimadas: si el sello dice periodicidad con
    # trimestres, el header debe traer columnas #Q20yyE (el contrato 1a que el
    # rebuild v3 se salto). Sin sello de periodicidad: n/a (modelo externo).
    import re as _re
    m = _re.search(r"periodicity=([a-z_]+)", kw)
    if m and m.group(1) in ("annual_plus_quarterly", "quarterly"):
        # (i) Operating (o Model legacy): el modelo SE CONSTRUYE sobre
        # trimestres — >=4 A y >=4 E en el header.
        q_actual = q_est = 0
        for ws in visible:
            if ws.title not in ("Operating", "Model"):
                continue
            for row in ws.iter_rows(min_row=1, max_row=8,
                                    max_col=min(ws.max_column, _MAX_SCAN_COLS)):
                for c in row:
                    if isinstance(c.value, str):
                        v = c.value.strip()
                        if _re.fullmatch(r"[1-4]Q20\d\dA", v):
                            q_actual += 1
                        elif _re.fullmatch(r"[1-4]Q20\d\dE", v):
                            q_est += 1
        problems: list[str] = []
        if q_actual < 4 or q_est < 4:
            problems.append(f"Operating: {q_actual} trimestres A / {q_est} E (min 4 y 4)")
        # (ii) Annual: solo anios FY en header (cero '#Q') y CERO inputs —
        # los agregados jamas se teclean (C8 estructural).
        annual = next((ws for ws in visible if ws.title == "Annual"), None)
        if annual is not None:
            q_in_annual = 0
            inputs_in_annual = 0
            for row in annual.iter_rows(min_row=1, max_row=8,
                                        max_col=min(annual.max_column, _MAX_SCAN_COLS)):
                for c in row:
                    if isinstance(c.value, str) and _re.fullmatch(
                            r"[1-4]Q20\d\d[AE]", c.value.strip()):
                        q_in_annual += 1
            for row in annual.iter_rows(max_row=min(annual.max_row, _MAX_SCAN_ROWS),
                                        max_col=min(annual.max_column, _MAX_SCAN_COLS)):
                for c in row:
                    if (c.fill is not None and c.fill.patternType == "solid"
                            and getattr(c.fill.fgColor, "rgb", None)
                            == Color.INPUT_FILL.value and c.value is not None):
                        inputs_in_annual += 1
            if q_in_annual:
                problems.append(f"Annual: {q_in_annual} columnas #Q (debe ser FY-solo)")
            if inputs_in_annual:
                problems.append(f"Annual: {inputs_in_annual} celdas de INPUT (agregados no se teclean)")
        findings.append(Finding(
            "F14 modelo trimestral-nativo", not problems,
            "; ".join(problems) if problems else
            f"Operating: {q_actual} A / {q_est} E; Annual limpio"))
    else:
        findings.append(Finding("F14 modelo trimestral-nativo", True,
                                "n/a (sin sello de periodicidad trimestral)"))
    # F19 roll de caja cerrado: inicio(t) + cambio(t) = cierre(t) e
    # inicio(t) = cierre(t-1), en TODAS las columnas incluido el historico.
    # Requiere valores calculados; sin ellos reporta [pendiente de recalc].
    try:
        roll = _cash_roll_violations(path)
    except Exception as exc:  # noqa: BLE001
        findings.append(Finding("F19 roll de caja cerrado", False,
                                f"no evaluado ({type(exc).__name__})", pending=True))
    else:
        if roll is None:
            findings.append(Finding("F19 roll de caja cerrado", True,
                                    "n/a (sin filas de roll identificables)"))
        else:
            roll_hits, evaluated, uncalculated = roll
            if roll_hits:
                findings.append(Finding("F19 roll de caja cerrado", False,
                                        "; ".join(roll_hits[:6])))
            elif uncalculated or not evaluated:
                # openpyxl guarda formulas SIN valores calculados: sin recalculo
                # no hay nada que comparar. Antes esto salia [ok].
                findings.append(Finding(
                    "F19 roll de caja cerrado", False,
                    f"{uncalculated} columnas con formulas sin valor calculado: "
                    "recalcula el libro (Excel COM: CalculateFullRebuild + Save) "
                    "y re-audita", pending=True))
            else:
                findings.append(Finding("F19 roll de caja cerrado", True,
                                        f"inicio+cambio=cierre e inicio=cierre previo "
                                        f"({evaluated} columnas)"))
    # F18 sin referencias circulares: grafo de dependencias + DFS. Caza la
    # causa raiz del "forecast que no calcula" (rendimiento/interes sobre
    # saldo de la MISMA columna) SIN necesitar Excel; ratios legitimos de la
    # misma columna no son ciclos y no aparecen.
    try:
        cycles = _find_cycles(wb)
    except Exception as exc:  # noqa: BLE001 - nunca tumbar el audit completo
        findings.append(Finding("F18 sin referencias circulares", False,
                                f"no evaluado ({type(exc).__name__})", pending=True))
    else:
        findings.append(Finding("F18 sin referencias circulares", not cycles,
                                ("CICLOS: " + " | ".join(cycles[:3]))
                                if cycles else "grafo de formulas aciclico"))
    # F21 formulas aceptables para Excel: aridad, parentesis, funciones sin
    # prefijo _xlfn. y operaciones de arreglo en formula normal. Sin esto el
    # audit daba verde sobre un libro que Excel abre en modo reparacion.
    bad_formulas: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(max_row=min(ws.max_row, _MAX_SCAN_ROWS),
                                max_col=min(ws.max_column, _MAX_SCAN_COLS)):
            for cell in row:
                value = cell.value
                is_array = isinstance(value, ArrayFormula)
                text = value.text if is_array else value
                for problem in lint_formula(text, array=is_array):
                    bad_formulas.append(f"{ws.title}!{cell.coordinate}: {problem}")
    findings.append(Finding(
        "F21 formulas aceptables para Excel", not bad_formulas,
        ("; ".join(bad_formulas[:6])
         + (f" (+{len(bad_formulas) - 6} mas)" if len(bad_formulas) > 6 else ""))
        if bad_formulas else "aridad, parentesis, prefijos y arreglos ok"))
    return findings


# --- Checks vivos de la tab Checks (formulas) --------------------------------

# Tabs con checks de formula viva: la lista (Checks) y la celda unica de
# error que los agrega (Cover).
LIVE_CHECK_SHEETS = ("Checks", "Cover")
_EXCEL_ERRORS = ("#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")


def live_checks_finding(path: str) -> Finding:
    """Estado de los checks de FORMULA VIVA de la tab ``Checks`` (C1-C5, C7...).

    Un check que nunca se evaluo no es un check verde. openpyxl guarda las
    formulas SIN valor calculado: mientras el libro no se recalcule con un
    motor real (``tools/recalc_validate.py``), estos checks estan SIN evaluar y
    el resumen por codigo no los cubre -> [pendiente]. Con valores calculados:
    cualquier ``Error`` o error de Excel (#VALUE!, #NAME?...) es falla.
    """
    name = "Checks vivos (tabs Checks y Cover)"
    formulas = load_workbook(path, data_only=False)
    sheets = [n for n in LIVE_CHECK_SHEETS if n in formulas.sheetnames]
    if not sheets:
        return Finding(name, True, "n/a (sin tab Checks ni Cover)")
    cached_book = load_workbook(path, data_only=True)
    total = 0
    unevaluated: list[str] = []
    failing: list[str] = []
    for sheet in sheets:
        ws, values = formulas[sheet], cached_book[sheet]
        for row in ws.iter_rows(max_row=min(ws.max_row, _MAX_SCAN_ROWS),
                                max_col=min(ws.max_column, _MAX_SCAN_COLS)):
            for cell in row:
                v = cell.value
                if not (isinstance(v, ArrayFormula)
                        or (isinstance(v, str) and v.startswith("="))):
                    continue
                total += 1
                cached = values.cell(row=cell.row, column=cell.column).value
                label = str(ws.cell(row=cell.row, column=2).value or "")[:40]
                where = f"{sheet}!{cell.coordinate}" + (f" ({label})" if label else "")
                if cached is None:
                    unevaluated.append(where)
                elif isinstance(cached, str) and (
                        cached.strip().upper().startswith("ERROR")
                        or cached.strip() in _EXCEL_ERRORS):
                    failing.append(f"{where} = {cached}")
    if total == 0:
        return Finding(name, True, "n/a (tabs Checks y Cover sin formulas vivas)")
    if failing:
        return Finding(name, False, "; ".join(failing[:8])
                       + (f" (+{len(failing) - 8} mas)" if len(failing) > 8 else ""))
    if unevaluated:
        return Finding(
            name, False,
            f"{len(unevaluated)} de {total} checks vivos SIN evaluar (formulas sin valor "
            "calculado): este reporte NO los cubre. Recalcula y valida con "
            "python tools/recalc_validate.py <modelo.xlsx>", pending=True)
    return Finding(name, True, f"{total} checks vivos evaluados, ninguno en Error")


def _print_report(findings: Iterable[Finding],
                  live: Optional[Finding] = None,
                  label: str = "Resumen F") -> int:
    """Imprime el reporte. Exit: 1 si hay fallas; 3 si solo hay pendientes.
    Los avisos se listan y se cuentan, pero no cambian el exit code.

    ``live``: estado de los checks de formula viva (``live_checks_finding``).
    No es un check F, pero entra al veredicto: un libro con checks vivos sin
    evaluar no puede salir verde solo porque su formato esta bien."""
    findings = list(findings)
    for f in findings:
        print(f"{f.mark} {f.check}: {f.detail}")
    if live is not None:
        print(f"{live.mark} {live.check}: {live.detail}")
        findings.append(live)
    failures = sum(1 for f in findings if f.blocking)
    pending = sum(1 for f in findings if f.pending)
    warnings = sum(1 for f in findings if f.warning and not f.pending)
    ok_count = sum(1 for f in findings if f.ok and not f.pending)
    print(f"{label}: {ok_count} ok, {failures} fallas, {pending} pendientes, "
          f"{warnings} avisos")
    if failures:
        return 1
    if pending:
        print("Veredicto: PENDIENTE - no es verde; resuelve los pendientes y re-audita")
        return 3
    return 0
