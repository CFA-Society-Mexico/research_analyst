"""Deterministic xlsx style/scaffold engine for the research_analyst plugin.

Doctrine: format is CODE, not model judgement. model-standards builds every
workbook THROUGH this module (never raw openpyxl styling), and /model-check
runs ``audit_format`` (checks F) against any workbook, plugin-built or not.

Empirical basis: CFI workbook corpus (3-Statement Model Complete, AMZN
Advanced case, Valuation Model, template library) extracted 2026-08-30.

Console output policy: ASCII only ([ok]/[x], no unicode symbols).

Usage as CLI:
    python tools/xlsx_builder.py audit <path.xlsx>     -> run checks F
    python tools/xlsx_builder.py demo  <path.xlsx>     -> build a v3 skeleton (self-test)

Exit codes: 0 = verde; 1 = alguna falla; 3 = sin fallas pero con checks
[pendiente] (p. ej. libro sin valores calculados: recalcular y re-auditar).
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from enum import Enum, auto
from typing import Iterable, Optional

from datetime import date

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------------------------------------------------------
# Contract constants (the ONLY allowed values; checks F enforce the whitelists)
# ---------------------------------------------------------------------------

BUILDER_STAMP_KEY = "research_analyst_builder"
BUILDER_STAMP_VALUE = "xlsx_builder v1"

FONT_NAME = "Aptos Narrow"


class Color(str, Enum):
    """Palette (ARGB). Identidad propia del plugin (2026-08-31): azul oscuro +
    azul claro. Los valores CFI legacy se conservan como referencia historica
    pero YA NO estan en la whitelist."""

    DARK_BLUE = "FF1F4E79"   # barra de marca + bandas de seccion (texto blanco)
    LIGHT_BLUE = "FFBDD7EE"  # sub-secciones + acento / tab color
    NAVY = "FF132E57"        # legacy CFI (fuera de whitelist)
    ORANGE = "FFED942D"      # legacy CFI (fuera de whitelist)
    TEAL = "FF1E8496"        # legacy CFI (fuera de whitelist)
    INPUT_BLUE = "FF0000FF"  # analyst input font
    LINK_GREEN = "FF00CC00"  # cross-sheet link font
    WARN_RED = "FFFF0000"    # error font
    WHITE = "FFFFFFFF"
    BLACK = "FF000000"
    INPUT_FILL = "FFFFF2CC"  # light yellow input shading
    SCENARIO_FILL = "FFF2F2F2"  # light gray scenario areas


class NumFmt(str, Enum):
    """Number-format whitelist (literal CFI strings)."""

    GENERAL = "General"
    NUM = '_-* #,##0_-;\\(#,##0\\)_-;_-* "-"_-;_-@_-'   # thousands, (neg), dash zero
    NUM_RED = "#,##0_);[Red](#,##0);-"                   # red negatives variant
    PCT1 = "0.0%"
    PCT2 = "0.00%"
    DEC2 = "0.00"
    MULT = "0.0\\x"                                       # 12.3x
    USD = '"$"#,##0_);\\("$"#,##0\\)'
    USD_CENTS = '"$"#,##0.00_);\\("$"#,##0.00\\)'
    YEAR_A = '0"A"'                                       # 2025A
    YEAR_E = '0"E"'                                       # 2026E
    DATE = "mm-dd-yy"
    HIDDEN = ";;;"


class CellRole(Enum):
    """Semantic cell roles -> font color mapping (traceability contract)."""

    LABEL = auto()      # black text
    INPUT = auto()      # blue font + yellow fill (analyst assumption)
    OBSERVED = auto()   # blue font, NO fill (historical, cited in comment)
    FORMULA = auto()    # black font
    LINK = auto()       # green font (pulls from another sheet)
    WARN = auto()       # red font


_ROLE_FONT_COLOR: dict[CellRole, Color] = {
    CellRole.LABEL: Color.BLACK,
    CellRole.INPUT: Color.INPUT_BLUE,
    CellRole.OBSERVED: Color.INPUT_BLUE,
    CellRole.FORMULA: Color.BLACK,
    CellRole.LINK: Color.LINK_GREEN,
    CellRole.WARN: Color.WARN_RED,
}

_THIN = Side(style="thin")
_DOUBLE = Side(style="double")

# Sheets that must carry frozen panes (data grids). Cover/Checks/Summary exempt.
FROZEN_SHEET_PREFIXES = ("Operating", "Annual", "Model", "Assumptions",
                         "Macro", "IS", "BS", "CF", "Ratios", "Schedules",
                         "Rev_Reconcile", "Val_", "Quarterly")

# Ratios completeness contract (check F13): these labels must exist in the
# Ratios section — build_ratios() writes exactly these, so builder output
# passes by construction and a lazy hand-built Ratios fails.
REQUIRED_RATIO_LABELS = (
    "Margen neto", "Rotacion de activos", "Apalancamiento", "ROE DuPont 3",
    "Carga fiscal", "Carga de interes", "Margen EBIT", "ROE DuPont 5",
    "NOPAT", "Capital invertido", "ROIC", "Economic profit",
    "Margen bruto", "Margen operativo", "Razon corriente", "Quick ratio",
    "Deuda / EBITDA aprox", "Cobertura de intereses",
    "DSO", "DIO", "DPO", "CCC",
    "DFL", "CFO / NI", "Accruals",
)

JUNK_SHEET_NAMES = ("Hoja1", "Hoja2", "Sheet1", "Sheet2", "Hoja 1", "Sheet 1")


@dataclass(frozen=True)
class PeriodHeader:
    """Year header spec: e.g. 2019..2031, actuals through 2025."""

    first_year: int
    last_year: int
    last_actual_year: int


# Rebrandable DECORATIVE slots (brand/DESIGN.md). Semantic colors — input blue,
# link green, warn red, input/scenario fills — are the traceability contract
# and are NEVER brandable.
BRAND_SLOTS = ("brand_primary", "brand_section", "brand_accent")


def load_brand(path: str) -> dict[str, str]:
    """Parse brand/DESIGN.md lines like ``brand_primary: #132E57`` -> ARGB.

    Deterministic: only the three BRAND_SLOTS are read; anything else in the
    file is prose for humans. Missing file or missing slot -> CFI default.
    """
    import re
    try:
        with open(path, encoding="utf-8") as fh:
            text = fh.read()
    except OSError:
        return {}
    out: dict[str, str] = {}
    for slot in BRAND_SLOTS:
        match = re.search(rf"{slot}\s*[:=]\s*#?([0-9A-Fa-f]{{6}})", text)
        if match:
            out[slot] = "FF" + match.group(1).upper()
    return out


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------


class ModelStyler:
    """Owns a Workbook and applies the format contract. One instance per model."""

    def __init__(self, units_label: str = "USD millones salvo indicado",
                 brand: Optional[dict[str, str]] = None) -> None:
        self.wb: Workbook = Workbook()
        self.units_label = units_label
        b = brand or {}
        self.color_primary: str = b.get("brand_primary", Color.DARK_BLUE.value)
        self.color_section: str = b.get("brand_section", Color.DARK_BLUE.value)
        self.color_accent: str = b.get("brand_accent", Color.LIGHT_BLUE.value)
        default = self.wb.active
        if default is not None:
            self.wb.remove(default)
        self.wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
        self._stamp()

    # -- workbook level -----------------------------------------------------

    def _stamp(self) -> None:
        """Custom doc property proving builder provenance (check F10)."""
        self.wb.properties.keywords = f"{BUILDER_STAMP_KEY}={BUILDER_STAMP_VALUE}"

    def set_periodicity(self, mode: str) -> None:
        """Stamp the model's periodicity (issuer-profile) into the workbook so
        check F14 can verify quarter columns without external context."""
        self.wb.properties.keywords = (
            f"{self.wb.properties.keywords or ''};periodicity={mode}")

    def define_constant(self, name: str, sheet: str, coord: str) -> None:
        """Named range for a labeled constant (e.g. DAYS_YEAR) — kills hardcodes."""
        from openpyxl.utils.cell import absolute_coordinate
        from openpyxl.workbook.defined_name import DefinedName
        # absolute_coordinate soporta columnas de 2+ letras ("AA10" -> "$AA$10");
        # el corte por caracter producia "$A$A10", una referencia invalida.
        quoted = sheet.replace("'", "''")
        ref = f"'{quoted}'!{absolute_coordinate(coord)}"
        self.wb.defined_names.add(DefinedName(name, attr_text=ref))

    def save(self, path: str) -> None:
        self.wb.save(path)

    # -- sheet level --------------------------------------------------------

    def new_sheet(self, name: str, freeze: Optional[str] = "C4",
                  tab_color: Optional[str] = None) -> Worksheet:
        """tab_color: ARGB string; use self.color_accent for branded tabs."""
        ws = self.wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        if freeze:
            ws.freeze_panes = freeze
        if tab_color is not None:
            ws.sheet_properties.tabColor = tab_color
        return ws

    def brand_bar(self, ws: Worksheet, title: str, last_col: int = 18) -> None:
        """Rows 1-2: brand bar + sheet title + units note. Row 3 reserved for checks."""
        for col in range(1, last_col + 1):
            for row in (1, 2):
                cell = ws.cell(row=row, column=col)
                cell.fill = PatternFill("solid", fgColor=self.color_primary)
        c = ws.cell(row=1, column=1, value="(c) research_analyst - todos los supuestos son del analista")
        c.font = Font(name=FONT_NAME, size=8, color=Color.WHITE.value)
        t = ws.cell(row=2, column=1, value=title)
        t.font = Font(name=FONT_NAME, size=16, bold=True, color=Color.WHITE.value)
        u = ws.cell(row=2, column=4, value=f"({self.units_label})")
        u.font = Font(name=FONT_NAME, size=11, color=Color.WHITE.value)

    def period_header(self, ws: Worksheet, row: int, first_col: int,
                      spec: PeriodHeader) -> None:
        """Year row with A/E suffix formats (2025A / 2026E)."""
        col = first_col
        for year in range(spec.first_year, spec.last_year + 1):
            cell = ws.cell(row=row, column=col, value=year)
            fmt = NumFmt.YEAR_A if year <= spec.last_actual_year else NumFmt.YEAR_E
            cell.number_format = fmt.value
            cell.font = Font(name=FONT_NAME, size=11, bold=True)
            cell.alignment = Alignment(horizontal="center")
            col += 1

    def check_row(self, ws: Worksheet, row: int, first_col: int, n_cols: int,
                  formula_template: str) -> None:
        """Per-column check row near the top (frozen visible). Template uses {col}."""
        label = ws.cell(row=row, column=2, value="Balance Sheet Check")
        label.font = Font(name=FONT_NAME, size=11, bold=True)
        self._nav_mark(ws, row)
        for i in range(n_cols):
            col_letter = get_column_letter(first_col + i)
            cell = ws.cell(row=row, column=first_col + i,
                           value=formula_template.format(col=col_letter))
            cell.font = Font(name=FONT_NAME, size=11, color=Color.WARN_RED.value)
            cell.alignment = Alignment(horizontal="center")

    # -- row/cell level -----------------------------------------------------

    def _nav_mark(self, ws: Worksheet, row: int,
                  color: str = Color.BLACK.value) -> None:
        """Navigation pattern: column A carries ONLY an 'x' on each header/
        sub-header row, so Ctrl+arrow on column A jumps section to section.
        Labels live in column B."""
        c = ws.cell(row=row, column=1, value="x")
        c.font = Font(name=FONT_NAME, size=8, color=color)

    def _require_breath(self, ws: Worksheet, row: int) -> None:
        """F16 por construccion: la fila previa a un header debe estar VACIA o
        ser otro header (consecutivos). Falla al construir, no en el audit."""
        if row <= 1:
            return
        prev_is_header = ws.cell(row=row - 1, column=1).value == "x"
        prev_has_content = any(ws.cell(row=row - 1, column=c).value is not None
                               for c in (2, 3))
        if prev_has_content and not prev_is_header:
            raise ValueError(
                f"F16: header en {ws.title}!fila {row} sin fila en blanco previa "
                f"(fila {row - 1} tiene contenido). Deja una fila de respiro.")

    def section_header(self, ws: Worksheet, row: int, title: str,
                       last_col: int = 18) -> None:
        """Banda azul oscuro, texto BLANCO bold 16 — una por seccion."""
        self._require_breath(ws, row)
        for col in range(1, last_col + 1):
            ws.cell(row=row, column=col).fill = PatternFill(
                "solid", fgColor=self.color_section)
        c = ws.cell(row=row, column=2, value=title)
        c.font = Font(name=FONT_NAME, size=16, bold=True,
                      color=Color.WHITE.value)
        self._nav_mark(ws, row, color=Color.WHITE.value)

    def subsection(self, ws: Worksheet, row: int, title: str,
                   last_col: int = 18) -> None:
        """Banda azul claro, texto negro bold 14 — sub-jerarquia."""
        self._require_breath(ws, row)
        for col in range(1, last_col + 1):
            ws.cell(row=row, column=col).fill = PatternFill(
                "solid", fgColor=self.color_accent)
        c = ws.cell(row=row, column=2, value=title)
        c.font = Font(name=FONT_NAME, size=14, bold=True)
        self._nav_mark(ws, row)

    def schedule_block_header(self, ws: Worksheet, row: int, name: str) -> None:
        """Header of one 'Sch: <name>' block inside the single Schedules sheet."""
        self.subsection(ws, row, f"Sch: {name}")

    def set_cell(self, ws: Worksheet, coord: str, value: object, role: CellRole,
                 numfmt: NumFmt = NumFmt.NUM, size: int = 11,
                 bold: bool = False) -> None:
        cell = ws[coord]
        cell.value = value
        cell.font = Font(name=FONT_NAME, size=size, bold=bold,
                         color=_ROLE_FONT_COLOR[role].value)
        cell.number_format = numfmt.value
        if role is CellRole.INPUT:
            cell.fill = PatternFill("solid", fgColor=Color.INPUT_FILL.value)

    def check_result(self, ws: Worksheet, row: int, col: int, result: object, *,
                     note: Optional[str] = None,
                     computed_at: Optional[str] = None) -> None:
        """Celda de resultado de un check (tab Checks o bloque de vigencia).

        Dos formas validas: FORMULA viva (empieza con '=') o resultado de un
        ESCANEO por codigo al construir — entonces ``computed_at`` (YYYY-MM-DD)
        y ``note`` son obligatorios: la fecha va en la celda contigua y la nota
        como comentario. Un literal sin fecha es un check congelado (queda
        verde aunque el hecho cambie) y se rechaza.
        """
        is_formula = isinstance(result, str) and result.startswith("=")
        if not is_formula:
            if not computed_at or not note:
                raise ValueError(
                    f"check_result {ws.title}!fila {row}: resultado literal "
                    f"{result!r} sin computed_at/note — un check no se afirma "
                    "sin fecha de calculo ni evidencia (usa formula o escaneo).")
            date.fromisoformat(computed_at)          # valida el formato
        self.set_cell(ws, f"{get_column_letter(col)}{row}", result,
                      CellRole.WARN, NumFmt.GENERAL)
        if not is_formula:
            self.set_cell(ws, f"{get_column_letter(col + 1)}{row}",
                          f"calc. {computed_at}", CellRole.LABEL, NumFmt.GENERAL,
                          size=9)
            ws.cell(row=row, column=col).comment = Comment(note, "research_analyst")

    def series_row(self, ws: Worksheet, row: int, label: str, first_col: int,
                   hist_values: list[object], forecast_values: list[object],
                   numfmt: NumFmt = NumFmt.NUM,
                   hist_role: CellRole = CellRole.FORMULA,
                   forecast_role: CellRole = CellRole.INPUT) -> None:
        """ONE continuous row across the whole horizon (contract F11).

        A series is never split into 'historical' and 'forecast' rows: the same
        line carries computed/observed history (black or blue-observed) and then
        forecast cells (blue input on yellow, or driver formula). The role
        switches at the boundary column; every horizon column gets a value.
        """
        lab = ws.cell(row=row, column=2, value=label)
        lab.font = Font(name=FONT_NAME, size=11, color=Color.BLACK.value)
        col = first_col
        for value in hist_values:
            self.set_cell(ws, f"{get_column_letter(col)}{row}", value,
                          hist_role, numfmt)
            col += 1
        for value in forecast_values:
            self.set_cell(ws, f"{get_column_letter(col)}{row}", value,
                          forecast_role, numfmt)
            col += 1

    def subtotal_border(self, ws: Worksheet, row: int, first_col: int,
                        n_cols: int) -> None:
        for i in range(n_cols):
            ws.cell(row=row, column=first_col + i).border = Border(top=_THIN)

    def total_border(self, ws: Worksheet, row: int, first_col: int,
                     n_cols: int) -> None:
        for i in range(n_cols):
            ws.cell(row=row, column=first_col + i).border = Border(
                top=_THIN, bottom=_DOUBLE)

    def group_rows(self, ws: Worksheet, start: int, end: int,
                   hidden: bool = False) -> None:
        """Outline level 1 so sections collapse to summary view.

        La fila EN BLANCO final del rango se deja FUERA del grupo: si la
        separacion vive dentro del bloque, al colapsar desaparece y los
        headers quedan pegados — justo en la vista donde mas se necesita.
        El respiro pertenece al esqueleto de la hoja, no al bloque.
        """
        last = end
        while last >= start:
            has_content = any(ws.cell(row=last, column=c).value is not None
                              for c in range(1, min(ws.max_column, _MAX_SCAN_COLS) + 1))
            if has_content:
                break
            last -= 1
        if last < start:
            # Rango entero sin contenido (scaffold aun sin poblar): agrupar
            # todo tal cual — no hay separacion que preservar todavia.
            last = end
        for r in range(start, last + 1):
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].hidden = hidden
        for r in range(last + 1, end + 1):
            ws.row_dimensions[r].outlineLevel = 0
            ws.row_dimensions[r].hidden = False

    def label_col_width(self, ws: Worksheet, width: float = 42.0) -> None:
        """Column A = narrow navigation column ('x' markers); B = labels."""
        ws.column_dimensions["A"].width = 2.5
        ws.column_dimensions["B"].width = width

    def quarter_header(self, ws: Worksheet, row: int, first_col: int,
                       quarters: list[str]) -> int:
        """Quarter labels ('1Q2026E' / '3Q2025A'). Returns the next free column."""
        col = first_col
        for label in quarters:
            cell = ws.cell(row=row, column=col, value=label)
            cell.font = Font(name=FONT_NAME, size=11, bold=True)
            cell.alignment = Alignment(horizontal="center")
            col += 1
        return col

    def interleaved_header(self, ws: Worksheet, row: int, first_col: int,
                           first_year: int, last_year: int,
                           quarterly_from_year: int, last_actual_year: int,
                           last_actual_quarter: int) -> dict[tuple[int, str], int]:
        """Quarterly-native header: per fiscal year, 4 quarter columns then the
        FY aggregate column; years before ``quarterly_from_year`` get FY only.

        DEPRECATED (diseño 2026-08-31): el modo quarterly usa dos hojas
        (Operating trimestral puro / Annual agregada). Se conserva para
        modelos legacy.

        Returns {(year, '1Q'|'2Q'|'3Q'|'4Q'|'FY'): column} so the build wires
        FY = aggregate-of-quarters formulas (C8 estructural) and populate knows
        where each period lives.
        """
        colmap: dict[tuple[int, str], int] = {}
        col = first_col
        for year in range(first_year, last_year + 1):
            if year >= quarterly_from_year:
                for q in (1, 2, 3, 4):
                    is_actual = (year < last_actual_year or
                                 (year == last_actual_year and q <= last_actual_quarter))
                    label = f"{q}Q{year}{'A' if is_actual else 'E'}"
                    cell = ws.cell(row=row, column=col, value=label)
                    cell.font = Font(name=FONT_NAME, size=11, bold=True)
                    cell.alignment = Alignment(horizontal="center")
                    colmap[(year, f"{q}Q")] = col
                    col += 1
            fy = ws.cell(row=row, column=col, value=year)
            fy.number_format = (NumFmt.YEAR_A if year <= last_actual_year
                                else NumFmt.YEAR_E).value
            fy.font = Font(name=FONT_NAME, size=11, bold=True)
            fy.alignment = Alignment(horizontal="center")
            colmap[(year, "FY")] = col
            col += 1
        return colmap

    # -- Ratios section (deterministic; fixes the "lazy ratios" failure) -----

    def build_ratios(self, ws: Worksheet, start_row: int, first_col: int,
                     n_cols: int, ref: dict[str, str],
                     wacc_ref: Optional[str] = None,
                     days_ref: str = "DAYS_YEAR") -> tuple[int, list[str]]:
        """Write the FULL Ratios section (blocks A-G of ratios-analytics.md).

        ``days_ref``: named range de dias para las razones de dias (DSO/DIO/
        DPO/CCC). REGLA FINANCIERA: el numerador es un STOCK promedio y el
        denominador un FLUJO — ambos deben cubrir la MISMA ventana. En hojas
        trimestrales, o el flujo es UDM (12 meses) con ``DAYS_YEAR``, o es el
        flujo del trimestre con ``DAYS_QUARTER``; mezclar flujo trimestral con
        DAYS_YEAR infla los dias ~4x.

        ``ref`` maps canon line -> absolute row reference WITHOUT column, e.g.
        {"rev": "Model!{c}27", ...} where "{c}" is replaced per period column
        and "{p}" by the previous column. Required canons: rev, cogs, gross,
        ebit, ebt, ni, interest, tax, ta, equity, cash, ar, inv, ap, ca, cl,
        debt, re, cfo, da. Missing canons skip their rows (returned in the
        skipped list) — but check F13 fails if the section is incomplete, so
        a skip is visible, never silent.
        """
        skipped: list[str] = []
        r = start_row
        # Auto-generar referencias de periodo previo: el caller pasa canons
        # planos ({c}); aqui se derivan los <canon>_p ({p}) que usan las
        # plantillas — evita el footgun de claves manuales.
        ref = dict(ref)
        for k, v in list(ref.items()):
            if not k.endswith("_p") and isinstance(v, str) and "{c}" in v:
                ref.setdefault(k + "_p", v.replace("{c}", "{p}"))

        def row_out(label: str, template: str, fmt: NumFmt,
                    needs: tuple[str, ...]) -> None:
            nonlocal r
            if any(k not in ref for k in needs):
                skipped.append(label)
                return
            ws.cell(row=r, column=2, value=label).font = Font(
                name=FONT_NAME, size=11)
            for i in range(1, n_cols):  # first period column has no prior year
                col = get_column_letter(first_col + i)
                prev = get_column_letter(first_col + i - 1)
                parts = {k: v.replace("{c}", col).replace("{p}", prev)
                         for k, v in ref.items()}
                formula = template.format(**parts)
                self.set_cell(ws, f"{col}{r}", formula, CellRole.FORMULA, fmt)
            r += 1

        def header(title: str) -> None:
            nonlocal r
            self.section_header(ws, r, title)
            r += 1

        AVG = "AVERAGE({p_ref},{c_ref})"

        def avg(canon: str) -> dict[str, str]:
            return {}

        header("DuPont")
        row_out("Margen neto (NI/Ventas)", '=IF({rev}=0,"",{ni}/{rev})', NumFmt.PCT1, ("ni", "rev"))
        row_out("Rotacion de activos (Ventas/Activos prom.)",
                '=IF(AVERAGE({ta_p},{ta})=0,"",{rev}/AVERAGE({ta_p},{ta}))', NumFmt.DEC2, ("rev", "ta", "ta_p"))
        row_out("Apalancamiento (Activos/Capital prom.)",
                '=IF(AVERAGE({equity_p},{equity})=0,"",AVERAGE({ta_p},{ta})/AVERAGE({equity_p},{equity}))', NumFmt.DEC2, ("ta", "ta_p", "equity", "equity_p"))
        row_out("ROE DuPont 3", '=IF(AVERAGE({equity_p},{equity})=0,"",{ni}/AVERAGE({equity_p},{equity}))', NumFmt.PCT1, ("ni", "equity", "equity_p"))
        row_out("Carga fiscal (NI/EBT)", '=IF({ebt}=0,"",{ni}/{ebt})', NumFmt.PCT1, ("ni", "ebt"))
        row_out("Carga de interes (EBT/EBIT)", '=IF({ebit}=0,"",{ebt}/{ebit})', NumFmt.PCT1, ("ebt", "ebit"))
        row_out("Margen EBIT (EBIT/Ventas)", '=IF({rev}=0,"",{ebit}/{rev})', NumFmt.PCT1, ("ebit", "rev"))
        row_out("ROE DuPont 5 (producto)",
                '=IF(OR({ebt}=0,{ebit}=0,{rev}=0,AVERAGE({equity_p},{equity})=0),"",'
                '{ni}/{ebt}*{ebt}/{ebit}*{ebit}/{rev}*{rev}/AVERAGE({ta_p},{ta})'
                '*AVERAGE({ta_p},{ta})/AVERAGE({equity_p},{equity}))', NumFmt.PCT1,
                ("ni", "ebt", "ebit", "rev", "ta", "ta_p", "equity", "equity_p"))
        r += 1

        header("ROIC y economic profit")
        row_out("Tasa efectiva (tax/EBT)", '=IF({ebt}=0,"",{tax}/{ebt})', NumFmt.PCT1, ("tax", "ebt"))
        row_out("NOPAT (EBIT x (1-t))", '=IF({ebt}=0,"",{ebit}*(1-{tax}/{ebt}))', NumFmt.NUM, ("ebit", "tax", "ebt"))
        # Capital invertido PROMEDIO (consistente con el resto de ratios que
        # usan denominadores de balance promediados — antes usaba el saldo de
        # cierre, inconsistente con ratios-analytics.md).
        _ic = ("(AVERAGE({debt_p},{debt})+AVERAGE({equity_p},{equity})"
               "-AVERAGE({cash_p},{cash}))")
        row_out("Capital invertido promedio (deuda+capital-caja)", "=" + _ic,
                NumFmt.NUM, ("debt", "debt_p", "equity", "equity_p", "cash", "cash_p"))
        # Guardas explicitas: capital invertido <= 0 (caja > deuda+capital,
        # caso real en emisoras con caja neta enorme) hace el ROIC absurdo —
        # se reporta "n/s" (no significativo), jamas un numero inflado.
        row_out("ROIC",
                '=IF(OR(' + _ic + '<=0,{ebt}=0),"n/s",{ebit}*(1-{tax}/{ebt})/' + _ic + ')',
                NumFmt.PCT1,
                ("debt", "debt_p", "equity", "equity_p", "cash", "cash_p",
                 "ebit", "tax", "ebt"))
        if wacc_ref:
            row_out(f"Economic profit (spread vs WACC {wacc_ref})",
                    '=IF(OR(' + _ic + '<=0,{ebt}=0),"n/s",'
                    '({ebit}*(1-{tax}/{ebt})/' + _ic + '-' + wacc_ref + ')*' + _ic + ')',
                    NumFmt.NUM,
                    ("ebit", "tax", "ebt", "debt", "debt_p", "equity",
                     "equity_p", "cash", "cash_p"))
        else:
            skipped.append("Economic profit (sin wacc_ref)")
        r += 1

        header("Rentabilidad, liquidez y solvencia")
        row_out("Margen bruto", '=IF({rev}=0,"",{gross}/{rev})', NumFmt.PCT1, ("gross", "rev"))
        row_out("Margen operativo", '=IF({rev}=0,"",{ebit}/{rev})', NumFmt.PCT1, ("ebit", "rev"))
        row_out("Razon corriente", '=IF({cl}=0,"",{ca}/{cl})', NumFmt.DEC2, ("ca", "cl"))
        row_out("Quick ratio", '=IF({cl}=0,"",({ca}-{inv})/{cl})', NumFmt.DEC2, ("ca", "cl", "inv"))
        row_out("Deuda / EBITDA aprox (EBIT+D&A)", '=IF(({ebit}+{da})=0,"",{debt}/({ebit}+{da}))', NumFmt.DEC2, ("debt", "ebit", "da"))
        row_out("Cobertura de intereses (EBIT/interes)", '=IF({interest}=0,"n/a",{ebit}/ABS({interest}))', NumFmt.DEC2, ("ebit", "interest"))
        r += 1

        header("Ciclo de conversion de efectivo")
        row_out("DSO (dias)", '=IF({rev}=0,"",AVERAGE({ar_p},{ar})/{rev}*' + days_ref + ')', NumFmt.DEC2, ("ar", "ar_p", "rev"))
        row_out("DIO (dias)", '=IF({cogs}=0,"",AVERAGE({inv_p},{inv})/{cogs}*' + days_ref + ')', NumFmt.DEC2, ("inv", "inv_p", "cogs"))
        row_out("DPO (dias)", '=IF({cogs}=0,"",AVERAGE({ap_p},{ap})/{cogs}*' + days_ref + ')', NumFmt.DEC2, ("ap", "ap_p", "cogs"))
        row_out("CCC (DSO+DIO-DPO)",
                '=IF({cogs}=0,"",AVERAGE({ar_p},{ar})/{rev}*' + days_ref
                + '+AVERAGE({inv_p},{inv})/{cogs}*' + days_ref
                + '-AVERAGE({ap_p},{ap})/{cogs}*' + days_ref + ')',
                NumFmt.DEC2, ("ar", "ar_p", "inv", "inv_p", "ap", "ap_p", "rev", "cogs"))
        r += 1

        header("Apalancamiento operativo y calidad")
        row_out("DFL (EBIT/(EBIT-interes))", '=IF(({ebit}-ABS({interest}))=0,"",{ebit}/({ebit}-ABS({interest})))', NumFmt.DEC2, ("ebit", "interest"))
        row_out("CFO / NI (calidad de utilidades)", '=IF({ni}=0,"",{cfo}/{ni})', NumFmt.DEC2, ("cfo", "ni"))
        row_out("Accruals proxy (NI-CFO)/Activos prom.",
                '=IF(AVERAGE({ta_p},{ta})=0,"",({ni}-{cfo})/AVERAGE({ta_p},{ta}))', NumFmt.PCT1, ("ni", "cfo", "ta", "ta_p"))
        return r, skipped


# ---------------------------------------------------------------------------
# Checks F — format audit (deterministic, any workbook)
# ---------------------------------------------------------------------------

_ALLOWED_FONT_COLORS = {c.value for c in (
    Color.INPUT_BLUE, Color.LINK_GREEN, Color.WARN_RED, Color.WHITE, Color.BLACK)}
_ALLOWED_FONT_COLORS.add("FF333333")  # near-black tolerated
_ALLOWED_FILLS = {c.value for c in (
    Color.DARK_BLUE, Color.LIGHT_BLUE, Color.INPUT_FILL, Color.SCENARIO_FILL)}
_ALLOWED_NUMFMTS = {f.value for f in NumFmt}

# Limites de escaneo del audit. Una hoja mas grande NO se audita en silencio
# a medias: el check F20 falla y nombra la hoja.
_MAX_SCAN_ROWS = 5000
_MAX_SCAN_COLS = 400


@dataclass(frozen=True)
class Finding:
    """Resultado de un check F. ``pending`` = no se pudo evaluar (p. ej. libro
    sin valores calculados): NO es verde ni falla; se reporta [pendiente]."""

    check: str
    ok: bool
    detail: str
    pending: bool = False


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


def _is_header_row(ws: Worksheet, r: int) -> bool:
    for col in (1, 2):
        f = ws.cell(row=r, column=col).font
        if f is not None and f.bold and (f.size or 0) >= 13:
            return True
    return False


# --- Deteccion de referencias circulares (check S11, sin Excel) ------------

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
    """S11: ciclos reales en el grafo de fórmulas (DFS con pila).

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


def audit_format(path: str, brand: Optional[dict[str, str]] = None) -> list[Finding]:
    """Run checks F (F1-F20) on a workbook. Pure read; returns findings.

    ``brand``: output of load_brand(brand/DESIGN.md) — its values extend the
    fill whitelist (F4) so a branded model audits green with its own DESIGN.md.
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
                            kw if stamped else "sin sello (no construido por xlsx_builder)"))
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
    return findings


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
        r_open = r_close = r_change = None
        for r in range(1, min(wsf.max_row, _MAX_SCAN_ROWS) + 1):
            lab = str(wsf.cell(row=r, column=2).value or "").lower()
            if not lab:
                continue
            if r_open is None and ("al inicio" in lab or "inicial" in lab
                                   or "apertura" in lab):
                r_open = r
            elif r_close is None and ("al cierre" in lab or "final" in lab):
                r_close = r
            elif r_change is None and ("cambio neto" in lab
                                       or "aumento (disminucion)" in lab
                                       or "variacion neta" in lab):
                r_change = r
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


def _print_report(findings: Iterable[Finding]) -> int:
    """Imprime el reporte. Exit: 1 si hay fallas; 3 si solo hay pendientes."""
    findings = list(findings)
    failures = pending = 0
    for f in findings:
        if f.pending:
            mark = "[pendiente]"
            pending += 1
        elif f.ok:
            mark = "[ok]"
        else:
            mark = "[x]"
            failures += 1
        print(f"{mark} {f.check}: {f.detail}")
    ok_count = sum(1 for f in findings if f.ok and not f.pending)
    print(f"Resumen F: {ok_count} ok, {failures} fallas, {pending} pendientes")
    if failures:
        return 1
    if pending:
        print("Veredicto: PENDIENTE - no es verde; resuelve los pendientes y re-audita")
        return 3
    return 0


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
    end_row, skipped = styler.build_ratios(op, 37, first, n_cols, ref,
                                           wacc_ref="WACC", days_ref="DAYS_QUARTER")
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
    print(f"[ok] demo escrito: {path}")


def main(argv: list[str]) -> int:
    if argv[1:2] == ["audit"] and len(argv) in (3, 4):
        brand = load_brand(argv[3]) if len(argv) == 4 else None
        return _print_report(audit_format(argv[2], brand=brand))
    if len(argv) == 3 and argv[1] == "demo":
        _demo(argv[2])
        return _print_report(audit_format(argv[2]))
    print("uso: python xlsx_builder.py audit <path.xlsx> [brand/DESIGN.md] | demo <path.xlsx>")
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
