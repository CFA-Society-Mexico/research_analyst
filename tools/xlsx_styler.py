"""Builder del xlsx: ``ModelStyler``, la unica via para escribir un modelo.

Hojas, barra de marca, headers de periodo, secciones y bloques, celdas por
rol (con lint F21 al escribir), series hist/forecast, bordes, agrupacion de
filas, bloques de valuacion (Gordon, crecimiento implicito, sensibilidad), la
tab Ratios (``build_ratios``), sellos del libro y vinculo con la capa de
captura (C9). Formato por construccion: lo que escribe pasa los checks F de
``xlsx_audit``. Se importa a traves de la fachada ``xlsx_builder``.
"""

# SIN ``from __future__ import annotations`` a proposito: ver el encabezado
# de ``xlsx_builder.py`` (el builder tambien se carga por ruta).

from datetime import date
from typing import Callable, Optional

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from xlsx_contract import (
    _MAX_SCAN_COLS,
    _ROLE_FONT_COLOR,
    _UNSET,
    BUILDER_STAMP_KEY,
    BUILDER_STAMP_VALUE,
    CAPTURE_NAME_PREFIX,
    FONT_NAME,
    RATIO_FLOW_CANONS,
    TRANSITION_STAMP_KEY,
    CellRole,
    Color,
    NumFmt,
    PeriodHeader,
    transition_year_of,
)
from xlsx_formulas import NO_GROWTH_BELOW_WACC, NO_REAL_SOLUTION, NOT_MEANINGFUL, lint_formula
from xlsx_scan import _has_quarter_header


_THIN = Side(style="thin")
_DOUBLE = Side(style="double")

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

    def set_transition_year(self, year: int) -> None:
        """Declara el ANIO DE TRANSICION: el primer anio del forecast que ya
        trae trimestres observados (9M reales + 4T guiado). Su columna lleva
        una formula propia (ancla en lo observado) y S5 la trata como tramo
        aparte. Sin este sello, ``series_row(transition_first=...)`` falla:
        una formula distinta en la primera columna E seria una violacion de
        S5 que nadie declaro."""
        if not 1990 <= int(year) <= 2100:
            raise ValueError(f"set_transition_year: anio fuera de rango: {year!r}")
        self.wb.properties.keywords = (
            f"{self.wb.properties.keywords or ''};{TRANSITION_STAMP_KEY}={int(year)}")

    def transition_year(self) -> Optional[int]:
        return transition_year_of(self.wb.properties.keywords)

    def define_constant(self, name: str, sheet: str, coord: str) -> None:
        """Named range for a labeled constant (e.g. DAYS_YEAR) — kills hardcodes."""
        from openpyxl.utils.cell import absolute_coordinate
        from openpyxl.workbook.defined_name import DefinedName
        # absolute_coordinate soporta columnas de 2+ letras ("AA10" -> "$AA$10");
        # el corte por caracter producia "$A$A10", una referencia invalida.
        quoted = sheet.replace("'", "''")
        ref = f"'{quoted}'!{absolute_coordinate(coord)}"
        self.wb.defined_names.add(DefinedName(name, attr_text=ref))

    def link_capture(self, ws: Worksheet, row: int, line_item: str,
                     scale: float = 1.0) -> None:
        """Declara que la fila ``row`` trae el historico de ``line_item`` de
        la capa de captura (model/inputs/canonical_*.csv) — contrato del C9.

        Escribe el nombre definido ``CAPTURE.<line_item>`` con alcance de la
        hoja, apuntando a la fila completa, y la escala en su comentario.
        ``scale`` convierte la unidad del modelo a la de la fuente:
        valor_csv = valor_xlsx * scale (1_000_000 si el modelo vive en
        millones y el CSV en unidades). El historico se escribe SIN redondear;
        el formato de celda redondea la vista.
        """
        import re
        from openpyxl.workbook.defined_name import DefinedName
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", line_item):
            raise ValueError(
                f"link_capture: line_item invalido {line_item!r} "
                "(letras, digitos y guion bajo; empieza con letra)")
        if isinstance(scale, bool) or not isinstance(scale, (int, float)) or not scale > 0:
            raise ValueError(f"link_capture: scale debe ser un numero > 0, no {scale!r}")
        name = CAPTURE_NAME_PREFIX + line_item
        if name in ws.defined_names:
            raise ValueError(
                f"link_capture: {line_item} ya esta vinculado en la hoja {ws.title} "
                "(una fila, una linea)")
        quoted = ws.title.replace("'", "''")
        ws.defined_names.add(DefinedName(
            name, attr_text=f"'{quoted}'!${row}:${row}",
            comment=f"scale={float(scale)!r}"))

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
        """Escribe una celda con su rol. F21 por construccion: una formula que
        Excel descartaria o evaluaria mal (``lint_formula``) se rechaza AQUI,
        con la celda, en vez de aparecer como libro reparado al abrirlo."""
        problems = lint_formula(value)
        if problems:
            raise ValueError(f"F21: {ws.title}!{coord} {value!r}: " + "; ".join(problems))
        self._write(ws, coord, value, role, numfmt, size, bold)

    def set_array_formula(self, ws: Worksheet, coord: str, formula: str,
                          role: CellRole = CellRole.FORMULA,
                          numfmt: NumFmt = NumFmt.NUM) -> None:
        """Formula de ARREGLO de una celda (CSE). Es la unica forma valida de
        escribir una operacion sobre rangos fuera de SUMPRODUCT, p. ej.
        ``=MEDIAN(IF(rango=1,valores))``: como formula normal Excel 365 aplica
        interseccion implicita y devuelve otra cosa. La aridad se valida igual."""
        problems = lint_formula(formula, array=True)
        if problems or not formula.startswith("="):
            raise ValueError(f"F21: {ws.title}!{coord} {formula!r}: "
                             + ("; ".join(problems) or "no es formula"))
        self._write(ws, coord, ArrayFormula(ref=coord, text=formula), role, numfmt, 11, False)

    def _write(self, ws: Worksheet, coord: str, value: object, role: CellRole,
               numfmt: NumFmt, size: int, bold: bool) -> None:
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
                   forecast_role: CellRole = CellRole.INPUT,
                   transition_first: object = _UNSET,
                   transition_role: CellRole = CellRole.FORMULA) -> None:
        """ONE continuous row across the whole horizon (contract F11).

        A series is never split into 'historical' and 'forecast' rows: the same
        line carries computed/observed history (black or blue-observed) and then
        forecast cells (blue input on yellow, or driver formula). The role
        switches at the boundary column; every horizon column gets a value.

        ``transition_first``: valor/formula de la PRIMERA columna del forecast
        cuando ese anio es de transicion (trimestres ya observados + resto
        guiado). Sustituye a ``forecast_values[0]`` -la lista sigue cubriendo
        todo el forecast- y se escribe con ``transition_role`` (default
        FORMULA: en un anio anclado el driver es RESULTADO del ancla, no
        input). Exige ``set_transition_year`` antes: asi S5 sabe que esa
        columna es un tramo propio y no una formula inconsistente.
        """
        if transition_first is not _UNSET:
            if self.transition_year() is None:
                raise ValueError(
                    f"series_row {ws.title}!fila {row}: transition_first sin anio de "
                    "transicion declarado. Llama antes a set_transition_year(<anio>).")
            if not forecast_values:
                raise ValueError(
                    f"series_row {ws.title}!fila {row}: transition_first sin forecast")
            forecast_values = list(forecast_values)
            first_forecast, forecast_values = forecast_values[0], forecast_values[1:]
            del first_forecast  # sustituido por transition_first
        lab = ws.cell(row=row, column=2, value=label)
        lab.font = Font(name=FONT_NAME, size=11, color=Color.BLACK.value)
        col = first_col
        for value in hist_values:
            self.set_cell(ws, f"{get_column_letter(col)}{row}", value,
                          hist_role, numfmt)
            col += 1
        if transition_first is not _UNSET:
            self.set_cell(ws, f"{get_column_letter(col)}{row}", transition_first,
                          transition_role, numfmt)
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

    # -- Valuacion: terminal normalizado, g implicita, sensibilidad ----------

    def _value_row(self, ws: Worksheet, row: int, col: int, label: str,
                   formula: str, numfmt: NumFmt) -> str:
        """Fila de valor unico: label en B, valor en ``col`` (C por contrato)."""
        ws.cell(row=row, column=2, value=label).font = Font(name=FONT_NAME, size=11)
        self.set_cell(ws, f"{get_column_letter(col)}{row}", formula,
                      CellRole.FORMULA, numfmt)
        return f"${get_column_letter(col)}${row}"

    def _normalized_terminal(self, ws: Worksheet, row: int, col: int, *, base_ref: str,
                             g_ref: str, rate_ref: str, return_ref: str,
                             labels: tuple[str, str, str]) -> dict[str, str]:
        """Terminal de Gordon con flujo normalizado, generico (tres filas).

        ``flujo = base x (1+g) x (1 - g/retorno)`` y ``TV = flujo / (tasa - g)``.
        Lado firma: base = NOPAT_n, tasa = WACC, retorno = RONIC. Lado equity:
        base = NI_n, tasa = ke, retorno = ROE terminal. Una sola matematica;
        los labels dicen cual es cual. Devuelve ``reinvestment``, ``flow``,
        ``tv`` y ``next_row``.
        """
        reinvestment = self._value_row(
            ws, row, col, labels[0],
            f'=IF({return_ref}<=0,"{NOT_MEANINGFUL}",{g_ref}/{return_ref})', NumFmt.PCT1)
        flow = self._value_row(
            ws, row + 1, col, labels[1],
            f'=IF({return_ref}<=0,"{NOT_MEANINGFUL}",'
            f"{base_ref}*(1+{g_ref})*(1-{g_ref}/{return_ref}))", NumFmt.NUM)
        tv = self._value_row(
            ws, row + 2, col, labels[2],
            f'=IF(OR({rate_ref}<={g_ref},NOT(ISNUMBER({flow}))),"{NOT_MEANINGFUL}",'
            f"{flow}/({rate_ref}-{g_ref}))", NumFmt.NUM)
        return {"reinvestment": reinvestment, "flow": flow, "tv": tv,
                "next_row": str(row + 3)}

    def gordon_terminal_block(self, ws: Worksheet, row: int, *, nopat_ref: str,
                              g_ref: str, wacc_ref: str, ronic_ref: str,
                              col: int = 3) -> dict[str, str]:
        """Terminal Gordon con FLUJO NORMALIZADO (tres filas, valores en C).

        ``flujo terminal = NOPAT_n x (1+g) x (1 - g/RONIC)``: la reinversion
        del terminal es la que el crecimiento g exige a un retorno RONIC, no
        la del ultimo anio explicito. Perpetuar ``FCFF_n`` arrastra el capex
        de una empresa que aun crece rapido sin darle ese crecimiento
        (TSMC: 1,025 por accion contra 2,065 con el flujo normalizado). RONIC
        es input etiquetado de Assumptions; con RONIC = WACC el crecimiento no
        crea valor (default defendible).

        Devuelve las referencias absolutas ``reinvestment``, ``flow``, ``tv``
        y ``next_row``.
        """
        return self._normalized_terminal(
            ws, row, col, base_ref=nopat_ref, g_ref=g_ref, rate_ref=wacc_ref,
            return_ref=ronic_ref,
            labels=("Tasa de reinversion terminal (g / RONIC)",
                    "Flujo terminal normalizado (NOPAT_n x (1+g) x (1 - g/RONIC))",
                    "TV Gordon (flujo terminal / (WACC - g))"))

    def implied_growth_block(self, ws: Worksheet, row: int, *, title: str,
                             tv_ref: str, nopat_ref: str, wacc_ref: str,
                             ronic_ref: str, col: int = 3) -> dict[str, str]:
        """g implicita en un valor terminal, forma cerrada, linea por linea.

        Despeje de ``TV = NOPAT_n (1+g)(1 - g/RONIC) / (WACC - g)``: la raiz
        menor de ``A g^2 - B g + C = 0`` con ``A = N/R``, ``B = TV + N(1-1/R)``
        y ``C = TV x W - N``. Se escribe sin literales (S4): ``m = B/(A+A)`` es
        el punto medio de las raices y ``g = m - RAIZ(m x m - C/A)``.

        Guarda obligatoria: con RONIC = WACC una raiz es exactamente g = WACC,
        solucion espuria (el TV explota). Una raiz >= WACC se reporta como
        texto (``NO_GROWTH_BELOW_WACC``), jamas como numero; igual si el
        discriminante es negativo (``NO_REAL_SOLUTION``). Con RONIC = WACC la
        cuadratica NO se usa: en punto flotante la raiz espuria sale
        0.08999999999999947 y pasaria por valida (visto con Excel). Ahi el
        despeje es lineal y exacto, ``g = TV x WACC / NOPAT_n - 1``.

        Sirve para el cruce de terminales (TV del multiplo de salida) y para
        el reverse DCF (TV implicita en el EV de mercado). Seis filas;
        devuelve la referencia de ``g`` y ``next_row``.
        """
        a = self._value_row(ws, row, col, f"{title}: A = NOPAT_n / RONIC",
                            f"={nopat_ref}/{ronic_ref}", NumFmt.NUM)
        b = self._value_row(ws, row + 1, col, f"{title}: B = TV + NOPAT_n x (1 - 1/RONIC)",
                            f"={tv_ref}+{nopat_ref}*(1-1/{ronic_ref})", NumFmt.NUM)
        c = self._value_row(ws, row + 2, col, f"{title}: C = TV x WACC - NOPAT_n",
                            f"={tv_ref}*{wacc_ref}-{nopat_ref}", NumFmt.NUM)
        mid = self._value_row(ws, row + 3, col, f"{title}: punto medio de las raices (B / 2A)",
                              f"={b}/({a}+{a})", NumFmt.PCT2)
        disc = self._value_row(ws, row + 4, col, f"{title}: discriminante reducido (m^2 - C/A)",
                               f"={mid}*{mid}-{c}/{a}", NumFmt.DEC2)
        root = f"{mid}-SQRT({disc})"
        linear = f"{tv_ref}*{wacc_ref}/{nopat_ref}-1"
        g = self._value_row(
            ws, row + 5, col, f"{title}: g implicita (raiz menor)",
            f'=IF(OR(NOT(ISNUMBER({disc})),{nopat_ref}<=0,{ronic_ref}<=0),"{NOT_MEANINGFUL}",'
            f'IF({ronic_ref}={wacc_ref},'
            f'IF({linear}>={wacc_ref},"{NO_GROWTH_BELOW_WACC}",{linear}),'
            f'IF({disc}<0,"{NO_REAL_SOLUTION}",'
            f'IF({root}>={wacc_ref},"{NO_GROWTH_BELOW_WACC}",{root}))))', NumFmt.PCT2)
        return {"g": g, "next_row": str(row + 6)}

    def sensitivity_grid(self, ws: Worksheet, top_row: int, left_col: int, *,
                         row_base_ref: str, col_base_ref: str,
                         row_step_ref: str, col_step_ref: str,
                         cell_formula: Callable[[str, str], str],
                         n_side: int = 2, corner_label: str = "WACC / g",
                         axis_fmt: NumFmt = NumFmt.PCT2,
                         value_fmt: NumFmt = NumFmt.DEC2) -> str:
        """Cuadricula de sensibilidad de 2 variables SIN data table.

        openpyxl no puede escribir un data table de Excel (``TABLE()``), asi
        que cada celda lleva la formula cerrada que devuelve
        ``cell_formula(ref_fila, ref_columna)`` -p. ej.
        ``dcf_per_share_formula``- con el encabezado de su fila y de su
        columna como variables. Encabezados: el central enlaza al valor base
        del modelo; los demas suman o restan el paso (constante con nombre en
        Assumptions), sin literales. Tamano ``(2 x n_side + 1)`` por lado.

        Devuelve la referencia absoluta de la celda CENTRAL, que debe
        reproducir el valor del modelo (``center_check_formula``).
        """
        if n_side < 1:
            raise ValueError(f"sensitivity_grid: n_side={n_side} invalido")
        size = 2 * n_side + 1
        corner = ws.cell(row=top_row, column=left_col, value=corner_label)
        corner.font = Font(name=FONT_NAME, size=11, bold=True)

        def axis(horizontal: bool, base_ref: str, step_ref: str) -> None:
            coords = [(f"{get_column_letter(left_col + 1 + k)}{top_row}" if horizontal
                       else f"{get_column_letter(left_col)}{top_row + 1 + k}")
                      for k in range(size)]
            for k, coord in enumerate(coords):
                if k == n_side:
                    formula = f"={base_ref}"
                elif k > n_side:
                    formula = f"={coords[k - 1]}+{step_ref}"
                else:
                    formula = f"={coords[k + 1]}-{step_ref}"
                self.set_cell(ws, coord, formula, CellRole.FORMULA, axis_fmt, bold=True)

        axis(True, col_base_ref, col_step_ref)
        axis(False, row_base_ref, row_step_ref)
        row_letter = get_column_letter(left_col)
        for i in range(size):
            for j in range(size):
                col_letter = get_column_letter(left_col + 1 + j)
                self.set_cell(ws, f"{col_letter}{top_row + 1 + i}",
                              cell_formula(f"${row_letter}{top_row + 1 + i}",
                                           f"{col_letter}${top_row}"),
                              CellRole.FORMULA, value_fmt)
        return f"${get_column_letter(left_col + 1 + n_side)}${top_row + 1 + n_side}"

    # -- Ratios section (deterministic; fixes the "lazy ratios" failure) -----

    def build_ratios(self, ws: Worksheet, start_row: int, first_col: int,
                     n_cols: int, ref: dict[str, str],
                     wacc_ref: Optional[str] = None,
                     days_ref: Optional[str] = None,
                     window: Optional[int] = None) -> tuple[int, list[str]]:
        """Write the FULL Ratios section (blocks A-G of ratios-analytics.md).

        ``window``: periodos por ventana anual. OBLIGATORIO en hojas con header
        trimestral (``Operating``): el builder no puede saber si ``ref`` apunta
        a flujos de un trimestre o a filas UDM que el caller ya construyo, y
        adivinar por el header anualiza dos veces en silencio.
          - ``window=4``: el caller pasa flujos de UN trimestre; el builder los
            suma (t-3..t, exactamente 4 trimestres) y promedia los saldos de
            los 5 cierres t-4..t (apertura y cierre del periodo de 12 meses,
            igual que (inicio+fin)/2 en modo anual).
          - ``window=1``: el caller ya pasa filas UDM (o quiere razones de un
            solo trimestre); flujos tal cual, saldos (t-1+t)/2.
        Sin ``window`` en hoja trimestral: ValueError. En hojas anuales el
        default es 1.

        ``days_ref``: named range de dias para DSO/DIO/DPO/CCC; default
        ``DAYS_YEAR``. Stock promedio y flujo deben cubrir la MISMA ventana:
        ``DAYS_QUARTER`` solo con ``window=1`` explicito (flujo de un
        trimestre); con ventana UDM es un error y se rechaza.

        ``ref`` maps canon line -> row reference with ``{c}`` for the period
        column, e.g. {"rev": "Operating!{c}27", ...}. Required canons: rev,
        cogs, gross, ebit, ebt, ni, interest, tax, ta, equity, cash, ar, inv,
        ap, ca, cl, debt, re, cfo, da. Missing canons skip their rows
        (returned in the skipped list) — but check F13 fails if the section is
        incomplete, so a skip is visible, never silent.
        """
        if window is None:
            if _has_quarter_header(ws):
                raise ValueError(
                    f"build_ratios: la hoja {ws.title} tiene header trimestral y "
                    "falta window. Pasa window=4 si ref apunta a flujos de un "
                    "trimestre (el builder suma t-3..t), o window=1 si ref ya "
                    "apunta a filas UDM.")
            window = 1
        if window < 1:
            raise ValueError(f"build_ratios: window={window} invalido")
        if days_ref is None:
            days_ref = "DAYS_YEAR"
        if window > 1 and days_ref == "DAYS_QUARTER":
            raise ValueError(
                "build_ratios: con ventana UDM (window=4) los flujos cubren 12 "
                "meses y los dias van con DAYS_YEAR; DAYS_QUARTER solo con "
                "window=1 explicito (flujo de un trimestre).")
        skipped: list[str] = []
        r = start_row
        tag = " [UDM]" if window > 1 else ""
        # Primera columna con ventana completa: los saldos se promedian sobre
        # window+1 cierres (t-window..t), asi que arranca en i = window. Las
        # previas quedan VACIAS (F15 admite hasta 4).
        first_i = window
        canons = [k for k, v in ref.items()
                  if not k.endswith("_p") and isinstance(v, str) and "{c}" in v]

        def at(canon: str, i: int) -> str:
            return ref[canon].replace("{c}", get_column_letter(first_col + i))

        def parts_for(i: int) -> dict[str, str]:
            parts: dict[str, str] = {}
            flow_span = range(i - window + 1, i + 1)     # t-window+1..t
            stock_span = range(i - window, i + 1)        # t-window..t (incluye apertura)
            for canon in canons:
                if window > 1 and canon in RATIO_FLOW_CANONS:
                    parts[canon] = "SUM(" + ",".join(at(canon, j) for j in flow_span) + ")"
                else:
                    parts[canon] = at(canon, i)
                parts["avg_" + canon] = ("AVERAGE(" + ",".join(at(canon, j) for j in stock_span)
                                         + ")")
            return parts

        def row_out(label: str, template: str, fmt: NumFmt,
                    needs: tuple[str, ...]) -> None:
            nonlocal r
            if any(k not in canons for k in needs):
                skipped.append(label)
                return
            ws.cell(row=r, column=2, value=label + tag).font = Font(
                name=FONT_NAME, size=11)
            for i in range(first_i, n_cols):
                col = get_column_letter(first_col + i)
                self.set_cell(ws, f"{col}{r}", template.format(**parts_for(i)),
                              CellRole.FORMULA, fmt)
            r += 1

        def header(title: str) -> None:
            nonlocal r
            self.section_header(ws, r, title)
            r += 1

        header("DuPont")
        row_out("Margen neto (NI/Ventas)", '=IF({rev}=0,"",{ni}/{rev})', NumFmt.PCT1, ("ni", "rev"))
        row_out("Rotacion de activos (Ventas/Activos prom.)",
                '=IF({avg_ta}=0,"",{rev}/{avg_ta})', NumFmt.DEC2, ("rev", "ta"))
        row_out("Apalancamiento (Activos/Capital prom.)",
                '=IF({avg_equity}=0,"",{avg_ta}/{avg_equity})', NumFmt.DEC2, ("ta", "equity"))
        row_out("ROE DuPont 3", '=IF({avg_equity}=0,"",{ni}/{avg_equity})', NumFmt.PCT1, ("ni", "equity"))
        row_out("Carga fiscal (NI/EBT)", '=IF({ebt}=0,"",{ni}/{ebt})', NumFmt.PCT1, ("ni", "ebt"))
        row_out("Carga de interes (EBT/EBIT)", '=IF({ebit}=0,"",{ebt}/{ebit})', NumFmt.PCT1, ("ebt", "ebit"))
        row_out("Margen EBIT (EBIT/Ventas)", '=IF({rev}=0,"",{ebit}/{rev})', NumFmt.PCT1, ("ebit", "rev"))
        row_out("ROE DuPont 5 (producto)",
                '=IF(OR({ebt}=0,{ebit}=0,{rev}=0,{avg_equity}=0),"",'
                '{ni}/{ebt}*{ebt}/{ebit}*{ebit}/{rev}*{rev}/{avg_ta}'
                '*{avg_ta}/{avg_equity})', NumFmt.PCT1,
                ("ni", "ebt", "ebit", "rev", "ta", "equity"))
        r += 1

        header("ROIC y economic profit")
        row_out("Tasa efectiva (tax/EBT)", '=IF({ebt}=0,"",{tax}/{ebt})', NumFmt.PCT1, ("tax", "ebt"))
        row_out("NOPAT (EBIT x (1-t))", '=IF({ebt}=0,"",{ebit}*(1-{tax}/{ebt}))', NumFmt.NUM, ("ebit", "tax", "ebt"))
        # Capital invertido PROMEDIO (consistente con el resto de ratios que
        # usan denominadores de balance promediados).
        _ic = "({avg_debt}+{avg_equity}-{avg_cash})"
        row_out("Capital invertido promedio (deuda+capital-caja)", "=" + _ic,
                NumFmt.NUM, ("debt", "equity", "cash"))
        # Guardas explicitas: capital invertido <= 0 (caja > deuda+capital,
        # caso real en emisoras con caja neta enorme) hace el ROIC absurdo —
        # se reporta "n/s" (no significativo), jamas un numero inflado.
        row_out("ROIC",
                '=IF(OR(' + _ic + '<=0,{ebt}=0),"n/s",{ebit}*(1-{tax}/{ebt})/' + _ic + ')',
                NumFmt.PCT1, ("debt", "equity", "cash", "ebit", "tax", "ebt"))
        if wacc_ref:
            row_out(f"Economic profit (spread vs WACC {wacc_ref})",
                    '=IF(OR(' + _ic + '<=0,{ebt}=0),"n/s",'
                    '({ebit}*(1-{tax}/{ebt})/' + _ic + '-' + wacc_ref + ')*' + _ic + ')',
                    NumFmt.NUM, ("ebit", "tax", "ebt", "debt", "equity", "cash"))
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
        row_out("DSO (dias)", '=IF({rev}=0,"",{avg_ar}/{rev}*' + days_ref + ')', NumFmt.DEC2, ("ar", "rev"))
        row_out("DIO (dias)", '=IF({cogs}=0,"",{avg_inv}/{cogs}*' + days_ref + ')', NumFmt.DEC2, ("inv", "cogs"))
        row_out("DPO (dias)", '=IF({cogs}=0,"",{avg_ap}/{cogs}*' + days_ref + ')', NumFmt.DEC2, ("ap", "cogs"))
        row_out("CCC (DSO+DIO-DPO)",
                '=IF(OR({cogs}=0,{rev}=0),"",{avg_ar}/{rev}*' + days_ref
                + '+{avg_inv}/{cogs}*' + days_ref
                + '-{avg_ap}/{cogs}*' + days_ref + ')',
                NumFmt.DEC2, ("ar", "inv", "ap", "rev", "cogs"))
        r += 1

        header("Apalancamiento operativo y calidad")
        row_out("DFL (EBIT/(EBIT-interes))", '=IF(({ebit}-ABS({interest}))=0,"",{ebit}/({ebit}-ABS({interest})))', NumFmt.DEC2, ("ebit", "interest"))
        row_out("CFO / NI (calidad de utilidades)", '=IF({ni}=0,"",{cfo}/{ni})', NumFmt.DEC2, ("cfo", "ni"))
        row_out("Accruals proxy (NI-CFO)/Activos prom.",
                '=IF({avg_ta}=0,"",({ni}-{cfo})/{avg_ta})', NumFmt.PCT1, ("ni", "cfo", "ta"))
        return r, skipped
