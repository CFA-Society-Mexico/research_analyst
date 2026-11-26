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
"""

# SIN ``from __future__ import annotations`` a proposito: con anotaciones
# como texto, ``@dataclass`` consulta ``sys.modules[cls.__module__]`` y truena
# si el modulo se cargo por ruta (``importlib.util.spec_from_file_location`` +
# ``exec_module``) sin registrarlo antes. Con anotaciones evaluadas el builder
# carga igual por ``import`` que por ruta (tests/test_formula_lint.py).

import csv
import sys
import types
from dataclasses import dataclass
from enum import Enum, auto
from pathlib import Path
from typing import Callable, Iterable, Optional

from datetime import date

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

# --- Fachada ------------------------------------------------------------------
# Los submodulos viven junto a este archivo. Al cargar la fachada por ruta sin
# tools/ en sys.path (importlib), los imports de abajo los encuentran igual.
_TOOLS_DIR = str(Path(__file__).resolve().parent)
if _TOOLS_DIR not in sys.path:
    sys.path.append(_TOOLS_DIR)

import xlsx_contract  # noqa: E402

_SUBMODULES = (xlsx_contract,)

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


_THIN = Side(style="thin")
_DOUBLE = Side(style="double")

# ---------------------------------------------------------------------------
# Lint de formulas (check F21): lo que Excel descarta o evalua mal
# ---------------------------------------------------------------------------
# openpyxl escribe cualquier texto que empiece con "=" sin validarlo. Excel si
# valida al abrir: una funcion con aridad invalida manda el libro a modo
# reparacion y BORRA la formula (cobertura TSMC: MAX(ABS(a,b,c)) borro siete
# checks vivos y el audit F seguia 19/19 en verde). El lint corre sin Excel.

# (minimo, maximo) de argumentos. Solo funciones de aridad cerrada o de uso
# comun en modelos: una funcion ausente no se valida (no se adivina).
_MANY = 255
FUNCTION_ARITY: dict[str, tuple[int, int]] = {
    "ABS": (1, 1), "SQRT": (1, 1), "LN": (1, 1), "EXP": (1, 1), "LOG10": (1, 1),
    "INT": (1, 1), "SIGN": (1, 1), "NOT": (1, 1), "ISNUMBER": (1, 1),
    "ISBLANK": (1, 1), "ISERROR": (1, 1), "ISERR": (1, 1), "ISNA": (1, 1),
    "ISTEXT": (1, 1), "N": (1, 1), "YEAR": (1, 1), "MONTH": (1, 1), "DAY": (1, 1),
    "LEN": (1, 1), "TRIM": (1, 1), "VALUE": (1, 1), "UPPER": (1, 1), "LOWER": (1, 1),
    "NA": (0, 0), "PI": (0, 0), "TRUE": (0, 0), "FALSE": (0, 0),
    "ROUND": (2, 2), "ROUNDUP": (2, 2), "ROUNDDOWN": (2, 2), "POWER": (2, 2),
    "MOD": (2, 2), "IFERROR": (2, 2), "IFNA": (2, 2), "COUNTIF": (2, 2),
    "EDATE": (2, 2), "EOMONTH": (2, 2), "TEXT": (2, 2), "SMALL": (2, 2),
    "LARGE": (2, 2), "PERCENTILE": (2, 2), "QUARTILE": (2, 2),
    "DATE": (3, 3), "MID": (3, 3), "XNPV": (3, 3),
    "LOG": (1, 2), "LEFT": (1, 2), "RIGHT": (1, 2), "IRR": (1, 2),
    "IF": (2, 3), "SUMIF": (2, 3), "AVERAGEIF": (2, 3), "MATCH": (2, 3),
    "YEARFRAC": (2, 3), "XIRR": (2, 3), "RANK": (2, 3),
    "VLOOKUP": (3, 4), "HLOOKUP": (3, 4), "INDEX": (2, 4),
    "PMT": (3, 5), "PV": (3, 5), "FV": (3, 5), "NPER": (3, 5), "RATE": (3, 6),
    "SUM": (1, _MANY), "AVERAGE": (1, _MANY), "MAX": (1, _MANY), "MIN": (1, _MANY),
    "MEDIAN": (1, _MANY), "COUNT": (1, _MANY), "COUNTA": (1, _MANY),
    "PRODUCT": (1, _MANY), "AND": (1, _MANY), "OR": (1, _MANY),
    "HARMEAN": (1, _MANY), "GEOMEAN": (1, _MANY), "STDEV": (1, _MANY),
    "VAR": (1, _MANY), "SUMPRODUCT": (1, _MANY), "CONCATENATE": (1, _MANY),
    "NPV": (2, _MANY), "CHOOSE": (2, _MANY), "SUMIFS": (3, _MANY),
    "COUNTIFS": (2, _MANY), "AVERAGEIFS": (3, _MANY),
}

# Funciones posteriores a Excel 2007: en el archivo deben llevar el prefijo
# ``_xlfn.`` (doc. de openpyxl). Sin el, Excel las carga como #NAME?.
XLFN_FUNCTIONS = frozenset({
    "XLOOKUP", "XMATCH", "IFS", "SWITCH", "MAXIFS", "MINIFS", "CONCAT",
    "TEXTJOIN", "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE",
    "SEQUENCE", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "PERCENTILE.INC",
    "PERCENTILE.EXC", "QUARTILE.INC", "QUARTILE.EXC", "RANK.EQ", "RANK.AVG",
})

# Evaluan sus argumentos como arreglo sin necesitar formula de arreglo.
_ARRAY_NATIVE = frozenset({"SUMPRODUCT"})
_XL_PREFIXES = ("_XLFN._XLWS.", "_XLFN.")


def _function_name(token_value: str) -> tuple[str, bool]:
    """('ABS', prefijado) a partir del token 'ABS(' o '_xlfn.XLOOKUP('."""
    name = token_value[:-1].upper()
    for prefix in _XL_PREFIXES:
        if name.startswith(prefix):
            return name[len(prefix):], True
    return name, False


def _is_multicell(operand: str) -> bool:
    """True si el operando es un rango de varias celdas (A1:B9, A:A, Hoja!A1:A9)."""
    ref = operand.rsplit("!", 1)[-1].replace("$", "")
    return ":" in ref


def lint_formula(formula: object, array: bool = False) -> list[str]:
    """Problemas que harian a Excel descartar o evaluar mal la formula.

    (a) aridad invalida en funciones comunes -> Excel repara el libro y BORRA
        la formula; (b) parentesis desbalanceados; (c) funcion moderna sin
        prefijo ``_xlfn.`` -> #NAME?; (d) operacion de arreglo (un rango de
        varias celdas como operando de un operador) en una formula NORMAL ->
        Excel aplica interseccion implicita y el resultado no es el del
        arreglo (``MEDIAN(IF(rango=1,...))`` sale "n/a"). (d) no aplica dentro
        de SUMPRODUCT ni cuando ``array=True`` (formula de arreglo declarada).
    Devuelve [] si no es formula o si esta limpia.
    """
    if not isinstance(formula, str) or not formula.startswith("=") or len(formula) < 2:
        return []
    try:
        tokens = [t for t in Tokenizer(formula).items if t.type != Token.WSPACE]
    except (TokenizerError, IndexError):
        # IndexError: el tokenizer de openpyxl saca de una pila vacia cuando
        # sobra un parentesis de cierre.
        return ["parentesis o comillas desbalanceados"]
    problems: list[str] = []
    # Pila de marcos abiertos: [nombre de funcion o None, separadores, hay contenido]
    stack: list[list] = []
    for i, token in enumerate(tokens):
        opens = token.subtype == Token.OPEN
        closes = token.subtype == Token.CLOSE
        if token.type == Token.FUNC and opens:
            name, prefixed = _function_name(token.value)
            if name in XLFN_FUNCTIONS and not prefixed:
                problems.append(f"{name}: funcion moderna sin prefijo _xlfn. (Excel la carga como #NAME?)")
            if stack:
                stack[-1][2] = True
            stack.append([name, 0, False])
            continue
        if token.type in (Token.PAREN, Token.ARRAY) and opens:
            if stack:
                stack[-1][2] = True
            stack.append([None, 0, False])
            continue
        if closes and token.type in (Token.FUNC, Token.PAREN, Token.ARRAY):
            if not stack:
                return problems + ["parentesis de cierre de mas"]
            name, separators, has_content = stack.pop()
            if token.type == Token.FUNC and name in FUNCTION_ARITY:
                count = separators + 1 if (separators or has_content) else 0
                low, high = FUNCTION_ARITY[name]
                if not low <= count <= high:
                    expected = str(low) if low == high else f"{low} a {high}"
                    problems.append(f"{name} con {count} argumentos (admite {expected}): "
                                    "Excel repara el libro y borra la formula")
            continue
        if stack:
            stack[-1][2] = True
            if token.type == Token.SEP and token.subtype == Token.ARG:
                stack[-1][1] += 1
        if (not array and token.type == Token.OPERAND and token.subtype == Token.RANGE
                and _is_multicell(token.value)
                and not any(frame[0] in _ARRAY_NATIVE for frame in stack)):
            before = tokens[i - 1] if i else None
            after = tokens[i + 1] if i + 1 < len(tokens) else None
            operated = ((before is not None and before.type in (Token.OP_IN, Token.OP_PRE)
                         and before.value != ",")
                        or (after is not None and after.type in (Token.OP_IN, Token.OP_POST)
                            and after.value != ","))
            if operated:
                problems.append(f"operacion de arreglo sobre {token.value} en formula normal "
                                "(Excel aplica interseccion implicita): usa set_array_formula "
                                "o SUMPRODUCT")
    if stack:
        problems.append("parentesis sin cerrar")
    return problems


# ---------------------------------------------------------------------------
# Formulas de valuacion (texto de formula; el builder las coloca)
# ---------------------------------------------------------------------------
# Viven aqui, y no en el script de cada cobertura, porque una formula de
# valuacion mal despejada no la detecta ningun check de formato. Ninguna lleva
# literales distintos de 0 y 1 fuera de un check (S4) ni operaciones de
# arreglo fuera de SUMPRODUCT (F21). Las referencias llegan como texto A1 o
# nombre definido.

NO_GROWTH_BELOW_WACC = "ninguna g < WACC reproduce el TV"
NO_REAL_SOLUTION = "sin solucion real"
NOT_MEANINGFUL = "n/s"
REVIEW_FLAG = "REVISAR"


def dcf_per_share_formula(*, wacc: str, g: str, fcff_range: str, exponent_range: str,
                          nopat_n: str, ronic: str, tv_exponent: str,
                          net_debt: str, shares: str) -> str:
    """Valor por accion del DCF en UNA celda, como funcion de (WACC, g).

    Es la forma cerrada que sustituye al data table de Excel (openpyxl no
    puede crearlo): ``(VP de los FCFF explicitos + VP del terminal normalizado
    - deuda neta) / acciones``. Los exponentes vienen de la fila del DCF, asi
    que respeta stub y mid-year. SUMPRODUCT evalua el arreglo sin CSE.
    """
    explicit = f"SUMPRODUCT({fcff_range}/(1+{wacc})^{exponent_range})"
    terminal = (f"{nopat_n}*(1+{g})*(1-{g}/{ronic})/({wacc}-{g})"
                f"/(1+{wacc})^{tv_exponent}")
    return (f'=IF({wacc}<={g},"{NOT_MEANINGFUL}",'
            f"({explicit}+{terminal}-{net_debt})/{shares})")


def comps_market_cap_formula(*, price: str, shares_per_ads: str, shares: str,
                             fx_to_reporting: str) -> str:
    """Market cap de un comparable EN LA MONEDA DE REPORTE del modelo.

    ``precio / acciones por ADS x acciones ordinarias x FX``. El precio de un
    ADR viene por ADS y en la moneda de cotizacion; multiplicarlo directo por
    las acciones ordinarias mezcla unidades y monedas (UMC en la cobertura de
    TSMC: EV/EBITDA 1.0x en vez de 10.1x).
    """
    return f"={price}/{shares_per_ads}*{shares}*{fx_to_reporting}"


def multiple_range_flag_formula(multiple: str, low: str, high: str) -> str:
    """Marca ``REVISAR`` si un multiplo cae fuera de [low, high] (check D12).

    Un EV/EBITDA de 1.0x casi nunca es una oportunidad: es un error de
    unidades o de moneda. Los limites son constantes con nombre (defaults del
    plugin 2x y 60x), no literales."""
    return (f'=IF(NOT(ISNUMBER({multiple})),"n/a",'
            f'IF(OR({multiple}<{low},{multiple}>{high}),"{REVIEW_FLAG}","OK"))')


def terminal_multiple_alarm_formula(implied_multiple: str, comps_median: str,
                                    floor_ratio: str) -> str:
    """Check D11: multiplo implicito del terminal Gordon vs mediana de comps.

    Por debajo de ``floor_ratio`` x la mediana (default 0.5) el terminal esta
    perpetuando algo que no es de estado estable -tipicamente capex de
    crecimiento- y se marca ``Error``; no es un dato informativo."""
    return (f'=IF(OR(NOT(ISNUMBER({implied_multiple})),NOT(ISNUMBER({comps_median}))),"n/a",'
            f'IF({implied_multiple}<{floor_ratio}*{comps_median},"Error","OK"))')


def center_check_formula(center: str, model_value: str) -> str:
    """La celda central de la cuadricula de sensibilidad reproduce el modelo."""
    return (f'=IF(OR(NOT(ISNUMBER({center})),NOT(ISNUMBER({model_value}))),"Error",'
            f'IF(ABS({center}-{model_value})>0.0001*MAX(1,ABS({model_value})),"Error","OK"))')


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
        reinvestment = self._value_row(
            ws, row, col, "Tasa de reinversion terminal (g / RONIC)",
            f'=IF({ronic_ref}<=0,"{NOT_MEANINGFUL}",{g_ref}/{ronic_ref})', NumFmt.PCT1)
        flow = self._value_row(
            ws, row + 1, col, "Flujo terminal normalizado (NOPAT_n x (1+g) x (1 - g/RONIC))",
            f'=IF({ronic_ref}<=0,"{NOT_MEANINGFUL}",'
            f"{nopat_ref}*(1+{g_ref})*(1-{g_ref}/{ronic_ref}))", NumFmt.NUM)
        tv = self._value_row(
            ws, row + 2, col, "TV Gordon (flujo terminal / (WACC - g))",
            f'=IF(OR({wacc_ref}<={g_ref},NOT(ISNUMBER({flow}))),"{NOT_MEANINGFUL}",'
            f"{flow}/({wacc_ref}-{g_ref}))", NumFmt.NUM)
        return {"reinvestment": reinvestment, "flow": flow, "tv": tv,
                "next_row": str(row + 3)}

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
