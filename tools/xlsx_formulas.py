"""Texto de formulas de Excel: validarlo (check F21) y generarlo (valuacion).

``lint_formula`` revisa aridad, parentesis, prefijos ``_xlfn.`` y operaciones
sobre rangos sin abrir Excel; el builder la corre en cada ``set_cell``. Las
formulas de valuacion (DCF, market cap de comps, banderas de rango) son texto
que el builder coloca en la celda. Se importa a traves de la fachada
``xlsx_builder``.
"""

# SIN ``from __future__ import annotations`` a proposito: ver el encabezado
# de ``xlsx_builder.py`` (el builder tambien se carga por ruta).

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError

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
