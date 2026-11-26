"""Contrato de formato del xlsx: los UNICOS valores permitidos.

Paleta (``Color``), formatos numericos (``NumFmt``), roles de celda
(``CellRole``) y su color de fuente, sellos del builder, contratos de hojas y
de la tab Ratios, slots de marca y limites de escaneo del audit. Lo leen el
builder (``xlsx_styler``) y el audit (``xlsx_audit``); los checks F hacen
cumplir las whitelists. Se importa a traves de la fachada ``xlsx_builder``.
"""

# SIN ``from __future__ import annotations`` a proposito: ver el encabezado
# de ``xlsx_builder.py`` (el builder tambien se carga por ruta).

from dataclasses import dataclass
from enum import Enum, auto
from typing import Optional

# ---------------------------------------------------------------------------
# Contract constants (the ONLY allowed values; checks F enforce the whitelists)
# ---------------------------------------------------------------------------

BUILDER_STAMP_KEY = "research_analyst_builder"
BUILDER_STAMP_VALUE = "xlsx_builder v1"
TRANSITION_STAMP_KEY = "transition"
_UNSET = object()   # centinela: None es un valor de celda valido


def transition_year_of(keywords: Optional[str]) -> Optional[int]:
    """Anio de transicion sellado en el libro (``set_transition_year``) o None."""
    import re
    match = re.search(rf"(?:^|;){TRANSITION_STAMP_KEY}=(\d{{4}})", keywords or "")
    return int(match.group(1)) if match else None

FONT_NAME = "Aptos Narrow"

# Vinculo fila <-> capa de captura (check C9). Un nombre definido con alcance
# de hoja por linea: Excel ajusta su rango al insertar o mover filas, asi que
# el vinculo sobrevive a una edicion manual del libro.
CAPTURE_NAME_PREFIX = "CAPTURE."
# Lineas que C9 usa como referencia de materialidad: ventas y activos totales
# (los nombres que emite xbrl_fetch.CONCEPT_MAP).
CAPTURE_REFERENCE_LINES = ("is_ns_total", "bs_ta")


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

# Canons de FLUJO: en ventana UDM se SUMAN (t-3..t); el resto son SALDOS y se
# promedian sobre la misma ventana (build_ratios).
RATIO_FLOW_CANONS = frozenset({"rev", "cogs", "gross", "ebit", "ebt", "ni",
                               "interest", "tax", "cfo", "da"})


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


# Limites de escaneo del audit. Una hoja mas grande NO se audita en silencio
# a medias: el check F20 falla y nombra la hoja.
_MAX_SCAN_ROWS = 5000
_MAX_SCAN_COLS = 400
