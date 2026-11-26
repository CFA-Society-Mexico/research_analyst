"""Deck de cobertura por codigo: numeros con procedencia, laminas y medicion.

Doctrina: en un deck, un numero no es texto. Sale del modelo recalculado, de
un hecho citado o de aritmetica entre ambos, y viaja con su origen hasta el
HTML (``data-src``). Un numero suelto en el texto impide generar el deck.

Este modulo es la base comun; el guion, la narrativa y las laminas a medida
viven en el ``research/notes/build_deck.py`` de cada cobertura.

Uso como biblioteca (desde el script de la cobertura):
    deck = Deck(root=ROOT, model="model/x-model_2026-01-01_v1.xlsx",
                ticker="X", as_of="2026-01-01")
    rev = deck.series("Annual", "Ingresos")
    deck.slide(title=f"Ventas crecen {rev.growth()[2026]:+.0%}", ...)
    deck.build()

Uso por consola:
    python tools/deck_builder.py demo <carpeta>
    python tools/deck_builder.py audit <deck.html> <modelo.xlsx>
                                 [--facts hechos.json] [--design brand/DESIGN.md]
                                 [--report reporte.md]
    python tools/deck_builder.py measure <deck.html> [--shots carpeta]
    python tools/deck_builder.py claims <deck.html> <salida.md>

Exit codes (los de model_check): 0 = verde; 1 = alguna falla; 2 = uso
incorrecto; 3 = sin fallas pero con [pendiente] (sin navegador para medir).

Solo lectura sobre el modelo. Consola ASCII-only.
"""

from __future__ import annotations

import argparse
import ast
import base64
import html
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Iterator, Optional, Sequence, Union

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

sys.path.insert(0, str(Path(__file__).resolve().parent))

import model_check as M  # noqa: E402
import xlsx_builder as B  # noqa: E402
from xlsx_builder import Finding  # noqa: E402


class DeckError(Exception):
    """El deck no se puede generar o auditar; el mensaje dice que corregir."""


def _ascii(text: object) -> str:
    """Texto seguro para la consola cp1252 (CONTRIBUTING 4.3)."""
    return str(text).encode("ascii", "replace").decode("ascii")


# ---------------------------------------------------------------------------
# Numero con procedencia
# ---------------------------------------------------------------------------
# Un Num formateado emite un marcador con caracteres de uso privado; el render
# lo convierte en <span data-src data-fmt>. Asi el origen sobrevive a un
# f-string sin que el script tenga que hacer nada.

_M_NUM, _M_LIT, _M_SEP, _M_END = "\ue000", "\ue001", "\ue002", "\ue003"
_MARKER = re.compile(
    f"{_M_NUM}([^{_M_SEP}]*){_M_SEP}([^{_M_SEP}]*){_M_SEP}([^{_M_END}]*){_M_END}"
    f"|{_M_LIT}([^{_M_SEP}]*){_M_SEP}([^{_M_END}]*){_M_END}")


def _const(value: float) -> str:
    number = float(value)
    return str(int(number)) if number.is_integer() and abs(number) < 1e15 else repr(number)


@dataclass(frozen=True)
class Num:
    """Valor + expresion de origen (``Annual!H14``, ``fact:clave`` o derivada)."""

    value: float
    src: str

    def _combine(self, other: object, symbol: str,
                 fn: Callable[[float, float], float], reverse: bool = False) -> "Num":
        if isinstance(other, Num):
            pair = (other.value, other.src)
        elif isinstance(other, (int, float)) and not isinstance(other, bool):
            pair = (float(other), _const(other))
        else:
            return NotImplemented
        left, right = (pair, (self.value, self.src)) if reverse else ((self.value, self.src), pair)
        try:
            result = fn(left[0], right[0])
        except ZeroDivisionError as exc:
            raise DeckError(f"division entre cero: {left[1]}{symbol}{right[1]}") from exc
        return Num(result, f"({left[1]}{symbol}{right[1]})")

    def __add__(self, other: object) -> "Num":
        return self._combine(other, "+", lambda a, b: a + b)

    def __radd__(self, other: object) -> "Num":
        return self._combine(other, "+", lambda a, b: a + b, reverse=True)

    def __sub__(self, other: object) -> "Num":
        return self._combine(other, "-", lambda a, b: a - b)

    def __rsub__(self, other: object) -> "Num":
        return self._combine(other, "-", lambda a, b: a - b, reverse=True)

    def __mul__(self, other: object) -> "Num":
        return self._combine(other, "*", lambda a, b: a * b)

    def __rmul__(self, other: object) -> "Num":
        return self._combine(other, "*", lambda a, b: a * b, reverse=True)

    def __truediv__(self, other: object) -> "Num":
        return self._combine(other, "/", lambda a, b: a / b)

    def __rtruediv__(self, other: object) -> "Num":
        return self._combine(other, "/", lambda a, b: a / b, reverse=True)

    def __pow__(self, other: object) -> "Num":
        return self._combine(other, "**", lambda a, b: a ** b)

    def __neg__(self) -> "Num":
        return Num(-self.value, f"(-{self.src})")

    def __abs__(self) -> "Num":
        return Num(abs(self.value), f"abs({self.src})")

    def __float__(self) -> float:
        return float(self.value)

    def __format__(self, spec: str) -> str:
        return f"{_M_NUM}{self.src}{_M_SEP}{spec}{_M_SEP}{format(self.value, spec)}{_M_END}"

    def __str__(self) -> str:
        return format(self, "")


def lit(text: str, why: str) -> str:
    """Valvula: un literal numerico declarado. Sale listado en el reporte."""
    if not why.strip():
        raise DeckError(f"lit({text!r}): falta el porque (why)")
    return f"{_M_LIT}{why}{_M_SEP}{text}{_M_END}"


def _esc(text: str) -> str:
    """Escapa texto para HTML sin tocar los marcadores de Num y lit."""
    out, last = [], 0
    for match in _MARKER.finditer(text):
        out.append(html.escape(text[last:match.start()], quote=False))
        out.append(match.group(0))
        last = match.end()
    out.append(html.escape(text[last:], quote=False))
    return "".join(out)


def render_markers(document: str) -> str:
    """Convierte marcadores en <span>/<tspan> con su origen. Dentro de una
    etiqueta (un atributo) el marcador se reduce a su texto."""
    found: list[re.Match[str]] = []

    def stash(match: re.Match[str]) -> str:
        found.append(match)
        return f"{_M_NUM}{len(found) - 1}{_M_END}"

    protected = _MARKER.sub(stash, document)
    token = re.compile(f"{_M_NUM}(\\d+){_M_END}")
    svg_depth = 0
    out: list[str] = []
    for piece in re.split(r"(<[^>]*>)", protected):
        if piece.startswith("<"):
            if re.match(r"<svg\b", piece):
                svg_depth += 1
            elif piece.startswith("</svg"):
                svg_depth -= 1
            out.append(token.sub(
                lambda m: html.escape(found[int(m.group(1))].group(3)
                                      or found[int(m.group(1))].group(5) or ""), piece))
            continue
        tag = "tspan" if svg_depth > 0 else "span"

        def expand(m: re.Match[str], tag: str = tag) -> str:
            src, fmt, text, why, literal = found[int(m.group(1))].groups()
            if literal is not None:
                return (f'<{tag} class="lit" data-why="{html.escape(why, quote=True)}">'
                        f"{html.escape(literal)}</{tag}>")
            return (f'<{tag} class="n" data-src="{html.escape(src, quote=True)}" '
                    f'data-fmt="{html.escape(fmt, quote=True)}">{html.escape(text)}</{tag}>')

        out.append(token.sub(expand, piece))
    return "".join(out)


# ---------------------------------------------------------------------------
# Lectura del modelo recalculado
# ---------------------------------------------------------------------------

PeriodKey = Union[int, str]


def _sheet_ref(sheet: str) -> str:
    return sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else f"'{sheet}'"


class Series:
    """Una fila del modelo sobre el eje de periodos: ``serie[2027]`` -> Num."""

    def __init__(self, label: str, points: dict[PeriodKey, Num],
                 actual: dict[PeriodKey, bool]) -> None:
        self.label = label
        self._points = dict(points)
        self._actual = dict(actual)

    @property
    def keys(self) -> list[PeriodKey]:
        return list(self._points)

    def __getitem__(self, key: PeriodKey) -> Num:
        if key not in self._points:
            raise DeckError(f"'{self.label}' no tiene valor en {key!r}; "
                            f"periodos: {self.keys[0]!r}..{self.keys[-1]!r}")
        return self._points[key]

    def __len__(self) -> int:
        return len(self._points)

    def is_actual(self, key: PeriodKey) -> bool:
        return self._actual[key]

    def _derived(self, label: str, points: dict[PeriodKey, Num]) -> "Series":
        return Series(label, points, {k: self._actual[k] for k in points})

    def growth(self) -> "Series":
        """Crecimiento contra el periodo anterior (desde el segundo)."""
        keys = self.keys
        return self._derived(f"{self.label} (crecimiento)", {
            key: self._points[key] / self._points[prev] - 1
            for prev, key in zip(keys, keys[1:])})

    def over(self, other: "Series") -> "Series":
        """Razon periodo a periodo contra otra serie (p. ej. un margen)."""
        return self._derived(f"{self.label} / {other.label}", {
            key: self._points[key] / other[key] for key in self.keys if key in other.keys})

    def cagr(self, first: int, last: int) -> Num:
        """Crecimiento compuesto anual entre dos anios del eje."""
        if not (isinstance(first, int) and isinstance(last, int) and last > first):
            raise DeckError(f"cagr: se esperan dos anios en orden, no {first!r}, {last!r}")
        return (self[last] / self[first]) ** (1.0 / (last - first)) - 1


class ModelBook:
    """El xlsx recalculado, leido por etiqueta. Solo valores calculados."""

    def __init__(self, path: Path) -> None:
        if not path.is_file():
            raise DeckError(f"no existe el modelo: {_ascii(path)}")
        self.path = path
        self._values = load_workbook(path, data_only=True)
        self._formulas = load_workbook(path, data_only=False)

    def _sheet(self, sheet: str) -> Worksheet:
        if sheet not in self._values.sheetnames:
            raise DeckError(f"el modelo no tiene la hoja '{_ascii(sheet)}'")
        return self._values[sheet]

    def num(self, sheet: str, coord: str) -> Num:
        value = self._sheet(sheet)[coord].value
        if value is None:
            formula = self._formulas[sheet][coord].value
            if isinstance(formula, str) and formula.startswith("="):
                raise DeckError(
                    f"{_ascii(sheet)}!{coord} es una formula sin valor calculado: "
                    f"recalcula con 'python tools/recalc_validate.py <modelo>' y reintenta")
            raise DeckError(f"{_ascii(sheet)}!{coord} esta vacia")
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise DeckError(f"{_ascii(sheet)}!{coord} no es un numero: {_ascii(value)!r}")
        return Num(float(value), f"{_sheet_ref(sheet)}!{coord}")

    def _row_of(self, sheet: str, label: str, section: Optional[str]) -> int:
        ws = self._sheet(sheet)
        first, last = 1, min(ws.max_row, B._MAX_SCAN_ROWS)
        if section is not None:
            bounds = M._section_rows(ws, section.strip().lower())
            if bounds is None:
                raise DeckError(f"{_ascii(sheet)}: no existe la seccion '{_ascii(section)}'")
            first, last = bounds
        rows = [r for r in range(first, last + 1)
                if isinstance(ws.cell(row=r, column=2).value, str)
                and ws.cell(row=r, column=2).value.strip() == label]
        if not rows:
            raise DeckError(f"{_ascii(sheet)}: no hay fila con etiqueta '{_ascii(label)}'")
        if len(rows) > 1:
            raise DeckError(f"{_ascii(sheet)}: la etiqueta '{_ascii(label)}' aparece en las "
                            f"filas {rows}; desambigua con section=")
        return rows[0]

    def series(self, sheet: str, label: str, section: Optional[str] = None) -> Series:
        ws = self._sheet(sheet)
        axis = M._period_axis(ws)
        if not axis:
            raise DeckError(f"{_ascii(sheet)}: no se reconoce un encabezado de periodos")
        row = self._row_of(sheet, label, section)
        points: dict[PeriodKey, Num] = {}
        actual: dict[PeriodKey, bool] = {}
        for period in axis:
            key: PeriodKey = (period.year if period.quarter is None
                              else f"{period.quarter}Q{period.year}")
            cell = ws.cell(row=row, column=period.column)
            formula = self._formulas[sheet].cell(row=row, column=period.column).value
            if cell.value is None and not (isinstance(formula, str) and formula.startswith("=")):
                continue
            if key in points:
                continue
            points[key] = self.num(sheet, cell.coordinate)
            actual[key] = period.actual
        if not points:
            raise DeckError(f"{_ascii(sheet)}: la fila '{_ascii(label)}' no tiene valores")
        return Series(label, points, actual)

    def cell(self, sheet: str, label: str, column: Optional[str] = None,
             section: Optional[str] = None) -> Num:
        ws = self._sheet(sheet)
        row = self._row_of(sheet, label, section)
        if column is not None:
            return self.num(sheet, f"{column.upper()}{row}")
        numeric = [c.coordinate for c in ws[row][2:]
                   if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)]
        if len(numeric) != 1:
            raise DeckError(f"{_ascii(sheet)}: la fila '{_ascii(label)}' tiene "
                            f"{len(numeric)} celdas numericas {numeric[:6]}; indica column=")
        return self.num(sheet, numeric[0])

    def range(self, sheet: str, ref: str) -> list[list[Num]]:
        return [[self.num(sheet, c.coordinate) for c in row] for row in self._sheet(sheet)[ref]]

    def axis_years(self) -> set[int]:
        """Anios presentes en algun encabezado de periodos del libro."""
        return {p.year for ws in self._values.worksheets for p in M._period_axis(ws)}


# ---------------------------------------------------------------------------
# Hechos externos citados
# ---------------------------------------------------------------------------

FACT_TAGS = ("observado", "guidance")
TEXT_SOURCES = frozenset({".md", ".txt", ".html", ".htm", ".csv", ".json"})
_FACT_KEY = re.compile(r"[A-Za-z0-9_.\-]+")


@dataclass(frozen=True)
class Fact:
    key: str
    value: float
    unit: str
    tag: str
    doc: str
    locator: str
    quote: str
    manual: bool   # True: la fuente no es texto; la cita se verifica a mano


def load_facts(path: Path, root: Path) -> dict[str, Fact]:
    """Lee y valida el registro de hechos. ``root`` = carpeta de la cobertura."""
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise DeckError(f"hechos ilegibles ({_ascii(path)}): {_ascii(exc)}") from exc
    facts: dict[str, Fact] = {}
    for raw in payload.get("facts", []):
        missing = [k for k in ("key", "value", "unit", "tag", "doc", "locator", "quote")
                   if k not in raw or raw[k] in ("", None)]
        key = str(raw.get("key", "?"))
        if missing:
            raise DeckError(f"hecho '{_ascii(key)}': faltan campos {missing}")
        if not _FACT_KEY.fullmatch(key):
            raise DeckError(f"hecho '{_ascii(key)}': la clave solo admite letras, digitos, _ . -")
        if key in facts:
            raise DeckError(f"hecho '{_ascii(key)}' duplicado")
        if raw["tag"] not in FACT_TAGS:
            raise DeckError(f"hecho '{_ascii(key)}': tag debe ser uno de {FACT_TAGS}")
        if isinstance(raw["value"], bool) or not isinstance(raw["value"], (int, float)):
            raise DeckError(f"hecho '{_ascii(key)}': value debe ser un numero")
        doc = root / raw["doc"]
        if not doc.is_file():
            raise DeckError(f"hecho '{_ascii(key)}': no existe el documento {_ascii(raw['doc'])}")
        manual = doc.suffix.lower() not in TEXT_SOURCES
        if not manual and raw["quote"] not in doc.read_text(encoding="utf-8", errors="replace"):
            raise DeckError(f"hecho '{_ascii(key)}': la cita {_ascii(raw['quote'])!r} no "
                            f"aparece en {_ascii(raw['doc'])}")
        facts[key] = Fact(key, float(raw["value"]), str(raw["unit"]), raw["tag"],
                          raw["doc"], str(raw["locator"]), str(raw["quote"]), manual)
    return facts


# ---------------------------------------------------------------------------
# Tokens de diseno (brand/DESIGN.md, slots deck_*)
# ---------------------------------------------------------------------------

_SLOT_LINE = re.compile(r"\s*(deck_[a-z0-9_]+)\s*[:=]\s*(\S.*)$")
_HEX = re.compile(r"#[0-9A-Fa-f]{6}")
MIN_SERIES_CONTRAST = 3.0
MIN_TEXT_CONTRAST = 4.5
MIN_CVD_DISTANCE = 40.0
# Machado et al. (2009), deuteranopia, sobre RGB lineal.
_DEUTAN = ((0.367322, 0.860646, -0.227968),
           (0.280085, 0.672501, 0.047413),
           (-0.011820, 0.042940, 0.968881))


@dataclass(frozen=True)
class Design:
    background: str = "#FFFFFF"
    text: str = "#1A1A1A"
    muted: str = "#5B6670"
    negative: str = "#B03A2E"
    # Paleta neutra: cada color pasa MIN_SERIES_CONTRAST sobre blanco y cada par
    # pasa MIN_CVD_DISTANCE (tests/test_deck_builder.py::DesignTests).
    series: tuple[str, ...] = ("#1F4E79", "#4A86BA", "#C8691B", "#2E7D6B", "#7B2D26", "#111111")
    font_heading: str = "Georgia"
    font_body: str = "Arial"
    margin: int = 56
    gap_max: int = 90
    logo: str = ""   # data URI; vacio = sin logo


def _linear(channel: int) -> float:
    c = channel / 255.0
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _rgb(color: str) -> tuple[int, int, int]:
    return int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)


def contrast(a: str, b: str) -> float:
    """Razon de contraste WCAG entre dos colores #RRGGBB (1 a 21)."""
    def luminance(color: str) -> float:
        r, g, bl = (_linear(c) for c in _rgb(color))
        return 0.2126 * r + 0.7152 * g + 0.0722 * bl
    hi, lo = sorted((luminance(a), luminance(b)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


def _deutan(color: str) -> tuple[float, float, float]:
    lin = [_linear(c) for c in _rgb(color)]
    return tuple(255.0 * max(0.0, min(1.0, sum(m * v for m, v in zip(row, lin)))) ** (1 / 2.2)
                 for row in _DEUTAN)  # type: ignore[return-value]


def load_design(path: Optional[Path], root: Optional[Path] = None) -> Design:
    """Lee los slots ``deck_*`` de brand/DESIGN.md. Sin archivo o sin slot:
    estilo neutro. Un valor mal formado es error (no se adivina)."""
    if path is None or not path.is_file():
        return Design()
    slots: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        match = _SLOT_LINE.match(line)
        if match:
            slots[match.group(1)] = re.sub(r"\s+#\s.*$", "", match.group(2)).strip()
    base = Design()
    series = list(base.series)
    values: dict[str, object] = {}
    for slot, value in slots.items():
        name = slot[len("deck_"):]
        if name.startswith("series_") and name[7:].isdigit() and 1 <= int(name[7:]) <= 6:
            if not _HEX.fullmatch(value):
                raise DeckError(f"{slot}: se espera #RRGGBB, no {_ascii(value)!r}")
            series[int(name[7:]) - 1] = value.upper()
        elif name in ("background", "text", "muted", "negative"):
            if not _HEX.fullmatch(value):
                raise DeckError(f"{slot}: se espera #RRGGBB, no {_ascii(value)!r}")
            values[name] = value.upper()
        elif name in ("margin", "gap_max"):
            if not value.isdigit():
                raise DeckError(f"{slot}: se espera un entero en px, no {_ascii(value)!r}")
            values[name] = int(value)
        elif name in ("font_heading", "font_body"):
            if not re.fullmatch(r"[A-Za-z0-9 \-]+", value):
                raise DeckError(f"{slot}: nombre de fuente invalido {_ascii(value)!r}")
            values[name] = value
        elif name == "logo":
            logo = (root or path.parent.parent) / value
            mime = {".svg": "image/svg+xml", ".png": "image/png"}.get(logo.suffix.lower())
            if mime is None or not logo.is_file():
                raise DeckError(f"{slot}: se espera un .svg o .png existente, no {_ascii(value)!r}")
            values["logo"] = f"data:{mime};base64," + base64.b64encode(logo.read_bytes()).decode("ascii")
    return Design(series=tuple(series), **values)  # type: ignore[arg-type]


def weak_series(design: Design) -> set[int]:
    """Indices de colores de serie con poco contraste contra el fondo: las
    graficas etiquetan todos sus valores cuando usan uno de ellos."""
    return {i for i, color in enumerate(design.series)
            if contrast(color, design.background) < MIN_SERIES_CONTRAST}


def design_findings(design: Design) -> list[Finding]:
    text_ratio = contrast(design.text, design.background)
    close = [f"{i + 1}-{j + 1}" for i in range(len(design.series))
             for j in range(i + 1, len(design.series))
             if math.dist(_deutan(design.series[i]), _deutan(design.series[j])) < MIN_CVD_DISTANCE]
    return [
        Finding("P9 contraste del texto", text_ratio >= MIN_TEXT_CONTRAST,
                f"razon {text_ratio:.1f} contra el fondo (minimo {MIN_TEXT_CONTRAST})",
                warning=text_ratio < MIN_TEXT_CONTRAST),
        Finding("P10 series distinguibles con daltonismo", not close,
                "sin pares cercanos" if not close
                else "pares de series cercanos (deuteranopia): " + ", ".join(close),
                warning=bool(close)),
    ]


# ---------------------------------------------------------------------------
# Lectura del HTML generado: cifras atadas, literales y texto visible
# ---------------------------------------------------------------------------

_VOID = frozenset({"br", "hr", "img", "meta", "link", "input", "col", "wbr"})
_BLOCK = frozenset({"h1", "h2", "h3", "h4", "p", "li", "td", "th", "figcaption",
                    "div", "aside", "header", "footer", "text"})
_EXEMPT_CLASSES = frozenset({"auto", "source"})
CLAIM_TAGS = frozenset({"h1", "h2", "h3", "h4", "p", "li", "figcaption"})


@dataclass
class Block:
    """Un bloque de texto visible: partes ('text' | 'num' | 'lit', contenido)."""

    slide: int
    tag: str
    in_svg: bool
    parts: list[tuple[str, str]] = field(default_factory=list)

    @property
    def text(self) -> str:
        return re.sub(r"\s+", " ", "".join(part for _, part in self.parts)).strip()


@dataclass(frozen=True)
class BoundNumber:
    slide: int
    text: str
    src: str
    fmt: str


class DeckDocument(HTMLParser):
    """Recorre el deck una vez y deja: cifras atadas, literales y bloques."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.numbers: list[BoundNumber] = []
        self.literals: list[tuple[int, str, str]] = []   # (lamina, texto, porque)
        self.blocks: list[Block] = []
        self.gap_max: Optional[int] = None
        self._stack: list[tuple[str, str, Optional[Block], dict[str, str]]] = []
        self._slide = 0
        self._svg = 0
        self._capture: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        if tag in _VOID:
            return
        attr = {k: (v or "") for k, v in attrs}
        classes = set(attr.get("class", "").split())
        if tag == "html" and attr.get("data-gap-max", "").isdigit():
            self.gap_max = int(attr["data-gap-max"])
        if tag == "section" and "slide" in classes:
            self._slide = int(attr.get("data-n", "0") or 0)
        if tag == "svg":
            self._svg += 1
        if tag in ("style", "script", "title"):
            kind = "skip"
        elif "n" in classes and "data-src" in attr:
            kind, self._capture = "num", []
        elif "lit" in classes:
            kind, self._capture = "lit", []
        elif classes & _EXEMPT_CLASSES:
            kind = "exempt"
        else:
            kind = ""
        block = None
        if tag in _BLOCK and kind == "":
            block = Block(self._slide, tag, self._svg > 0)
            self.blocks.append(block)
        self._stack.append((tag, kind, block, attr))

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        return   # autocerradas (rect, line, br): sin texto

    def handle_endtag(self, tag: str) -> None:
        index = next((i for i in range(len(self._stack) - 1, -1, -1)
                      if self._stack[i][0] == tag), None)
        if index is None:
            return
        _, kind, _, attr = self._stack[index]
        del self._stack[index:]
        if tag == "svg":
            self._svg -= 1
        if kind in ("num", "lit"):
            text = "".join(self._capture)
            if kind == "num":
                self.numbers.append(BoundNumber(self._slide, text, attr["data-src"],
                                                attr.get("data-fmt", "")))
            else:
                self.literals.append((self._slide, text, attr.get("data-why", "")))
            owner = self._current_block()
            if owner is not None:
                owner.parts.append((kind, text))

    def _current_block(self) -> Optional[Block]:
        return next((b for _, _, b, _ in reversed(self._stack) if b is not None), None)

    def handle_data(self, data: str) -> None:
        kinds = [kind for _, kind, _, _ in self._stack]
        if "skip" in kinds or "exempt" in kinds:
            return
        if kinds and kinds[-1] in ("num", "lit"):
            self._capture.append(data)
            return
        owner = self._current_block()
        if owner is not None:
            owner.parts.append(("text", data))
        elif data.strip():
            loose = Block(self._slide, "", self._svg > 0, [("text", data)])
            self.blocks.append(loose)


def read_deck(document: str) -> DeckDocument:
    parser = DeckDocument()
    parser.feed(document)
    parser.close()
    return parser


# Numero suelto = palabra que, sin puntuacion ni simbolo de moneda, es un
# numero con sufijo de unidad opcional. 'N2', '20-F', '3nm' y '10-K' no lo
# son; '38%', '11x', '2,150' y 'US$5' si.
_WORD_SPLIT = re.compile(r"[\s/\u2013\u2014]+|(?<=\d)-(?=\d)")
_EDGE_PUNCT = "()[]{}.,;:!?\u00bf\u00a1\"'\u201c\u201d\u2018\u2019\u00ab\u00bb"
_CURRENCY = re.compile(r"^(?:[A-Za-z]{0,3}\$|\u20ac|\u00a3|\u00a5)")
_NUMERIC_WORD = re.compile(r"[-+\u2212]?\d[\d.,]*(?:%|x|pp|pb|bps?|k|m|mm|mn|bn)?", re.IGNORECASE)


def loose_numbers(text: str, years: set[int]) -> list[str]:
    """Palabras numericas de ``text`` que no son un anio del eje del modelo."""
    hits: list[str] = []
    for word in _WORD_SPLIT.split(text):
        core = _CURRENCY.sub("", word.strip(_EDGE_PUNCT)).strip(_EDGE_PUNCT)
        if not core or not _NUMERIC_WORD.fullmatch(core):
            continue
        if re.fullmatch(r"\d{4}", core) and int(core) in years:
            continue
        hits.append(word.strip(_EDGE_PUNCT))
    return hits


def unbound_numbers(doc: DeckDocument, years: set[int]) -> list[tuple[int, str, str]]:
    """(lamina, numero, contexto) de cada numero visible sin origen."""
    hits: list[tuple[int, str, str]] = []
    for block in doc.blocks:
        for kind, part in block.parts:
            if kind != "text":
                continue
            for number in loose_numbers(part, years):
                hits.append((block.slide, number, block.text[:60]))
    return hits


# ---------------------------------------------------------------------------
# El deck: anatomia de lamina y render
# ---------------------------------------------------------------------------

SLIDE_W, SLIDE_H = 1280, 720
DEFAULT_LEGAL = ("Documento con fines informativos. No constituye una recomendación de "
                 "inversión ni una oferta. Los supuestos y las conclusiones son del "
                 "analista que lo firma; las cifras proyectadas pueden no cumplirse.")
_AS_OF = re.compile(r"\d{4}-\d{2}-\d{2}")
_TICKER = re.compile(r"[A-Za-z0-9.\-]+")

_CSS = """
*{box-sizing:border-box}
html{background:#D9DCE1}
body{margin:0;font-family:%(body)s,Arial,sans-serif;color:%(text)s}
.slide{position:relative;width:1280px;height:720px;margin:24px auto;padding:%(m)spx;
 background:%(bg)s;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,.25)}
.slide h1,.slide h2{font-family:%(head)s,Georgia,serif;margin:0;color:%(text)s}
.slide h1{font-size:52px;line-height:1.1}
.slide h2{font-size:30px;line-height:1.2}
.lede{font-size:18px;line-height:1.4;margin:10px 0 0;color:%(muted)s}
.body{margin-top:18px}
.visual svg{display:block;width:100%%;height:auto}
.note{margin-top:14px;padding:12px 16px;border-left:4px solid %(s1)s;background:rgba(0,0,0,.035);
 font-size:15px;line-height:1.45}
.note b{display:block;font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:%(s1)s}
.note p{margin:4px 0 0}
.slide footer{position:absolute;left:%(m)spx;right:%(m)spx;bottom:%(fb)spx}
.takeaway{font-size:16px;font-weight:bold;margin:0 0 8px;padding-top:10px;border-top:1px solid %(muted)s}
.source{font-size:11px;color:%(muted)s;margin:0;padding-right:60px}
.page{position:absolute;right:0;bottom:0;font-size:11px;color:%(muted)s}
.cover,.act,.legal{display:flex;flex-direction:column;justify-content:center}
.cover .sub{font-size:22px;color:%(muted)s;margin:14px 0 0}
.cover .asof{font-size:14px;color:%(muted)s;margin:28px 0 0}
.cover .logo{position:absolute;right:%(m)spx;top:%(m)spx;max-height:56px}
.figs{display:flex;gap:40px;margin-top:40px}
.figs div{border-top:3px solid %(s1)s;padding-top:10px;min-width:180px}
.figs b{display:block;font-size:34px;font-family:%(head)s,Georgia,serif}
.figs span.lbl{font-size:13px;color:%(muted)s}
.act .num{font-size:96px;font-family:%(head)s,Georgia,serif;color:%(s1)s;line-height:1}
.act .q{font-size:20px;color:%(muted)s;margin:12px 0 0}
.legal p{font-size:14px;line-height:1.5;color:%(muted)s;max-width:900px}
svg text{font-family:%(body)s,Arial,sans-serif;font-size:13px;fill:%(text)s}
svg text.auto{fill:%(muted)s;font-size:12px}
svg text.val{font-weight:bold}
table.grid{border-collapse:collapse;width:100%%;font-size:14px}
table.grid th,table.grid td{padding:7px 10px;text-align:right;border-bottom:1px solid rgba(0,0,0,.12)}
table.grid th:first-child,table.grid td:first-child{text-align:left}
table.grid th{font-size:12px;color:%(muted)s}
table.grid td.hl{outline:2px solid %(text)s;font-weight:bold}
html.shot{background:%(bg)s}
html.shot .slide{margin:0;box-shadow:none}
@page{size:1280px 720px;margin:0}
@media print{html{background:none}.slide{margin:0;box-shadow:none;break-after:page}}
"""

# Mide cada lamina de contenido y deja el resultado en <html data-measure>.
# Con ?shot=N muestra solo la lamina N, pegada al origen (capturas).
_MEASURE_JS = """
(function(){
var slides=[].slice.call(document.querySelectorAll('section.slide'));
var shot=new URLSearchParams(location.search).get('shot');
if(shot!==null){
 document.documentElement.classList.add('shot');
 slides.forEach(function(s,i){if(String(i+1)!==shot){s.style.display='none';}});
 return;}
function clip(t){return (t||'').replace(/\\s+/g,' ').trim().slice(0,40);}
var out=slides.map(function(s){
 var kind=s.getAttribute('data-kind'),r={n:Number(s.getAttribute('data-n')),kind:kind};
 if(kind!=='content'&&kind!=='raw'){return r;}
 var body=s.querySelector('.body'),foot=s.querySelector('footer'),vis=s.querySelector('.visual');
 var sr=s.getBoundingClientRect(),pad=parseFloat(getComputedStyle(s).paddingLeft)||0;
 r.gap=foot.getBoundingClientRect().top-body.getBoundingClientRect().bottom;
 r.visual_h=vis?vis.getBoundingClientRect().height:null;
 r.overflow=[];r.svg_out=[];
 [].forEach.call(s.querySelectorAll('header *,.body *'),function(el){
  if(el.closest('svg')){return;}
  var b=el.getBoundingClientRect();
  if(b.width===0){return;}
  if(b.right>sr.right-pad+1||b.left<sr.left+pad-1||el.scrollWidth>el.clientWidth+1){
   r.overflow.push(el.tagName.toLowerCase()+': '+clip(el.textContent));}});
 [].forEach.call(s.querySelectorAll('svg'),function(svg){
  var b=svg.getBoundingClientRect();
  [].forEach.call(svg.querySelectorAll('text'),function(t){
   var q=t.getBoundingClientRect();
   if(q.left<b.left-1||q.right>b.right+1||q.top<b.top-1||q.bottom>b.bottom+1){
    r.svg_out.push(clip(t.textContent));}});});
 return r;});
document.documentElement.setAttribute('data-measure',JSON.stringify(out));
})();
"""


class Deck:
    """Un deck de cobertura. El script pide numeros aqui y arma las laminas."""

    def __init__(self, root: Path, model: str, ticker: str, as_of: str,
                 facts: Optional[str] = None, design: Optional[str] = None) -> None:
        if not _AS_OF.fullmatch(as_of):
            raise DeckError(f"as_of debe ser YYYY-MM-DD, no {_ascii(as_of)!r}")
        if not _TICKER.fullmatch(ticker):
            raise DeckError(f"ticker invalido: {_ascii(ticker)!r}")
        self.root = Path(root)
        self.ticker = ticker
        self.as_of = as_of
        self.book = ModelBook(self.root / model)
        self.facts = load_facts(self.root / facts, self.root) if facts else {}
        self.design = load_design(self.root / design if design else None, self.root)
        self._weak = weak_series(self.design)
        self._slides: list[tuple[str, str]] = []   # (tipo, html con marcadores)
        self._acts = 0

    # -- numeros ------------------------------------------------------------

    def series(self, sheet: str, label: str, section: Optional[str] = None) -> Series:
        return self.book.series(sheet, label, section)

    def cell(self, sheet: str, label: str, column: Optional[str] = None,
             section: Optional[str] = None) -> Num:
        return self.book.cell(sheet, label, column, section)

    def range(self, sheet: str, ref: str) -> list[list[Num]]:
        return self.book.range(sheet, ref)

    def fact(self, key: str) -> Num:
        if key not in self.facts:
            raise DeckError(f"hecho '{_ascii(key)}' no esta en el registro de hechos")
        return Num(self.facts[key].value, f"fact:{key}")

    # -- laminas ------------------------------------------------------------

    def cover(self, title: str, subtitle: str = "",
              figures: Sequence[tuple[str, str]] = ()) -> None:
        logo = f'<img class="logo" src="{self.design.logo}" alt=""/>' if self.design.logo else ""
        figs = "".join(f'<div><b>{_esc(value)}</b><span class="lbl">{_esc(label)}</span></div>'
                       for label, value in figures)
        self._slides.append(("cover", (
            f"{logo}<h1>{_esc(title)}</h1>"
            + (f'<p class="sub">{_esc(subtitle)}</p>' if subtitle else "")
            + (f'<div class="figs">{figs}</div>' if figs else "")
            + f'<p class="asof auto">{self.ticker} · {self.as_of}</p>')))

    def section(self, title: str, question: str = "") -> None:
        self._acts += 1
        self._slides.append(("act", (
            f'<div class="num auto">{self._acts:02d}</div><h1>{_esc(title)}</h1>'
            + (f'<p class="q">{_esc(question)}</p>' if question else ""))))

    def slide(self, title: str, lede: str, visual: str, note: str, source: str,
              takeaway: str = "") -> None:
        for name, value in (("title", title), ("visual", visual), ("note", note),
                            ("source", source)):
            if not value.strip():
                raise DeckError(f"slide '{_ascii(title)[:40]}': falta {name}")
        self._slides.append(("content", (
            f"<header><h2>{_esc(title)}</h2>"
            + (f'<p class="lede">{_esc(lede)}</p>' if lede else "") + "</header>"
            f'<div class="body"><div class="visual">{visual}</div>'
            f'<aside class="note"><b>Nota</b><p>{_esc(note)}</p></aside></div>'
            + self._footer(source, takeaway))))

    def raw(self, content: str, source: str, takeaway: str = "") -> None:
        """Lamina a medida: HTML/SVG propio. Pasa por bloqueo y medicion."""
        if not source.strip():
            raise DeckError("raw: falta source")
        self._slides.append(("raw", f'<div class="body">{content}</div>'
                             + self._footer(source, takeaway)))

    def legal(self, text: str = DEFAULT_LEGAL) -> None:
        self._slides.append(("legal", f"<p>{_esc(text)}</p>"))

    def _footer(self, source: str, takeaway: str) -> str:
        return ("<footer>"
                + (f'<p class="takeaway">{_esc(takeaway)}</p>' if takeaway else "")
                + f'<p class="source">{_esc(source)}</p></footer>')

    # -- salida -------------------------------------------------------------

    @property
    def output_path(self) -> Path:
        return (self.root / "research" / "notes"
                / f"{self.ticker.lower()}-deck_{self.as_of}.html")

    def render(self) -> str:
        if not self._slides:
            raise DeckError("el deck no tiene laminas")
        d = self.design
        css = _CSS % {"body": d.font_body, "head": d.font_heading, "text": d.text,
                      "muted": d.muted, "bg": d.background, "s1": d.series[0],
                      "m": d.margin, "fb": max(d.margin - 24, 16)}
        sections = []
        for number, (kind, content) in enumerate(self._slides, start=1):
            page = (f'<span class="page auto">{number}</span>'
                    if kind in ("content", "raw") else "")
            content = content.replace("</footer>", page + "</footer>")
            sections.append(f'<section class="slide {kind}" data-n="{number}" '
                            f'data-kind="{kind}">{content}</section>')
        document = (
            '<!doctype html>\n<html lang="es" '
            f'data-gap-max="{d.gap_max}" data-as-of="{self.as_of}">\n<head>\n'
            '<meta charset="utf-8"/>\n'
            f"<title>{html.escape(self.ticker)} · {self.as_of}</title>\n"
            f"<style>{css}</style>\n</head>\n<body>\n"
            + "\n".join(sections)
            + f"\n<script>{_MEASURE_JS}</script>\n</body>\n</html>\n")
        return render_markers(document)

    def build(self) -> Path:
        """Escribe el deck. Un numero sin origen lo impide (regla de bloqueo)."""
        document = self.render()
        loose = unbound_numbers(read_deck(document), self.book.axis_years())
        if loose:
            lines = [f"  lamina {slide}: {_ascii(number)!r} en '{_ascii(context)}'"
                     for slide, number, context in loose[:20]]
            more = f"\n  (+{len(loose) - 20} mas)" if len(loose) > 20 else ""
            raise DeckError(
                f"{len(loose)} numero(s) sin origen; usa deck.series/cell/fact o lit():\n"
                + "\n".join(lines) + more)
        out = self.output_path
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(document, encoding="utf-8", newline="\n")
        return out


    # -- graficas (SVG sin librerias) -----------------------------------------
    # Reciben Num; la geometria usa float(). Las marcas de eje las calcula el
    # tool y van con class="auto" (no son cifras del analista).

    def _color(self, index: int) -> str:
        return self.design.series[index % len(self.design.series)]

    def bars(self, series: Series, keys: Optional[Sequence[PeriodKey]] = None,
             show: Sequence[PeriodKey] = (), fmt: str = ",.0f", scale: float = 1.0,
             tick_fmt: Optional[str] = None, w: int = 1100, h: int = 320) -> str:
        """Barras por periodo. Observado y proyectado en colores distintos, con
        corte rotulado; negativos en el color de negativo (sirve para
        crecimientos). ``show``: periodos que llevan etiqueta de valor."""
        keys = list(keys or series.keys)
        values = [series[k] / scale if scale != 1.0 else series[k] for k in keys]
        ticks = _nice_ticks(min(0.0, *(float(v) for v in values)),
                            max(0.0, *(float(v) for v in values)))
        lo, hi = ticks[0], ticks[-1]
        left, right, top, bottom = 64, 16, 26, 30
        plot_w, plot_h = w - left - right, h - top - bottom
        slot = plot_w / len(keys)

        def y(value: float) -> float:
            return top + plot_h * (hi - value) / (hi - lo)

        label_all = bool(self._weak & {0, 1})
        out = [_svg_open(w, h, series.label)]
        out += _y_axis(ticks, y, left, w - right, tick_fmt or fmt)
        split = next((i for i, k in enumerate(keys) if not series.is_actual(k)), None)
        if split not in (None, 0):
            x = left + slot * split
            out.append(f'<line x1="{x:.1f}" y1="{top}" x2="{x:.1f}" y2="{top + plot_h}" '
                       f'stroke="{self.design.muted}" stroke-dasharray="4 4"/>')
            out.append(f'<text class="auto" x="{x + 6:.1f}" y="{top + 12}">Proyectado</text>')
        for i, (key, value) in enumerate(zip(keys, values)):
            v = float(value)
            color = (self.design.negative if v < 0
                     else self._color(0) if series.is_actual(key) else self._color(1))
            x, y0, y1 = left + slot * i + slot * 0.15, y(max(v, 0.0)), y(min(v, 0.0))
            out.append(f'<rect x="{x:.1f}" y="{y0:.1f}" width="{slot * 0.7:.1f}" '
                       f'height="{max(y1 - y0, 0.5):.1f}" fill="{color}"/>')
            out.append(f'<text class="auto" x="{x + slot * 0.35:.1f}" y="{h - 10}" '
                       f'text-anchor="middle">{html.escape(str(key))}</text>')
            if label_all or key in show:
                ly = y0 - 5 if v >= 0 else y1 + 14
                out.append(f'<text class="val" x="{x + slot * 0.35:.1f}" y="{ly:.1f}" '
                           f'text-anchor="middle">{value:{fmt}}</text>')
        out.append("</svg>")
        return "".join(out)

    def lines(self, named: Sequence[tuple[str, Series]],
              keys: Optional[Sequence[PeriodKey]] = None, show: Sequence[PeriodKey] = (),
              fmt: str = ".0%", tick_fmt: Optional[str] = None,
              w: int = 1100, h: int = 320) -> str:
        """Lineas de una o varias series sobre el mismo eje; la zona
        proyectada va sombreada y cada linea lleva su nombre al final."""
        first = named[0][1]
        keys = list(keys or first.keys)
        flat = [float(s[k]) for _, s in named for k in keys]
        ticks = _nice_ticks(min(flat), max(flat))
        lo, hi = ticks[0], ticks[-1]
        left, right, top, bottom = 64, 170, 26, 30
        plot_w, plot_h = w - left - right, h - top - bottom
        step = plot_w / max(len(keys) - 1, 1)

        def x(i: int) -> float:
            return left + step * i

        def y(value: float) -> float:
            return top + plot_h * (hi - value) / (hi - lo)

        out = [_svg_open(w, h, ", ".join(name for name, _ in named))]
        split = next((i for i, k in enumerate(keys) if not first.is_actual(k)), None)
        if split not in (None, 0):
            x0 = x(split) - step / 2
            out.append(f'<rect x="{x0:.1f}" y="{top}" width="{left + plot_w - x0:.1f}" '
                       f'height="{plot_h}" fill="{self.design.muted}" opacity="0.10"/>')
            out.append(f'<text class="auto" x="{x0 + 6:.1f}" y="{top + 12}">Proyectado</text>')
        out += _y_axis(ticks, y, left, left + plot_w, tick_fmt or fmt)
        for i, key in enumerate(keys):
            if i % max(len(keys) // 8, 1) == 0 or i == len(keys) - 1:
                out.append(f'<text class="auto" x="{x(i):.1f}" y="{h - 10}" '
                           f'text-anchor="middle">{html.escape(str(key))}</text>')
        for index, (name, series) in enumerate(named):
            color = self._color(index)
            points = " ".join(f"{x(i):.1f},{y(float(series[k])):.1f}" for i, k in enumerate(keys))
            out.append(f'<polyline points="{points}" fill="none" stroke="{color}" '
                       f'stroke-width="2.5"/>')
            out.append(f'<text x="{left + plot_w + 8:.1f}" '
                       f'y="{y(float(series[keys[-1]])) + 4:.1f}" fill="{color}">'
                       f"{_esc(name)}</text>")
            for i, key in enumerate(keys):
                if index in self._weak or key in show:
                    out.append(f'<text class="val" x="{x(i):.1f}" '
                               f'y="{y(float(series[key])) - 8:.1f}" text-anchor="middle">'
                               f"{series[key]:{fmt}}</text>")
        out.append("</svg>")
        return "".join(out)

    def football(self, rows: Sequence[tuple[str, Num, Num, Num]],
                 marker: Optional[tuple[str, Num]] = None, fmt: str = ",.0f",
                 w: int = 1100, h: Optional[int] = None) -> str:
        """Football field: por fila (metodo, minimo, maximo, punto central);
        ``marker`` = (rotulo, valor) de la linea vertical (precio actual)."""
        h = h or 60 * len(rows) + 60
        values = [float(v) for _, low, high, mid in rows for v in (low, high, mid)]
        if marker is not None:
            values.append(float(marker[1]))
        ticks = _nice_ticks(min(values), max(values))
        lo, hi = ticks[0], ticks[-1]
        left, right, top = 300, 70, 34
        plot_w = w - left - right

        def x(value: float) -> float:
            return left + plot_w * (value - lo) / (hi - lo)

        out = [_svg_open(w, h, "Rangos de valor por metodo")]
        for tick in ticks:
            out.append(f'<text class="auto" x="{x(tick):.1f}" y="{h - 8}" '
                       f'text-anchor="middle">{tick:{fmt}}</text>')
        for i, (label, low, high, mid) in enumerate(rows):
            cy, color = top + 60 * i + 22, self._color(i)
            out.append(f'<text x="0" y="{cy + 4}">{_esc(label)}</text>')
            out.append(f'<rect x="{x(float(low)):.1f}" y="{cy - 12}" '
                       f'width="{x(float(high)) - x(float(low)):.1f}" height="24" '
                       f'fill="{color}" opacity="0.30"/>')
            out.append(f'<circle cx="{x(float(mid)):.1f}" cy="{cy}" r="6" fill="{color}"/>')
            out.append(f'<text class="val" x="{x(float(mid)):.1f}" y="{cy - 18}" '
                       f'text-anchor="middle">{mid:{fmt}}</text>')
            out.append(f'<text x="{x(float(low)) - 6:.1f}" y="{cy + 4}" '
                       f'text-anchor="end">{low:{fmt}}</text>')
            out.append(f'<text x="{x(float(high)) + 6:.1f}" y="{cy + 4}">{high:{fmt}}</text>')
        if marker is not None:
            mx = x(float(marker[1]))
            out.append(f'<line x1="{mx:.1f}" y1="{top - 6}" x2="{mx:.1f}" y2="{h - 26}" '
                       f'stroke="{self.design.text}" stroke-width="2"/>')
            out.append(f'<text class="val" x="{mx:.1f}" y="{top - 14}" text-anchor="middle">'
                       f"{_esc(marker[0])} {marker[1]:{fmt}}</text>")
        out.append("</svg>")
        return "".join(out)

    def waterfall(self, steps: Sequence[tuple[str, Num]], total: tuple[str, Num],
                  fmt: str = ",.0f", w: int = 1100, h: int = 320) -> str:
        """Puente: cada paso suma o resta sobre el acumulado; ``total`` es la
        barra final, leida del modelo (no se calcula aqui)."""
        running, spans = 0.0, []
        for label, value in steps:
            start, running = running, running + float(value)
            spans.append((label, value, start, running))
        spans.append((total[0], total[1], 0.0, float(total[1])))
        ticks = _nice_ticks(min(0.0, *(min(a, b) for _, _, a, b in spans)),
                            max(0.0, *(max(a, b) for _, _, a, b in spans)))
        lo, hi = ticks[0], ticks[-1]
        left, right, top, bottom = 64, 16, 26, 44
        plot_w, plot_h = w - left - right, h - top - bottom
        slot = plot_w / len(spans)

        def y(value: float) -> float:
            return top + plot_h * (hi - value) / (hi - lo)

        out = [_svg_open(w, h, "Puente de valor")]
        out += _y_axis(ticks, y, left, w - right, fmt)
        for i, (label, value, start, end) in enumerate(spans):
            is_total = i == len(spans) - 1
            color = (self._color(0) if is_total
                     else self.design.negative if end < start else self._color(1))
            x0, y0, y1 = left + slot * i + slot * 0.15, y(max(start, end)), y(min(start, end))
            out.append(f'<rect x="{x0:.1f}" y="{y0:.1f}" width="{slot * 0.7:.1f}" '
                       f'height="{max(y1 - y0, 0.5):.1f}" fill="{color}"/>')
            out.append(f'<text class="val" x="{x0 + slot * 0.35:.1f}" y="{y0 - 5:.1f}" '
                       f'text-anchor="middle">{value:{fmt}}</text>')
            out.append(f'<text x="{x0 + slot * 0.35:.1f}" y="{h - 14}" text-anchor="middle">'
                       f"{_esc(label)}</text>")
        out.append("</svg>")
        return "".join(out)

    def heatmap(self, grid: Sequence[Sequence[Num]], rows: Sequence[Union[Num, str]],
                cols: Sequence[Union[Num, str]], fmt: str = ",.0f", row_fmt: str = ".1%",
                col_fmt: str = ".1%", corner: str = "",
                center: Optional[tuple[int, int]] = None) -> str:
        """Cuadricula de sensibilidad como tabla; el fondo crece con el valor
        y ``center`` (fila, columna) marca el caso base."""
        flat = [float(v) for line in grid for v in line]
        lo, span = min(flat), (max(flat) - min(flat)) or 1.0

        def head(item: Union[Num, str], spec: str) -> str:
            return f"{item:{spec}}" if isinstance(item, Num) else _esc(item)

        r, g, b = _rgb(self._color(0))
        out = ['<table class="grid heat"><tr>', f"<th>{_esc(corner)}</th>"]
        out += [f"<th>{head(c, col_fmt)}</th>" for c in cols]
        out.append("</tr>")
        for i, line in enumerate(grid):
            out.append(f"<tr><th>{head(rows[i], row_fmt)}</th>")
            for j, value in enumerate(line):
                alpha = 0.06 + 0.34 * (float(value) - lo) / span
                hl = ' class="hl"' if center == (i, j) else ""
                out.append(f'<td{hl} style="background:rgba({r},{g},{b},{alpha:.2f})">'
                           f"{value:{fmt}}</td>")
            out.append("</tr>")
        out.append("</table>")
        return "".join(out)

    def table(self, header: Sequence[str], rows: Sequence[Sequence[str]]) -> str:
        """Tabla simple. Las celdas son texto; los numeros llegan ya
        formateados desde un Num (f-string)."""
        out = ['<table class="grid"><tr>']
        out += [f"<th>{_esc(cell)}</th>" for cell in header]
        out.append("</tr>")
        for row in rows:
            out.append("<tr>" + "".join(f"<td>{_esc(cell)}</td>" for cell in row) + "</tr>")
        out.append("</table>")
        return "".join(out)


def _svg_open(w: int, h: int, label: str) -> str:
    return (f'<svg viewBox="0 0 {w} {h}" role="img" '
            f'aria-label="{html.escape(_MARKER.sub("", label), quote=True)}">')


def _nice_ticks(lo: float, hi: float, count: int = 4) -> list[float]:
    """Marcas de eje redondas que cubren [lo, hi]."""
    if hi <= lo:
        hi = lo + 1.0
    raw = (hi - lo) / count
    magnitude = 10 ** math.floor(math.log10(raw))
    step = next(m * magnitude for m in (1, 2, 2.5, 5, 10) if m * magnitude >= raw * (1 - 1e-9))
    tick = math.floor(lo / step + 1e-9) * step
    ticks = [round(tick, 10)]
    while ticks[-1] < hi - step * 1e-9:
        ticks.append(round(ticks[-1] + step, 10))
    return ticks


def _y_axis(ticks: Sequence[float], y: Callable[[float], float], x0: float, x1: float,
            fmt: str) -> list[str]:
    out = []
    for tick in ticks:
        out.append(f'<line x1="{x0:.1f}" y1="{y(tick):.1f}" x2="{x1:.1f}" y2="{y(tick):.1f}" '
                   f'stroke="currentColor" opacity="0.12"/>')
        out.append(f'<text class="auto" x="{x0 - 8:.1f}" y="{y(tick) + 4:.1f}" '
                   f'text-anchor="end">{tick:{fmt}}</text>')
    return out


# ---------------------------------------------------------------------------
# Auditoria: cada cifra del deck contra su origen
# ---------------------------------------------------------------------------

_SRC_TOKEN = re.compile(
    r"(?:'(?P<quoted>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_.]*))!(?P<coord>[A-Z]{1,3}\d+)"
    r"|fact:(?P<fact>[A-Za-z0-9_.\-]+)")
_BINARY = {ast.Add: lambda a, b: a + b, ast.Sub: lambda a, b: a - b,
           ast.Mult: lambda a, b: a * b, ast.Div: lambda a, b: a / b,
           ast.Pow: lambda a, b: a ** b}


def evaluate(src: str, book: ModelBook, facts: dict[str, Fact]) -> float:
    """Re-evalua la expresion de origen de una cifra. Solo admite referencias
    a celdas, ``fact:clave``, constantes, + - * / ** y abs()."""
    names: dict[str, float] = {}

    def bind(match: re.Match[str]) -> str:
        if match.group("fact") is not None:
            key = match.group("fact")
            if key not in facts:
                raise DeckError(f"hecho '{_ascii(key)}' no esta en el registro")
            value = facts[key].value
        else:
            value = book.num(match.group("quoted") or match.group("sheet"),
                             match.group("coord")).value
        names[f"_v{len(names)}"] = value
        return f"_v{len(names) - 1}"

    try:
        tree = ast.parse(_SRC_TOKEN.sub(bind, src), mode="eval")
    except SyntaxError as exc:
        raise DeckError(f"origen ilegible: {_ascii(src)!r}") from exc

    def walk(node: ast.AST) -> float:
        if isinstance(node, ast.Expression):
            return walk(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) \
                and not isinstance(node.value, bool):
            return float(node.value)
        if isinstance(node, ast.Name) and node.id in names:
            return names[node.id]
        if isinstance(node, ast.BinOp) and type(node.op) in _BINARY:
            return _BINARY[type(node.op)](walk(node.left), walk(node.right))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            value = walk(node.operand)
            return -value if isinstance(node.op, ast.USub) else value
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) \
                and node.func.id == "abs" and len(node.args) == 1 and not node.keywords:
            return abs(walk(node.args[0]))
        raise DeckError(f"origen con una operacion no admitida: {_ascii(src)!r}")

    try:
        return walk(tree)
    except ZeroDivisionError as exc:
        raise DeckError(f"origen con division entre cero: {_ascii(src)!r}") from exc


def audit_deck(deck_html: Path, model: Path, facts_path: Optional[Path] = None,
               root: Optional[Path] = None) -> tuple[list[Finding], DeckDocument, dict[str, Fact]]:
    """P1 a P4: cifras atadas, sin numeros sueltos, literales y hechos manuales."""
    doc = read_deck(deck_html.read_text(encoding="utf-8"))
    book = ModelBook(model)
    root = root or model.resolve().parent.parent
    facts = load_facts(facts_path, root) if facts_path else {}
    stale: list[str] = []
    for number in doc.numbers:
        try:
            expected = format(evaluate(number.src, book, facts), number.fmt)
        except DeckError as exc:
            stale.append(f"lamina {number.slide}: {exc}")
            continue
        if expected != number.text:
            stale.append(f"lamina {number.slide}: muestra {_ascii(number.text)!r}, "
                         f"el origen da {_ascii(expected)!r} ({_ascii(number.src)})")
    loose = unbound_numbers(doc, book.axis_years())
    used = {m.group("fact") for n in doc.numbers for m in _SRC_TOKEN.finditer(n.src)
            if m.group("fact")}
    manual = sorted(key for key in used if key in facts and facts[key].manual)
    findings = [
        Finding("P1 cifras iguales a su origen", not stale,
                f"{len(doc.numbers)} cifras verificadas" if not stale
                else f"{len(stale)} de {len(doc.numbers)} no coinciden (deck desactualizado "
                     f"o alterado): " + M._examples(stale)),
        Finding("P2 sin numeros sueltos", not loose,
                "todo numero visible tiene origen" if not loose
                else f"{len(loose)} sin origen: " + M._examples(
                    [f"lamina {s}: {_ascii(n)!r}" for s, n, _ in loose])),
        Finding("P3 literales declarados con lit()", not doc.literals,
                "ninguno" if not doc.literals
                else f"{len(doc.literals)} a revisar en el gate: " + M._examples(
                    [f"lamina {s}: {_ascii(t)!r} ({_ascii(w)})" for s, t, w in doc.literals]),
                warning=bool(doc.literals)),
        Finding("P4 hechos con verificacion manual", not manual,
                "ninguno" if not manual
                else "fuente no textual, cotejar a mano: " + ", ".join(_ascii(k) for k in manual),
                warning=bool(manual)),
    ]
    return findings, doc, facts


def claims(doc: DeckDocument) -> list[tuple[int, str]]:
    """(lamina, frase) de todo el texto corrido del deck, en orden."""
    out: list[tuple[int, str]] = []
    for block in doc.blocks:
        if block.in_svg or block.tag not in CLAIM_TAGS or not block.text:
            continue
        for sentence in re.split(r"(?<=[.!?])\s+", block.text):
            if sentence.strip():
                out.append((block.slide, sentence.strip()))
    return out


def _md_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def claims_table(doc: DeckDocument) -> str:
    lines = ["| Lámina | Frase | Evidencia |", "|---|---|---|"]
    lines += [f"| {slide} | {_md_cell(sentence)} |  |" for slide, sentence in claims(doc)]
    return "\n".join(lines) + "\n"


def write_report(path: Path, deck_html: Path, groups: Sequence[tuple[str, Sequence[Finding]]],
                 doc: DeckDocument, facts: dict[str, Fact]) -> None:
    """Reporte de verificacion del gate 2 (partes deterministas). La columna
    Evidencia de las afirmaciones la llena quien verifica."""
    out = [f"# Verificación del deck — {deck_html.name}", ""]
    for title, findings in groups:
        out += [f"## {title}", "", "| Estado | Check | Detalle |", "|---|---|---|"]
        out += [f"| {f.mark} | {f.check} | {_md_cell(f.detail)} |" for f in findings]
        out.append("")
    out += ["## Cifras y su origen", "", "| Lámina | Texto | Origen | Formato |", "|---|---|---|---|"]
    out += [f"| {n.slide} | {_md_cell(n.text)} | `{_md_cell(n.src)}` | `{n.fmt}` |"
            for n in doc.numbers]
    out += ["", "## Literales declarados", "", "| Lámina | Texto | Porqué |", "|---|---|---|"]
    out += [f"| {s} | {_md_cell(t)} | {_md_cell(w)} |" for s, t, w in doc.literals]
    out += ["", "## Hechos citados", "",
            "| Clave | Valor | Etiqueta | Documento | Localizador | Verificación |",
            "|---|---|---|---|---|---|"]
    out += [f"| {f.key} | {f.value:g} {f.unit} | {f.tag} | {f.doc} | {_md_cell(f.locator)} | "
            f"{'manual' if f.manual else 'cita encontrada'} |" for f in facts.values()]
    out += ["", "## Afirmaciones", "", claims_table(doc)]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(out), encoding="utf-8", newline="\n")


# ---------------------------------------------------------------------------
# Medicion de espacio con el navegador ya instalado
# ---------------------------------------------------------------------------

MEASURE_TIMEOUT_SECONDS = 120
MEASURED_KINDS = ("content", "raw")
_BROWSER_NAMES = ("msedge", "chrome", "google-chrome", "google-chrome-stable",
                  "chromium", "chromium-browser")
_BROWSER_CANDIDATES = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
)
_MEASURE_ATTR = re.compile(r'data-measure="([^"]*)"')


class MeasureError(Exception):
    """El navegador no pudo abrir o medir el deck."""


def find_browser() -> Optional[str]:
    """Edge, Chrome o Chromium ya instalado. ``RA_BROWSER`` fija la ruta."""
    forced = os.environ.get("RA_BROWSER")
    if forced:
        return forced if Path(forced).is_file() else None
    for name in _BROWSER_NAMES:
        found = shutil.which(name)
        if found:
            return found
    return next((c for c in _BROWSER_CANDIDATES if Path(c).is_file()), None)


def _run_browser(browser: str, args: Sequence[str]) -> str:
    with tempfile.TemporaryDirectory() as profile:
        command = [browser, "--headless=new", "--disable-gpu", "--no-first-run",
                   f"--user-data-dir={profile}", f"--window-size={SLIDE_W},{SLIDE_H}", *args]
        try:
            done = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                                  errors="replace", timeout=MEASURE_TIMEOUT_SECONDS, check=False)
        except subprocess.TimeoutExpired as exc:
            raise MeasureError(f"el navegador no respondio en {MEASURE_TIMEOUT_SECONDS}s") from exc
        except OSError as exc:
            raise MeasureError(f"no se pudo ejecutar el navegador ({_ascii(exc)})") from exc
    return done.stdout or ""


def dump_measure(browser: str, deck_html: Path) -> list[dict[str, object]]:
    """Abre el deck sin ventana y devuelve la medicion que deja su script."""
    dom = _run_browser(browser, ["--dump-dom", deck_html.resolve().as_uri()])
    match = _MEASURE_ATTR.search(dom)
    if match is None:
        raise MeasureError("el deck no dejo medicion (no lo genero deck_builder, o el "
                           "navegador no ejecuto su script)")
    return json.loads(html.unescape(match.group(1)))


def measure_findings(data: Sequence[dict[str, object]], gap_max: int) -> list[Finding]:
    """P5 a P8 a partir de la medicion por lamina."""
    measured = [d for d in data if d.get("kind") in MEASURED_KINDS]
    overlap = [f"lamina {d['n']}: {float(d['gap']):.0f} px" for d in measured
               if float(d["gap"]) <= 0]
    empty = []
    for d in measured:
        gap = float(d["gap"])
        if gap >= gap_max:
            visual = (f" (el visual mide {float(d['visual_h']):.0f} px)"
                      if d.get("visual_h") is not None else "")
            empty.append(f"lamina {d['n']}: {gap:.0f} px libres{visual}")
    wide = [f"lamina {d['n']}: {_ascii(item)}" for d in measured for item in d.get("overflow", [])]
    outside = [f"lamina {d['n']}: {_ascii(item)!r}" for d in measured
               for item in d.get("svg_out", [])]
    total = f"{len(measured)} laminas medidas"
    return [
        Finding("P5 sin contenido encimado", not overlap,
                total if not overlap else "contenido sobre el cierre: " + M._examples(overlap)),
        Finding("P6 sin lamina vacia", not empty,
                f"{total}, umbral {gap_max} px" if not empty
                else f"mas de {gap_max} px libres: " + M._examples(empty)),
        Finding("P7 sin desborde horizontal", not wide,
                total if not wide else "texto mas ancho que su caja: " + M._examples(wide)),
        Finding("P8 etiquetas dentro de su grafica", not outside,
                total if not outside else "etiqueta fuera del area: " + M._examples(outside)),
    ]


def measure_deck(deck_html: Path, shots: Optional[Path] = None) -> list[Finding]:
    """Mide el deck. Sin navegador: [pendiente], nunca verde."""
    doc = read_deck(deck_html.read_text(encoding="utf-8"))
    gap_max = doc.gap_max if doc.gap_max is not None else Design().gap_max
    browser = find_browser()
    if browser is None:
        return [Finding("P5-P8 medicion de espacio", False,
                        "sin Edge, Chrome ni Chromium en esta maquina: el deck NO se midio. "
                        "Instala uno o fija RA_BROWSER y vuelve a correr measure", pending=True)]
    try:
        data = dump_measure(browser, deck_html)
        findings = measure_findings(data, gap_max)
        if shots is not None:
            shots.mkdir(parents=True, exist_ok=True)
            for item in data:
                target = (shots / f"slide-{int(item['n']):02d}.png").resolve()  # type: ignore[call-overload]
                _run_browser(browser, [f"--screenshot={target}",
                                       f"{deck_html.resolve().as_uri()}?shot={item['n']}"])
                if not target.is_file():
                    raise MeasureError(f"no se genero la captura de la lamina {item['n']}")
    except MeasureError as exc:
        return [Finding("P5-P8 medicion de espacio", False, str(exc))]
    return findings


# ---------------------------------------------------------------------------
# Demo y CLI
# ---------------------------------------------------------------------------

DEMO_AS_OF = "2026-01-01"
_DEMO_ROWS = (
    ("Ingresos", (820.0, 905.0, 1010.0, 1140.0, 1265.0, 1380.0)),
    ("Utilidad de operacion", (148.0, 172.0, 202.0, 239.0, 272.0, 304.0)),
    ("Flujo libre", (96.0, 110.0, 131.0, 150.0, 176.0, 201.0)),
)


def _demo_model(path: Path) -> None:
    """Libro sintetico de valores (sin formulas: no hace falta recalcular)."""
    styler = B.ModelStyler()
    styler.set_periodicity("annual")
    ws = styler.new_sheet("Model")
    styler.brand_bar(ws, "Modelo demo")
    styler.label_col_width(ws)
    styler.period_header(ws, 4, 3, B.PeriodHeader(2022, 2027, 2024))
    for offset, (label, values) in enumerate(_DEMO_ROWS):
        row = 6 + offset
        styler.set_cell(ws, f"B{row}", label, B.CellRole.LABEL, numfmt=B.NumFmt.GENERAL)
        for i, value in enumerate(values):
            styler.set_cell(ws, f"{get_column_letter(3 + i)}{row}", value, B.CellRole.OBSERVED)
    val = styler.new_sheet("Valuation", freeze=None)
    styler.brand_bar(val, "Valuacion demo")
    styler.label_col_width(val)
    for row, (label, value, numfmt) in enumerate((
            ("WACC", 0.09, B.NumFmt.PCT1), ("Valor por accion", 48.0, B.NumFmt.DEC2),
            ("Rango bajo", 41.0, B.NumFmt.DEC2), ("Rango alto", 56.0, B.NumFmt.DEC2),
            ("Precio actual", 44.0, B.NumFmt.DEC2)), start=5):
        styler.set_cell(val, f"B{row}", label, B.CellRole.LABEL, numfmt=B.NumFmt.GENERAL)
        styler.set_cell(val, f"C{row}", value, B.CellRole.OBSERVED, numfmt=numfmt)
    path.parent.mkdir(parents=True, exist_ok=True)
    styler.save(str(path))


def build_demo(folder: Path) -> tuple[Path, Path, Path]:
    """Cobertura sintetica minima y su deck. Devuelve (deck, modelo, hechos)."""
    root = folder / "demo"
    model = root / "model" / f"demo-model_{DEMO_AS_OF}_v1.xlsx"
    _demo_model(model)
    report = root / "research" / "industry" / f"industry-report_{DEMO_AS_OF}.md"
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text("# Industria demo\n\nParticipacion de la emisora: 31.5% del mercado.\n",
                      encoding="utf-8", newline="\n")
    facts = root / "research" / "notes" / f"demo-facts_{DEMO_AS_OF}.json"
    facts.parent.mkdir(parents=True, exist_ok=True)
    facts.write_text(json.dumps({"as_of": DEMO_AS_OF, "facts": [{
        "key": "share", "value": 31.5, "unit": "%", "tag": "observado",
        "doc": f"research/industry/industry-report_{DEMO_AS_OF}.md",
        "locator": "parrafo unico", "quote": "31.5%"}]}, indent=1), encoding="utf-8", newline="\n")

    deck = Deck(root=root, model=str(model.relative_to(root)), ticker="DEMO", as_of=DEMO_AS_OF,
                facts=str(facts.relative_to(root)))
    rev = deck.series("Model", "Ingresos")
    ebit = deck.series("Model", "Utilidad de operacion")
    fcf = deck.series("Model", "Flujo libre")
    margin, growth = ebit.over(rev), rev.growth()
    wacc = deck.cell("Valuation", "WACC")
    value = deck.cell("Valuation", "Valor por accion")
    price = deck.cell("Valuation", "Precio actual")
    src = "Fuente: modelo demo y reporte de industria demo."
    deck.cover("Emisora demo", "Deck de ejemplo generado por deck_builder",
               figures=[("Ventas 2024", f"{rev[2024]:,.0f}"),
                        ("Margen de operación", f"{margin[2024]:.1%}"),
                        ("Participación de mercado", f"{deck.fact('share'):.1f}%")])
    deck.section("Crecimiento", "¿De dónde sale el crecimiento?")
    deck.slide(
        title=f"Las ventas crecen {rev.cagr(2024, 2027):.1%} al año hasta 2027",
        lede="El ritmo proyectado es menor al observado.",
        visual=deck.bars(rev, show=(2022, 2024, 2027), h=330),
        note=(f"Entre 2022 y 2024 las ventas crecieron {rev.cagr(2022, 2024):.1%} anual. "
              f"El modelo supone {growth[2027]:.1%} en 2027."),
        source=src, takeaway="El crecimiento se modera, no se detiene.")
    deck.slide(
        title=f"El margen de operación llega a {margin[2027]:.1%} en 2027",
        lede="La escala explica la mejora.",
        visual=deck.lines([("Margen de operación", margin)], show=(2022, 2027),
                          fmt=".1%", h=350),
        note=(f"El margen pasó de {margin[2022]:.1%} a {margin[2024]:.1%}. "
              f"El flujo libre proyectado para 2027 es {fcf[2027]:,.0f}."),
        source=src)
    deck.section("Valuación")
    deck.slide(
        title=f"El valor estimado es {value:,.0f} por acción contra un precio de {price:,.0f}",
        lede=f"Descontado al {wacc:.1%}.",
        visual=deck.football(
            [("Flujos descontados", deck.cell("Valuation", "Rango bajo"),
              deck.cell("Valuation", "Rango alto"), value)],
            marker=("Precio", price), h=200)
        + deck.table(["Concepto", "Valor"],
                     [["Tasa de descuento", f"{wacc:.1%}"],
                      ["Valor por acción", f"{value:,.2f}"],
                      ["Precio actual", f"{price:,.2f}"],
                      ["Diferencia", f"{value / price - 1:+.1%}"]]),
        note=(f"El rango refleja la sensibilidad a la tasa de descuento. El punto central "
              f"queda {value / price - 1:.0%} arriba del precio."),
        source=src, takeaway="El precio está dentro del rango, debajo del punto central.")
    deck.legal()
    return deck.build(), model, facts


def _report(groups: Sequence[tuple[str, Sequence[Finding]]]) -> int:
    """Imprime por familia. Exit: 1 fallas; 3 solo pendientes; 0 verde."""
    for title, findings in groups:
        print(f"-- {title}")
        for finding in findings:
            print(f"{finding.mark} {finding.check}: {_ascii(finding.detail)}")
    everything = [f for _, findings in groups for f in findings]
    failures = sum(1 for f in everything if f.blocking)
    pending = sum(1 for f in everything if f.pending)
    warnings = sum(1 for f in everything if f.warning and not f.pending)
    passed = sum(1 for f in everything if f.ok and not f.pending)
    print(f"Resumen: {passed} ok, {failures} fallas, {pending} pendientes, {warnings} avisos")
    if failures:
        print("Veredicto: FALLA")
        return 1
    if pending:
        print("Veredicto: PENDIENTE - no es verde; resuelve los pendientes y vuelve a correr")
        return 3
    print("Veredicto: VERDE")
    return 0


def main(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="deck_builder.py",
        description="Deck de cobertura por codigo: demo, auditoria de cifras, medicion "
                    "de espacio y lista de afirmaciones.")
    sub = parser.add_subparsers(dest="command", required=True)
    demo = sub.add_parser("demo", help="genera un deck sintetico y lo audita")
    demo.add_argument("folder")
    audit = sub.add_parser("audit", help="cada cifra del deck contra el modelo")
    audit.add_argument("deck")
    audit.add_argument("model")
    audit.add_argument("--facts", help="registro de hechos (JSON)")
    audit.add_argument("--root", help="carpeta de la cobertura (default: dos niveles "
                                      "arriba del modelo)")
    audit.add_argument("--design", help="brand/DESIGN.md: agrega los avisos de contraste")
    audit.add_argument("--report", help="escribe el reporte de verificacion (Markdown)")
    measure = sub.add_parser("measure", help="mide el espacio de cada lamina")
    measure.add_argument("deck")
    measure.add_argument("--shots", help="carpeta para una captura por lamina")
    claim = sub.add_parser("claims", help="lista las frases del deck en un Markdown")
    claim.add_argument("deck")
    claim.add_argument("out")
    try:
        args = parser.parse_args(list(argv[1:]))
    except SystemExit as exc:
        return 2 if exc.code not in (0, None) else 0
    try:
        if args.command == "demo":
            deck, model, facts = build_demo(Path(args.folder))
            print(f"[ok] deck demo -> {_ascii(deck)}")
            findings, _, _ = audit_deck(deck, model, facts)
            return _report([("Cifras", findings)])
        if args.command == "audit":
            deck = Path(args.deck)
            findings, doc, facts = audit_deck(
                deck, Path(args.model), Path(args.facts) if args.facts else None,
                Path(args.root) if args.root else None)
            groups = [("Cifras", findings)]
            if args.design:
                groups.append(("Diseno", design_findings(load_design(Path(args.design)))))
            if args.report:
                write_report(Path(args.report), deck, groups, doc, facts)
                print(f"[ok] reporte -> {_ascii(args.report)}")
            return _report(groups)
        if args.command == "measure":
            return _report([("Espacio", measure_deck(
                Path(args.deck), Path(args.shots) if args.shots else None))])
        doc = read_deck(Path(args.deck).read_text(encoding="utf-8"))
        Path(args.out).write_text(claims_table(doc), encoding="utf-8", newline="\n")
        print(f"[ok] {len(claims(doc))} frases -> {_ascii(args.out)}")
        return 0
    except (DeckError, OSError) as exc:
        print(f"[x] {_ascii(exc)}")
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
