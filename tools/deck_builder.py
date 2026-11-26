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
