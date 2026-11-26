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
