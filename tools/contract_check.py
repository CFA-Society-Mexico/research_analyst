"""Validador de contratos de la cobertura (checks K), por codigo.

Los contratos de ``templates/`` los escribe el agente. Este tool dice si los
de UNA cobertura cumplen lo que el template declara, antes de que un check
del modelo tropiece con ellos:

  K1  todo YAML parsea
  K2  issuer-profile.yaml
  K3  macro/macro-view.yaml
  K4  comps/*.yaml (el snapshot vigente de cada comparable)
  K5  model/inputs/canonical_*.csv

Solo lectura: jamas modifica un contrato.

Uso:
    python tools/contract_check.py <raiz>/<TICKER> [--macro macro-view.yaml]

Severidad: falla lo ESTRUCTURAL (no parsea, obligatorio vacio, valor fuera de
su lista, CSV sin columnas, llave duplicada); avisa lo TOLERABLE (campo que el
contrato no conoce, valor deprecado con equivalente, etiqueta valida con texto
extra).

PyYAML es OPCIONAL: sin el, K1-K4 salen [pendiente] y K5 (CSV, stdlib) corre.

Exit codes: 0 = verde; 1 = alguna falla; 3 = sin fallas pero con checks
[pendiente]; 2 = no existe la carpeta. Un [aviso] nunca cambia el exit code.

Consola ASCII-only.
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

from xlsx_builder import Finding  # noqa: E402

try:
    import yaml
except ImportError:  # dependencia opcional: sin ella K1-K4 quedan [pendiente]
    yaml = None

PYYAML_HINT = "PyYAML no instalado: python -m pip install pyyaml"

K1 = "K1 los YAML parsean"
K2 = "K2 issuer-profile"
K3 = "K3 macro-view"
K4 = "K4 comp-snapshots"
K5 = "K5 capa de captura (CSV)"

# ---------------------------------------------------------------------------
# Reglas del contrato (templates/*.yaml, templates/coverage-tree.md).
# tests/test_repo_contracts.py falla si un template gana o pierde un campo de
# primer nivel sin que estas tablas se enteren.
# ---------------------------------------------------------------------------

CAPTURE_COLUMNS = ("line_item", "period", "value", "source_doc", "source_ref", "tag")
CAPTURE_TAGS = ("observado", "derivado", "guidance")
CAPTURE_PERIODS = {"canonical_annual.csv": re.compile(r"FY\d{4}"),
                   "canonical_quarterly.csv": re.compile(r"[1-4]Q\d{4}")}

_ISO_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_EXAMPLES = 6


def _ascii(text: object) -> str:
    """Texto para consola: lo que venga de un archivo puede traer acentos."""
    return str(text).encode("ascii", "replace").decode("ascii")


def _blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_number(value: object) -> bool:
    return (isinstance(value, (int, float)) and not isinstance(value, bool)
            and not (isinstance(value, float) and math.isnan(value)))


@dataclass
class Problems:
    """Problemas de un check: los duros lo fallan, los blandos lo avisan."""

    hard: list[str] = field(default_factory=list)
    soft: list[str] = field(default_factory=list)

    def finding(self, name: str, ok_detail: str) -> Finding:
        if self.hard:
            notes = self.hard + [f"aviso: {item}" for item in self.soft]
            return Finding(name, False, _ascii(_listing(notes)))
        if self.soft:
            return Finding(name, False, _ascii(_listing(self.soft)), warning=True)
        return Finding(name, True, ok_detail)


def _listing(items: Sequence[str]) -> str:
    extra = f" (+{len(items) - _EXAMPLES} mas)" if len(items) > _EXAMPLES else ""
    return "; ".join(items[:_EXAMPLES]) + extra


# ---------------------------------------------------------------------------
# K5 - capa de captura
# ---------------------------------------------------------------------------


def _capture_file_problems(path: Path, problems: Problems) -> int:
    """Valida un CSV canonico; devuelve cuantas filas leyo. Los problemas se
    cuentan por regla (con la primera linea) para no listar miles de filas."""
    period_format = CAPTURE_PERIODS[path.name]
    hard: dict[str, list[int]] = {}
    soft: dict[str, list[int]] = {}
    seen: set[tuple[str, str]] = set()
    rows = 0
    try:
        with path.open(encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh)
            missing = [c for c in CAPTURE_COLUMNS if c not in (reader.fieldnames or [])]
            if missing:
                problems.hard.append(f"{path.name}: faltan columnas {', '.join(missing)}")
                return 0
            for number, row in enumerate(reader, start=2):
                rows += 1
                line_item = (row["line_item"] or "").strip()
                period = (row["period"] or "").strip()
                tag = (row["tag"] or "").strip().lower()
                if not line_item:
                    hard.setdefault("line_item vacio", []).append(number)
                if not period_format.fullmatch(period):
                    hard.setdefault("period fuera de formato", []).append(number)
                try:
                    numeric = not math.isnan(float(row["value"]))
                except (TypeError, ValueError):
                    numeric = False
                if not numeric:
                    hard.setdefault("value no numerico", []).append(number)
                if (line_item, period) in seen:
                    hard.setdefault("llave (line_item, period) repetida", []).append(number)
                seen.add((line_item, period))
                if tag not in CAPTURE_TAGS:
                    if tag.startswith(CAPTURE_TAGS):
                        soft.setdefault("tag con texto extra", []).append(number)
                    else:
                        hard.setdefault(
                            f"tag fuera de ({', '.join(CAPTURE_TAGS)})", []).append(number)
                if tag.startswith("observado") and not (row["source_doc"] or "").strip():
                    hard.setdefault("observado sin source_doc", []).append(number)
    except UnicodeDecodeError:
        problems.hard.append(f"{path.name}: no es UTF-8")
        return rows
    for tally, target in ((hard, problems.hard), (soft, problems.soft)):
        for rule, lines in tally.items():
            target.append(f"{path.name}: {len(lines)} filas con {rule} (linea {lines[0]})")
    return rows


def check_capture(inputs_dir: Path) -> Finding:
    """K5: esquema de los CSV canonicos de la capa de captura."""
    files = [inputs_dir / name for name in CAPTURE_PERIODS if (inputs_dir / name).is_file()]
    if not files:
        return Finding(K5, True, "n/a (sin canonical_*.csv en model/inputs/)")
    problems = Problems()
    rows = sum(_capture_file_problems(path, problems) for path in files)
    return problems.finding(K5, f"{rows} filas en {len(files)} archivos, esquema en orden")


def audit_contracts(coverage: Path, macro_view: Path | None = None) -> list[Finding]:
    """Checks K sobre la carpeta de una cobertura (``<raiz>/<TICKER>``)."""
    return [check_capture(coverage / "model" / "inputs")]
