# Validador de contratos de la cobertura (checks K) — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que `python tools/contract_check.py <raiz>/<TICKER>` diga, por código, si los contratos de una cobertura (`issuer-profile.yaml`, `macro-view.yaml`, `comps/*.yaml`, `canonical_*.csv`) cumplen lo que `templates/` declara.

**Architecture:** Un tool nuevo de solo lectura que emite cinco `Finding` (K1–K5) con los estados y exit codes de la casa. Las reglas viven en tablas dentro del tool; una prueba de contrato del repo las ata a los campos de `templates/*.yaml`. PyYAML es opcional: sin él, K1–K4 salen `[pendiente]` y K5 (CSV, stdlib) corre igual.

**Tech Stack:** Python 3.10+, stdlib (`csv`, `re`, `dataclasses`, `argparse`), PyYAML (opcional en runtime, declarada en `requirements.txt`), `unittest`.

**Spec:** `docs/superpowers/specs/2026-11-26-contract-check-design.md`

**Rama:** `feat/contract-check` (sale de `main` con el PR 57).

**Convenciones del repo** (CONTRIBUTING.md): stdout ASCII-only, incluido el `detail` de cada `Finding`; commits Conventional Commits, sujeto en español, ASCII, 72 caracteres o menos; suite con `python -m unittest discover -s tests`.

---

## Mapa de archivos

| Archivo | Responsabilidad |
|---|---|
| `tools/contract_check.py` (nuevo) | Tablas de reglas, carga de YAML/CSV, checks K1–K5, reporte y CLI |
| `tests/test_contract_check.py` (nuevo) | Cada check en los dos sentidos, modo sin PyYAML, CLI |
| `tests/test_repo_contracts.py` | Campos de `templates/*.yaml` = campos que conoce el validador; K documentados = emitidos |
| `requirements.txt` | `pyyaml>=6.0` |
| `skills/model-standards/references/integrity-checks.md` | Sección K |
| `commands/*.md`, `AGENTS.md`, `CONTRIBUTING.md` | Dónde se corre y la dependencia |

---

### Task 1: Esqueleto del tool y K5 (CSV canónicos)

K5 no depende de PyYAML: va primero y deja el andamiaje (`Problems`, reporte) que usan los demás.

**Files:**
- Create: `tools/contract_check.py`
- Create: `tests/test_contract_check.py`

- [ ] **Step 1: Escribir las pruebas que fallan**

Crear `tests/test_contract_check.py`:

```python
"""Pruebas de tools/contract_check.py: checks K en los dos sentidos.

Cada check tiene una cobertura que lo pasa y un contrato roto que lo falla o
lo avisa. Stdlib unittest, sin red. Las pruebas de YAML piden PyYAML (lo
instala requirements.txt); el modo degradado se prueba sin el.

Correr: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import contract_check as K  # noqa: E402
from xlsx_builder import Finding  # noqa: E402

PROFILE_OK = """\
ticker: "TICK"
name: "Emisora de prueba"
as_of: 2026-11-20
listing: "sec"
issuer_type: "non_financial"
adr: false
framework_governing: "us_gaap"
framework_confirmed_by_user: true
currency_reporting: "USD"
units: "USD millones"
fiscal_year_end: "09-27"
life_cycle_stage: "mature"
model_periodicity: "quarterly"
quarterly_history_years: 10
periodicity_confirmed_by_user: true
valuation_methods:
  dcf_fcff: true
  comps: true
industry_staleness_months: 6
materiality:
  critical_pct_target_price: 10
  capture_unlinked_pct: 0.5
mnpi_confirmed_public: true
"""

MACRO_OK = """\
updated_at: "2026-10-01"
owner: "analista"
staleness_warn_months: 3
rates:
  risk_free_10y: {value: 4.1, label: observado, source: "FRED DGS10"}
"""

COMP_OK = """\
ticker: "MSFT"
name: "Microsoft Corporation"
as_of: "2026-11-23 16:00 EST"
stale_days: 5
market_data:
  price: {value: 410.5, label: observado, source: "cierre 2026-11-23"}
  price_currency: "USD"
  shares_per_ads: {value: 1, label: observado, source: "accion ordinaria"}
  fx_to_reporting: {value: 1, source: "misma moneda", as_of: "2026-11-23"}
  shares_outstanding: {value: 7430000000, label: observado, source: "10-Q p.1"}
balance_items:
  total_debt: {value: 60000, label: observado, source: "10-Q p.5"}
  preferred_equity: {}
income_items:
  period: "ltm"
  revenue: {value: 245000, label: observado, source: "10-Q p.3"}
framework: "us_gaap"
"""

CSV_HEADER = "line_item,period,value,source_doc,source_ref,tag\n"
QUARTERLY_OK = CSV_HEADER + (
    "is_ns_total,1Q2026,100.0,10-Q acc 1,p.3,observado\n"
    "is_ns_total,2Q2026,103.0,10-Q acc 2,p.3,observado\n"
    "is_ni,4Q2025,9.0,derivado,FY menos 3Q,derivado\n")
ANNUAL_OK = CSV_HEADER + "is_ns_total,FY2025,400.0,10-K acc 9,p.30,observado\n"


class _Coverage(unittest.TestCase):
    """Base: una cobertura valida en disco; cada prueba rompe UNA pieza."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.coverage = self.root / "TICK"
        self._write("TICK/issuer-profile.yaml", PROFILE_OK)
        self._write("macro/macro-view.yaml", MACRO_OK)
        self._write("TICK/comps/MSFT.yaml", COMP_OK)
        self._write("TICK/model/inputs/canonical_quarterly.csv", QUARTERLY_OK)
        self._write("TICK/model/inputs/canonical_annual.csv", ANNUAL_OK)

    def _write(self, relative: str, text: str) -> Path:
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    def _checks(self) -> dict[str, Finding]:
        return {f.check.split()[0]: f for f in K.audit_contracts(self.coverage)}

    def assertFails(self, finding: Finding, *fragments: str) -> None:
        self.assertTrue(finding.blocking, finding.detail)
        for fragment in fragments:
            self.assertIn(fragment, finding.detail)

    def assertWarns(self, finding: Finding, *fragments: str) -> None:
        self.assertTrue(finding.warning, finding.detail)
        self.assertFalse(finding.blocking, finding.detail)
        for fragment in fragments:
            self.assertIn(fragment, finding.detail)


class CaptureTests(_Coverage):

    def _k5(self) -> Finding:
        return K.check_capture(self.coverage / "model" / "inputs")

    def test_valid_capture_is_green(self) -> None:
        finding = self._k5()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("4 filas en 2 archivos", finding.detail)

    def test_wide_format_fails_on_its_columns(self) -> None:
        self._write("TICK/model/inputs/canonical_quarterly.csv",
                    "line_item,region,1Q2016,2Q2016\naccounts_payable,consolidado,1,2\n")
        self.assertFails(self._k5(), "canonical_quarterly.csv: faltan columnas period, value")

    def test_period_must_match_the_file(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    ANNUAL_OK + "is_ni,4Q2025,9.0,10-K,p.1,observado\n")
        self.assertFails(self._k5(), "canonical_annual.csv", "period fuera de formato", "linea 3")

    def test_non_numeric_value_fails(self) -> None:
        for bad in ("n/d", "", "nan"):
            with self.subTest(value=bad):
                self._write("TICK/model/inputs/canonical_annual.csv",
                            CSV_HEADER + f"is_ni,FY2025,{bad},10-K,p.1,observado\n")
                self.assertFails(self._k5(), "value no numerico")

    def test_duplicated_key_fails(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    ANNUAL_OK + "is_ns_total,FY2025,401.0,10-K/A,p.30,observado\n")
        self.assertFails(self._k5(), "llave (line_item, period) repetida")

    def test_tag_outside_the_contract_fails(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,10-K,p.1,estimado\n")
        self.assertFails(self._k5(), "tag fuera de (observado, derivado, guidance)")

    def test_valid_tag_with_extra_text_warns(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,10-K,p.1,observado (signo normalizado)\n")
        self.assertWarns(self._k5(), "tag con texto extra")

    def test_observed_figure_needs_its_document(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,,p.1,observado\n")
        self.assertFails(self._k5(), "observado sin source_doc")

    def test_blank_line_item_fails(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + ",FY2025,9.0,10-K,p.1,observado\n")
        self.assertFails(self._k5(), "line_item vacio")

    def test_coverage_without_capture_does_not_apply(self) -> None:
        for name in ("canonical_quarterly.csv", "canonical_annual.csv"):
            (self.coverage / "model" / "inputs" / name).unlink()
        finding = self._k5()
        self.assertTrue(finding.ok)
        self.assertIn("n/a", finding.detail)

    def test_many_bad_rows_are_counted_not_listed(self) -> None:
        rows = "".join(f"is_ni,FY{2000 + i},9.0,10-K,p.1,estimado\n" for i in range(40))
        self._write("TICK/model/inputs/canonical_annual.csv", CSV_HEADER + rows)
        finding = self._k5()
        self.assertFails(finding, "40 filas con tag fuera de")
        self.assertLess(len(finding.detail), 400)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_contract_check`
Expected: `ModuleNotFoundError: No module named 'contract_check'`.

- [ ] **Step 3: Implementar**

Crear `tools/contract_check.py`:

```python
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
```

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_contract_check`
Expected: `OK` (11 pruebas).

- [ ] **Step 5: Commit**

```bash
git add tools/contract_check.py tests/test_contract_check.py
git commit -m "feat: contract_check valida el esquema de los CSV canonicos (K5)"
```

---

### Task 2: Carga de YAML, K1 y modo sin PyYAML

**Files:**
- Modify: `tools/contract_check.py`
- Modify: `tests/test_contract_check.py`
- Modify: `requirements.txt`

- [ ] **Step 1: Escribir las pruebas que fallan**

En `tests/test_contract_check.py`, agregar antes de `if __name__ == "__main__":`:

```python
# Mapa en linea sin cerrar: ningun parser de YAML lo acepta.
BROKEN_YAML = """\
updated_at: "2026-10-01"
rates:
  risk_free_10y: {value: 4.1, source: "FRED DGS10"
fx: 1
"""


@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class YamlParsingTests(_Coverage):

    def test_valid_coverage_passes_the_five_checks(self) -> None:
        found = self._checks()
        self.assertEqual(sorted(found), ["K1", "K2", "K3", "K4", "K5"])
        for check_id, finding in found.items():
            with self.subTest(check=check_id):
                self.assertTrue(finding.ok, finding.detail)

    def test_broken_yaml_fails_k1_with_file_and_line(self) -> None:
        self._write("macro/macro-view.yaml", BROKEN_YAML)
        found = self._checks()
        self.assertFails(found["K1"], "macro-view.yaml", "linea")
        # Lo que no parsea no se puede revisar por dentro: pendiente, no verde.
        self.assertTrue(found["K3"].pending)

    def test_yaml_that_is_not_a_mapping_fails_k1(self) -> None:
        self._write("TICK/comps/MSFT.yaml", "- solo\n- una lista\n")
        self.assertFails(self._checks()["K1"], "MSFT.yaml", "no es un mapa")

    def test_macro_flag_points_at_another_file(self) -> None:
        other = self._write("otra/macro-view.yaml", BROKEN_YAML)
        found = {f.check.split()[0]: f for f in K.audit_contracts(self.coverage, other)}
        self.assertFails(found["K1"], "macro-view.yaml")


class WithoutPyYamlTests(_Coverage):

    def test_yaml_checks_are_pending_and_csv_still_runs(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv",
                    CSV_HEADER + "is_ni,FY2025,9.0,10-K,p.1,estimado\n")
        with mock.patch.object(K, "yaml", None):
            found = self._checks()
        for check_id in ("K1", "K2", "K3", "K4"):
            with self.subTest(check=check_id):
                self.assertTrue(found[check_id].pending)
                self.assertIn("pip install pyyaml", found[check_id].detail)
        self.assertTrue(found["K5"].blocking)      # el CSV se valida igual
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_contract_check.YamlParsingTests tests.test_contract_check.WithoutPyYamlTests`
Expected: fallan las 5 (`audit_contracts` solo devuelve K5: `KeyError: 'K1'` o listas distintas).

- [ ] **Step 3: Implementar**

En `requirements.txt`, agregar una línea:

```
pyyaml>=6.0
```

En `tools/contract_check.py`, insertar antes del comentario de sección `# K5 - capa de captura`:

```python
# ---------------------------------------------------------------------------
# Carga de YAML y K1
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Loaded:
    """Un contrato YAML leido: ``data`` es None si no parsea o no es un mapa."""

    path: Path
    data: dict[str, Any] | None
    error: str = ""


def load_yaml(path: Path) -> Loaded:
    """Lee un contrato YAML. Requiere PyYAML (el llamador ya lo comprobo)."""
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except UnicodeDecodeError:
        return Loaded(path, None, f"{path.name}: no es UTF-8")
    except yaml.YAMLError as exc:
        mark = getattr(exc, "problem_mark", None)
        where = f" linea {mark.line + 1}, columna {mark.column + 1}" if mark else ""
        problem = getattr(exc, "problem", None) or "YAML invalido"
        return Loaded(path, None, f"{path.name}{where}: {problem}")
    if not isinstance(data, dict):
        return Loaded(path, None, f"{path.name}: el documento no es un mapa de campos")
    return Loaded(path, data)


def check_yaml_parses(loaded: Sequence[Loaded]) -> Finding:
    """K1: todo contrato YAML de la cobertura es YAML valido."""
    problems = Problems(hard=[item.error for item in loaded if item.data is None])
    return problems.finding(K1, f"{len(loaded)} archivos YAML validos")


def _unparsed(name: str) -> Finding:
    return Finding(name, False, "no evaluado: el archivo no parsea (ver K1)", pending=True)
```

Reemplazar `audit_contracts` completa por:

```python
def audit_contracts(coverage: Path, macro_view: Path | None = None) -> list[Finding]:
    """Checks K sobre la carpeta de una cobertura (``<raiz>/<TICKER>``)."""
    capture = check_capture(coverage / "model" / "inputs")
    if yaml is None:
        return [*(Finding(name, False, PYYAML_HINT, pending=True) for name in (K1, K2, K3, K4)),
                capture]
    macro_path = macro_view or coverage.parent / "macro" / "macro-view.yaml"
    profile_path = coverage / "issuer-profile.yaml"
    comps_dir = coverage / "comps"
    profile = load_yaml(profile_path) if profile_path.is_file() else None
    macro = load_yaml(macro_path) if macro_path.is_file() else None
    comps = [load_yaml(p) for p in sorted(comps_dir.glob("*.yaml"))] if comps_dir.is_dir() else []
    loaded = [item for item in (profile, macro, *comps) if item is not None]
    return [check_yaml_parses(loaded), check_profile(profile), check_macro(macro),
            check_comps(comps), capture]


def check_profile(profile: Loaded | None) -> Finding:
    if profile is not None and profile.data is None:
        return _unparsed(K2)
    return Finding(K2, True, "n/a (reglas en la Task 3)")


def check_macro(macro: Loaded | None) -> Finding:
    if macro is not None and macro.data is None:
        return _unparsed(K3)
    return Finding(K3, True, "n/a (reglas en la Task 4)")


def check_comps(comps: Sequence[Loaded]) -> Finding:
    return Finding(K4, True, "n/a (reglas en la Task 4)")
```

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_contract_check`
Expected: `OK` (16 pruebas).

- [ ] **Step 5: Commit**

```bash
git add tools/contract_check.py tests/test_contract_check.py requirements.txt
git commit -m "feat: contract_check carga YAML, K1 y modo sin PyYAML"
```

---

### Task 3: K2 — `issuer-profile.yaml`

**Files:**
- Modify: `tools/contract_check.py`
- Modify: `tests/test_contract_check.py`

- [ ] **Step 1: Escribir las pruebas que fallan**

Agregar antes de `class WithoutPyYamlTests`:

```python
@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class ProfileTests(_Coverage):

    def _k2(self, text: str | None = None) -> Finding:
        if text is not None:
            self._write("TICK/issuer-profile.yaml", text)
        return self._checks()["K2"]

    def test_required_field_left_blank_fails(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace('units: "USD millones"', 'units: ""')),
                         "units vacio")

    def test_required_field_missing_fails(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace('listing: "sec"\n', "")), "listing vacio")

    def test_value_outside_its_list_fails(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace('listing: "sec"', 'listing: "nyse"')),
                         "listing: 'nyse' no esta en (bmv, sec, dual, private)")

    def test_deprecated_periodicity_warns_with_its_equivalent(self) -> None:
        text = PROFILE_OK.replace('model_periodicity: "quarterly"',
                                  'model_periodicity: "annual_plus_quarterly"')
        self.assertWarns(self._k2(text), "annual_plus_quarterly", "quarterly")

    def test_unconfirmed_profile_fails(self) -> None:
        for flag in K.PROFILE_CONFIRMATIONS:
            with self.subTest(flag=flag):
                self.assertFails(self._k2(PROFILE_OK.replace(f"{flag}: true", f"{flag}: false")),
                                 f"{flag} no esta en true")

    def test_unknown_field_warns(self) -> None:
        self.assertWarns(self._k2(PROFILE_OK + 'listing_notes: "ADR nivel 3"\n'),
                         "campo desconocido: listing_notes")

    def test_dates_and_fiscal_year_end_have_a_format(self) -> None:
        self.assertFails(self._k2(PROFILE_OK.replace("as_of: 2026-11-20", 'as_of: "noviembre"')),
                         "as_of sin fecha YYYY-MM-DD")
        self.assertFails(self._k2(PROFILE_OK.replace('"09-27"', '"septiembre"')),
                         "fiscal_year_end no es MM-DD")

    def test_typed_fields_reject_the_wrong_type(self) -> None:
        cases = {
            "quarterly_history_years: 10": ("quarterly_history_years: -1", "quarterly_history_years"),
            "adr: false": ('adr: "no"', "adr no es booleano"),
            "  comps: true": ('  comps: "si"', "valuation_methods.comps"),
            "  capture_unlinked_pct: 0.5": ("  capture_unlinked_pct: -2", "materiality.capture_unlinked_pct"),
        }
        for good, (bad, fragment) in cases.items():
            with self.subTest(field=fragment):
                self.assertFails(self._k2(PROFILE_OK.replace(good, bad)), fragment)

    def test_optional_field_absent_is_not_a_finding(self) -> None:
        finding = self._k2(PROFILE_OK.replace("quarterly_history_years: 10\n", ""))
        self.assertTrue(finding.ok, finding.detail)

    def test_coverage_without_profile_fails(self) -> None:
        (self.coverage / "issuer-profile.yaml").unlink()
        self.assertFails(self._checks()["K2"], "no existe issuer-profile.yaml")

    def test_detail_stays_ascii_with_accented_values(self) -> None:
        finding = self._k2(PROFILE_OK.replace('listing: "sec"', 'listing: "bolsa de México"'))
        self.assertTrue(finding.detail.isascii(), finding.detail)
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_contract_check.ProfileTests`
Expected: fallan 9 de 11 (K2 devuelve siempre `n/a`); pasan `test_optional_field_absent_is_not_a_finding` y `test_detail_stays_ascii_with_accented_values`.

- [ ] **Step 3: Implementar**

En `tools/contract_check.py`, después de la constante `CAPTURE_PERIODS`:

```python
PROFILE_REQUIRED = ("ticker", "name", "as_of", "listing", "issuer_type",
                    "framework_governing", "currency_reporting", "units",
                    "fiscal_year_end", "model_periodicity")
PROFILE_ENUMS = {
    "listing": ("bmv", "sec", "dual", "private"),
    "issuer_type": ("non_financial", "bank", "insurer", "fibra", "private"),
    "framework_governing": ("ifrs", "us_gaap", "nif", "cnbv_criteria"),
    "model_periodicity": ("annual", "quarterly"),
    "life_cycle_stage": ("embryonic", "growth", "mature", "decline"),
}
# valor deprecado -> como se trata hoy
PROFILE_DEPRECATED = {"model_periodicity": {"annual_plus_quarterly": "quarterly"}}
PROFILE_CONFIRMATIONS = ("framework_confirmed_by_user", "periodicity_confirmed_by_user",
                         "mnpi_confirmed_public")
PROFILE_POSITIVE_INTS = ("quarterly_history_years", "industry_staleness_months")
PROFILE_KEYS = frozenset(PROFILE_REQUIRED + PROFILE_CONFIRMATIONS + PROFILE_POSITIVE_INTS + (
    "adr", "framework_notes", "currency_functional", "segments", "life_cycle_stage",
    "valuation_methods", "materiality", "retention_note", "sources"))
```

Después de `_is_number`:

```python
def _date_of(value: object) -> date | None:
    """Fecha de un campo: PyYAML ya convierte ``2026-11-20`` sin comillas a
    ``date``; entre comillas llega como texto ('2026-08-31 11:05 KST')."""
    if isinstance(value, date):
        return value
    found = _ISO_DATE.search(value) if isinstance(value, str) else None
    if found is None:
        return None
    try:
        return date(int(found.group(1)), int(found.group(2)), int(found.group(3)))
    except ValueError:
        return None


def _positive_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _unknown_fields(data: dict[str, Any], known: frozenset[str], prefix: str = "") -> list[str]:
    return [f"campo desconocido: {prefix}{key}" for key in sorted(map(str, set(data) - known))]
```

Reemplazar `check_profile` completa por:

```python
def check_profile(profile: Loaded | None) -> Finding:
    """K2: issuer-profile.yaml cumple su contrato."""
    if profile is None:
        return Finding(K2, False, "no existe issuer-profile.yaml en la cobertura")
    if profile.data is None:
        return _unparsed(K2)
    data, problems = profile.data, Problems()
    problems.hard += [f"{key} vacio" for key in PROFILE_REQUIRED if _blank(data.get(key))]
    for key, allowed in PROFILE_ENUMS.items():
        value = data.get(key)
        if _blank(value):
            continue
        replacement = PROFILE_DEPRECATED.get(key, {}).get(value)
        if replacement:
            problems.soft.append(f"{key}: {value} esta deprecado (se trata como {replacement})")
        elif value not in allowed:
            problems.hard.append(f"{key}: {value!r} no esta en ({', '.join(allowed)})")
    if not _blank(data.get("as_of")) and _date_of(data["as_of"]) is None:
        problems.hard.append("as_of sin fecha YYYY-MM-DD")
    fiscal_year_end = data.get("fiscal_year_end")
    if not _blank(fiscal_year_end) and not re.fullmatch(r"\d{2}-\d{2}", str(fiscal_year_end)):
        problems.hard.append("fiscal_year_end no es MM-DD")
    problems.hard += [f"{key} no esta en true (falta la confirmacion del analista)"
                      for key in PROFILE_CONFIRMATIONS if data.get(key) is not True]
    if "adr" in data and not isinstance(data["adr"], bool):
        problems.hard.append("adr no es booleano")
    problems.hard += [f"{key} debe ser un entero > 0" for key in PROFILE_POSITIVE_INTS
                      if key in data and not _positive_int(data[key])]
    for key, value in (data.get("materiality") or {}).items():
        if not _is_number(value) or value < 0:
            problems.hard.append(f"materiality.{key} debe ser un numero >= 0")
    for key, value in (data.get("valuation_methods") or {}).items():
        if not isinstance(value, bool):
            problems.hard.append(f"valuation_methods.{key} no es booleano")
    problems.soft += _unknown_fields(data, PROFILE_KEYS)
    return problems.finding(K2, "obligatorios, listas y confirmaciones en orden")
```

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_contract_check`
Expected: `OK` (27 pruebas).

- [ ] **Step 5: Commit**

```bash
git add tools/contract_check.py tests/test_contract_check.py
git commit -m "feat: check K2 valida issuer-profile.yaml contra su contrato"
```

---

### Task 4: K3 (`macro-view.yaml`) y K4 (`comps/*.yaml`)

**Files:**
- Modify: `tools/contract_check.py`
- Modify: `tests/test_contract_check.py`

- [ ] **Step 1: Escribir las pruebas que fallan**

Agregar antes de `class WithoutPyYamlTests`:

```python
@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class MacroTests(_Coverage):

    def test_macro_view_without_date_fails(self) -> None:
        self._write("macro/macro-view.yaml", MACRO_OK.replace('"2026-10-01"', '""'))
        self.assertFails(self._checks()["K3"], "updated_at vacio o sin fecha")

    def test_missing_macro_view_warns(self) -> None:
        (self.root / "macro" / "macro-view.yaml").unlink()
        self.assertWarns(self._checks()["K3"], "sin macro-view.yaml")

    def test_unknown_field_and_bad_window(self) -> None:
        self._write("macro/macro-view.yaml", MACRO_OK + 'gate_status: "propuesta"\n')
        self.assertWarns(self._checks()["K3"], "campo desconocido: gate_status")
        self._write("macro/macro-view.yaml",
                    MACRO_OK.replace("staleness_warn_months: 3", "staleness_warn_months: 0"))
        self.assertFails(self._checks()["K3"], "staleness_warn_months")


@unittest.skipIf(K.yaml is None, "PyYAML no instalado")
class CompTests(_Coverage):

    def _k4(self, text: str | None = None, name: str = "MSFT.yaml") -> Finding:
        if text is not None:
            self._write(f"TICK/comps/{name}", text)
        return self._checks()["K4"]

    @staticmethod
    def _without(text: str, *keys: str) -> str:
        return "".join(line for line in text.splitlines(keepends=True)
                       if not line.strip().startswith(tuple(f"{k}:" for k in keys)))

    def test_missing_mandatory_market_fields_are_named(self) -> None:
        text = self._without(COMP_OK, "price_currency", "shares_per_ads", "fx_to_reporting")
        self.assertFails(self._k4(text), "MSFT.yaml",
                         "market_data sin price_currency, shares_per_ads, fx_to_reporting")

    def test_old_currency_key_is_reported_as_unknown(self) -> None:
        text = COMP_OK.replace('  price_currency: "USD"',
                               '  price_currency: "USD"\n  currency: "USD"')
        self.assertWarns(self._k4(text), "campo desconocido: market_data.currency")

    def test_price_currency_must_be_iso_4217(self) -> None:
        self.assertFails(self._k4(COMP_OK.replace('price_currency: "USD"', 'price_currency: "usd"')),
                         "price_currency no es ISO 4217")

    def test_captured_datum_needs_value_and_source(self) -> None:
        no_source = COMP_OK.replace('{value: 410.5, label: observado, source: "cierre 2026-11-23"}',
                                    "{value: 410.5, label: observado}")
        self.assertFails(self._k4(no_source), "market_data.price sin source")
        no_value = COMP_OK.replace("{value: 60000, label: observado", "{value: n/d, label: observado")
        self.assertFails(self._k4(no_value), "balance_items.total_debt sin value numerico")

    def test_fx_needs_its_as_of(self) -> None:
        text = COMP_OK.replace('{value: 1, source: "misma moneda", as_of: "2026-11-23"}',
                               '{value: 1, source: "misma moneda"}')
        self.assertFails(self._k4(text), "market_data.fx_to_reporting sin as_of")

    def test_snapshot_needs_ticker_and_dated_as_of(self) -> None:
        self.assertFails(self._k4(COMP_OK.replace('as_of: "2026-11-23 16:00 EST"', 'as_of: ""')),
                         "as_of sin fecha")

    def test_foreign_framework_warns(self) -> None:
        self.assertWarns(self._k4(COMP_OK.replace('framework: "us_gaap"', 'framework: "k_ifrs"')),
                         "framework", "k_ifrs")

    def test_only_the_latest_snapshot_of_a_comparable_is_checked(self) -> None:
        old = self._without(COMP_OK, "price_currency").replace("2026-11-23 16:00 EST", "2026-08-01")
        self._write("TICK/comps/MSFT_2026-08-01.yaml", old)   # historia: incompleto, no cuenta
        finding = self._k4()
        self.assertTrue(finding.ok, finding.detail)
        self.assertIn("1 snapshots vigentes", finding.detail)

    def test_coverage_without_comps_does_not_apply(self) -> None:
        (self.coverage / "comps" / "MSFT.yaml").unlink()
        finding = self._k4()
        self.assertTrue(finding.ok)
        self.assertIn("n/a", finding.detail)

    def test_empty_optional_datum_is_not_a_finding(self) -> None:
        finding = self._k4()                                   # preferred_equity: {}
        self.assertTrue(finding.ok, finding.detail)
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_contract_check.MacroTests tests.test_contract_check.CompTests`
Expected: fallan 11 de 13 (K3 y K4 devuelven `n/a`); pasan `test_coverage_without_comps_does_not_apply` y `test_empty_optional_datum_is_not_a_finding`.

- [ ] **Step 3: Implementar**

En `tools/contract_check.py`, después de `PROFILE_KEYS`:

```python
MACRO_KEYS = frozenset({"updated_at", "owner", "staleness_warn_months", "rates", "fx",
                        "inflation", "gdp", "commodity_decks", "equity_risk_premium", "notes"})

COMP_KEYS = frozenset({"ticker", "name", "selected_because", "as_of", "stale_days",
                       "market_data", "balance_items", "income_items", "framework",
                       "comparability_notes"})
COMP_MARKET_REQUIRED = ("price", "price_currency", "shares_per_ads", "fx_to_reporting",
                        "shares_outstanding")
COMP_MARKET_KEYS = frozenset(COMP_MARKET_REQUIRED + ("market_cap",))
# Datos de mercado con forma {value, source}; price_currency es un texto.
COMP_MARKET_DATA = ("price", "shares_per_ads", "fx_to_reporting", "shares_outstanding",
                    "market_cap")
COMP_FRAMEWORKS = ("ifrs", "us_gaap", "nif")
_ISO_4217 = re.compile(r"[A-Z]{3}")
_DATED_SUFFIX = re.compile(r"_\d{4}-\d{2}-\d{2}$")
```

Reemplazar `check_macro` y `check_comps` completas por:

```python
def check_macro(macro: Loaded | None) -> Finding:
    """K3: macro/macro-view.yaml cumple su contrato."""
    if macro is None:
        return Finding(K3, False, "sin macro-view.yaml en la carpeta macro/ del workspace",
                       warning=True)
    if macro.data is None:
        return _unparsed(K3)
    data, problems = macro.data, Problems()
    if _date_of(data.get("updated_at")) is None:
        problems.hard.append("updated_at vacio o sin fecha")
    if "staleness_warn_months" in data and not _positive_int(data["staleness_warn_months"]):
        problems.hard.append("staleness_warn_months debe ser un entero > 0")
    problems.soft += _unknown_fields(data, MACRO_KEYS)
    return problems.finding(K3, "updated_at y plazo en orden")


def _empty_datum(value: object) -> bool:
    """Dato no capturado: ausente, texto vacio o ``{}``."""
    return _blank(value) or value == {}


def _datum_problem(name: str, value: object) -> str | None:
    """Un dato capturado lleva ``value`` numerico y ``source``."""
    if not isinstance(value, dict):
        return f"{name} no es {{value, source}}"
    if not _is_number(value.get("value")):
        return f"{name} sin value numerico"
    if _blank(value.get("source")):
        return f"{name} sin source"
    return None


def _snapshot_problems(data: dict[str, Any]) -> tuple[list[str], list[str]]:
    hard: list[str] = []
    soft: list[str] = []
    if _blank(data.get("ticker")):
        hard.append("ticker vacio")
    if _date_of(data.get("as_of")) is None:
        hard.append("as_of sin fecha")
    if "stale_days" in data and not _positive_int(data["stale_days"]):
        hard.append("stale_days debe ser un entero > 0")
    market = data.get("market_data")
    if not isinstance(market, dict):
        hard.append("sin market_data")
        market = {}
    else:
        missing = [key for key in COMP_MARKET_REQUIRED if _empty_datum(market.get(key))]
        if missing:
            hard.append("market_data sin " + ", ".join(missing))
        currency = market.get("price_currency")
        if not _blank(currency) and not _ISO_4217.fullmatch(str(currency)):
            hard.append("price_currency no es ISO 4217 (tres letras mayusculas)")
        for key in COMP_MARKET_DATA:
            if not _empty_datum(market.get(key)):
                problem = _datum_problem(f"market_data.{key}", market[key])
                if problem:
                    hard.append(problem)
        fx = market.get("fx_to_reporting")
        if isinstance(fx, dict) and fx and _blank(fx.get("as_of")):
            hard.append("market_data.fx_to_reporting sin as_of")
        soft += _unknown_fields(market, COMP_MARKET_KEYS, prefix="market_data.")
    for section in ("balance_items", "income_items"):
        for key, value in (data.get(section) or {}).items():
            if key == "period" or _empty_datum(value):
                continue
            problem = _datum_problem(f"{section}.{key}", value)
            if problem:
                hard.append(problem)
    framework = data.get("framework")
    if not _blank(framework) and framework not in COMP_FRAMEWORKS:
        soft.append(f"framework {framework!r} fuera de ({', '.join(COMP_FRAMEWORKS)})")
    soft += _unknown_fields(data, COMP_KEYS)
    return hard, soft


def _current_snapshots(comps: Sequence[Loaded]) -> list[Loaded]:
    """El snapshot vigente de cada comparable: el de ``as_of`` mas reciente
    (mismo criterio que D3). Los anteriores son historia y no se validan."""
    latest: dict[str, tuple[date, Loaded]] = {}
    for item in comps:
        if item.data is None:
            continue
        key = str(item.data.get("ticker") or _DATED_SUFFIX.sub("", item.path.stem))
        as_of = _date_of(item.data.get("as_of")) or date.min
        if key not in latest or as_of > latest[key][0]:
            latest[key] = (as_of, item)
    return [item for _, item in sorted(latest.values(), key=lambda pair: pair[1].path.name)]


def check_comps(comps: Sequence[Loaded]) -> Finding:
    """K4: el snapshot vigente de cada comparable cumple su contrato."""
    if not comps:
        return Finding(K4, True, "n/a (sin snapshots en comps/)")
    current = _current_snapshots(comps)
    if not current:
        return _unparsed(K4)
    problems = Problems()
    for item in current:
        hard, soft = _snapshot_problems(item.data or {})
        if hard:
            problems.hard.append(f"{item.path.name}: " + ", ".join(hard))
        if soft:
            problems.soft.append(f"{item.path.name}: " + ", ".join(soft))
    return problems.finding(K4, f"{len(current)} snapshots vigentes completos")
```

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_contract_check`
Expected: `OK` (40 pruebas).

- [ ] **Step 5: Commit**

```bash
git add tools/contract_check.py tests/test_contract_check.py
git commit -m "feat: checks K3 y K4 validan macro-view y los comp-snapshots"
```

---

### Task 5: CLI, reporte y contrato del repo

**Files:**
- Modify: `tools/contract_check.py`
- Modify: `tests/test_contract_check.py`
- Modify: `tests/test_repo_contracts.py`

- [ ] **Step 1: Escribir las pruebas que fallan**

En `tests/test_contract_check.py`, agregar antes de `if __name__ == "__main__":`:

```python
class CliTests(_Coverage):

    def _run(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = K.main(["contract_check.py", *args])
        return code, out.getvalue()

    @unittest.skipIf(K.yaml is None, "PyYAML no instalado")
    def test_valid_coverage_exits_green(self) -> None:
        code, out = self._run(str(self.coverage))
        self.assertEqual(code, 0)
        self.assertIn("[ok] K5", out)
        self.assertIn("Veredicto: VERDE", out)
        self.assertTrue(out.isascii())

    @unittest.skipIf(K.yaml is None, "PyYAML no instalado")
    def test_warning_does_not_change_the_exit_code(self) -> None:
        self._write("TICK/issuer-profile.yaml", PROFILE_OK + 'listing_notes: "x"\n')
        code, out = self._run(str(self.coverage))
        self.assertEqual(code, 0)
        self.assertIn("[aviso] K2", out)
        self.assertIn("1 avisos", out)

    def test_broken_contract_exits_failure(self) -> None:
        self._write("TICK/model/inputs/canonical_annual.csv", "line_item,region\nx,y\n")
        code, out = self._run(str(self.coverage))
        self.assertEqual(code, 1)
        self.assertIn("[x] K5", out)
        self.assertIn("Veredicto: FALLA", out)

    def test_without_pyyaml_exits_pending(self) -> None:
        with mock.patch.object(K, "yaml", None):
            code, out = self._run(str(self.coverage))
        self.assertEqual(code, 3)
        self.assertIn("[pendiente] K1", out)

    def test_missing_folder_is_reported_without_traceback(self) -> None:
        code, out = self._run(str(self.root / "NOEXISTE"))
        self.assertEqual(code, 2)
        self.assertIn("[x] no existe la carpeta de cobertura", out)
```

En `tests/test_repo_contracts.py`:

Reemplazar

```python
_CHECK_ROW = re.compile(r"^\| ([SCDF]\d+[a-z]?) \|", re.MULTILINE)
```

por

```python
_CHECK_ROW = re.compile(r"^\| ([SCDFK]\d+[a-z]?) \|", re.MULTILINE)
```

Después de la línea `CHECKER = ROOT / "tools" / "model_check.py"`:

```python
CONTRACT_CHECKER = ROOT / "tools" / "contract_check.py"
```

Y agregar, antes de `if __name__ == "__main__":`:

```python
class ContractCheckTests(unittest.TestCase):
    """El validador de contratos conoce exactamente los campos de los templates."""

    @classmethod
    def setUpClass(cls) -> None:
        sys.path.insert(0, str(ROOT / "tools"))
        import contract_check
        cls.K = contract_check

    def _template(self, name: str) -> dict:
        if self.K.yaml is None:
            self.skipTest("PyYAML no instalado")
        return self.K.yaml.safe_load(_read(f"templates/{name}"))

    def test_validator_knows_exactly_the_template_fields(self) -> None:
        for name, known in (("issuer-profile.yaml", self.K.PROFILE_KEYS),
                            ("macro-view.yaml", self.K.MACRO_KEYS),
                            ("comp-snapshot.yaml", self.K.COMP_KEYS)):
            with self.subTest(template=name):
                self.assertEqual(set(self._template(name)), set(known),
                                 "campos del template vs campos que conoce contract_check.py")

    def test_validator_knows_the_market_data_fields(self) -> None:
        market = self._template("comp-snapshot.yaml")["market_data"]
        self.assertEqual(set(market), set(self.K.COMP_MARKET_KEYS))

    def test_documented_k_checks_are_exactly_the_emitted_ones(self) -> None:
        documented = {i for i in _CHECK_ROW.findall(CHECKS_DOC.read_text(encoding="utf-8"))
                      if i.startswith("K")}
        emitted = set(re.findall(r'"(K\d+) ', CONTRACT_CHECKER.read_text(encoding="utf-8")))
        self.assertEqual(emitted, documented,
                         "checks K documentados vs emitidos por tools/contract_check.py")
```

Agregar `import sys` al bloque de imports de `tests/test_repo_contracts.py` si no está.

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_contract_check.CliTests tests.test_repo_contracts.ContractCheckTests`
Expected: `CliTests` falla con `AttributeError: module 'contract_check' has no attribute 'main'`; `test_documented_k_checks...` falla (K emitidos, ninguno documentado todavía). Los dos de campos de templates pasan.

- [ ] **Step 3: Implementar**

Al final de `tools/contract_check.py`:

```python
# ---------------------------------------------------------------------------
# Reporte y CLI
# ---------------------------------------------------------------------------


def print_report(findings: Sequence[Finding]) -> int:
    """Imprime el reporte. Exit: 1 fallas; 3 solo pendientes; 0 verde.
    Los avisos se listan y se cuentan, pero no cambian el exit code."""
    for finding in findings:
        print(f"{finding.mark} {finding.check}: {finding.detail}")
    failures = sum(1 for f in findings if f.blocking)
    pending = sum(1 for f in findings if f.pending)
    warnings = sum(1 for f in findings if f.warning and not f.pending)
    ok = sum(1 for f in findings if f.ok and not f.pending)
    print(f"Resumen: {ok} ok, {failures} fallas, {pending} pendientes, {warnings} avisos")
    if failures:
        print("Veredicto: FALLA")
        return 1
    if pending:
        print("Veredicto: PENDIENTE - no es verde; resuelve los pendientes y re-valida")
        return 3
    print("Veredicto: VERDE (contratos de la cobertura)")
    return 0


def main(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Checks K: contratos de una cobertura contra templates/ (solo lectura).")
    parser.add_argument("coverage", help="carpeta de la cobertura: <raiz>/<TICKER>")
    parser.add_argument("--macro",
                        help="macro-view.yaml (default: <raiz>/macro/macro-view.yaml)")
    args = parser.parse_args(list(argv[1:]))
    coverage = Path(args.coverage)
    if not coverage.is_dir():
        print(f"[x] no existe la carpeta de cobertura: {_ascii(args.coverage)}")
        return 2
    return print_report(audit_contracts(coverage, Path(args.macro) if args.macro else None))


if __name__ == "__main__":
    sys.exit(main(sys.argv))
```

`test_documented_k_checks...` sigue en rojo hasta la Task 6 (la sección K del documento). Es esperado: va en el mismo commit que la doctrina.

- [ ] **Step 4: Correr las pruebas del tool**

Run: `python -m unittest tests.test_contract_check`
Expected: `OK` (45 pruebas).

- [ ] **Step 5: No se commitea todavía**

La prueba de contrato del repo exige la sección K documentada. La Task 6 la escribe y cierra el commit.

---

### Task 6: Doctrina y commit del CLI

Sin código. Cada reemplazo es texto exacto: si el texto "viejo" no aparece una sola vez, detente y revisa.

**Files:**
- Modify: `skills/model-standards/references/integrity-checks.md`
- Modify: `commands/update-quarter.md`, `commands/model-check.md`, `commands/update-industry.md`, `commands/update-macro.md`, `commands/init-coverage.md`
- Modify: `AGENTS.md`, `CONTRIBUTING.md`

- [ ] **Step 1: `integrity-checks.md` — sección K**

Insertar inmediatamente antes de la línea `## Reporte de /model-check`:

```
## De contratos de la cobertura (checks K — `tools/contract_check.py`)

Corren con `python tools/contract_check.py <raiz>/<TICKER>`, sobre los insumos
y no sobre el xlsx: sirven antes de que exista un modelo. Solo lectura, mismos
exit codes. Falla lo estructural; avisa lo tolerable (campo que el contrato no
conoce, valor deprecado con equivalente, etiqueta válida con texto extra).
PyYAML es opcional: sin él K1–K4 salen `[pendiente]` y K5 corre.

| # | Check | Regla |
|---|---|---|
| K1 | Todo YAML parsea | `issuer-profile.yaml`, `macro-view.yaml` y cada `comps/*.yaml` son YAML válido y un mapa de campos. Falla con archivo, línea y columna. Un archivo que no parsea deja su check de contenido `[pendiente]` |
| K2 | `issuer-profile.yaml` | Obligatorios no vacíos (ticker, name, as_of, listing, issuer_type, framework_governing, currency_reporting, units, fiscal_year_end, model_periodicity); valores de su lista; las tres confirmaciones del analista en `true`; tipos de los campos numéricos y booleanos. Aviso: campo desconocido, `annual_plus_quarterly` (deprecado) |
| K3 | `macro/macro-view.yaml` | `updated_at` con fecha; `staleness_warn_months` entero > 0 si está. Sin archivo o con campo desconocido: aviso |
| K4 | `comps/*.yaml` | Sobre el snapshot VIGENTE de cada comparable: `ticker`, `as_of` con fecha, los cinco campos obligatorios de `market_data` (price, price_currency ISO 4217, shares_per_ads, fx_to_reporting con `as_of`, shares_outstanding) y cada dato capturado con `value` numérico y `source`. Aviso: `framework` fuera de lista, campo desconocido |
| K5 | `model/inputs/canonical_*.csv` | Las seis columnas del contrato; `period` `FY2016` en el anual y `3Q2016` en el trimestral; `value` numérico; llave `(line_item, period)` única; `tag` de la lista; `observado` con `source_doc`. Aviso: etiqueta válida con texto extra |

Las tablas de reglas viven en el tool; `tests/test_repo_contracts.py` falla si
un template gana o pierde un campo sin que el validador se entere.

```

- [ ] **Step 2: comandos — dónde se corre**

En `commands/update-quarter.md`, reemplazar

```
1. Árbol de cobertura e `issuer-profile.yaml` existen y el perfil está confirmado.
```

por

```
1. Árbol de cobertura e `issuer-profile.yaml` existen y el perfil está confirmado:
   `python tools/contract_check.py <raiz>/<TICKER>` sin fallas (checks K; un
   `[aviso]` se reporta y no detiene).
```

En `commands/model-check.md`, reemplazar

```
   checks D1, D2, D7). Si es modelo externo: los checks D que dependen de contratos
   del plugin se reportan como `n/a`, no como falla.
```

por

```
   checks D1, D2, D7). Si es modelo externo: los checks D que dependen de contratos
   del plugin se reportan como `n/a`, no como falla.
3. Si el modelo vive en una cobertura: `python tools/contract_check.py
   <raiz>/<TICKER>` (checks K) y su resultado va en el reporte, antes de los S.
   Un contrato roto explica la mitad de los rojos del modelo.
```

En `commands/update-industry.md`, reemplazar

```
2. `macro-view.yaml` dentro de staleness (aviso, no bloqueo — igual que D6).
```

por

```
2. `macro-view.yaml` dentro de staleness (aviso, no bloqueo — igual que D6).
3. Contratos sin fallas: `python tools/contract_check.py <raiz>/<TICKER>`.
```

En `commands/update-macro.md`, reemplazar

```
1. `macro/` existe con `macro-view.yaml` (si no: crear estructura desde
   template y avisar — no hay nada que actualizar todavía).
```

por

```
1. `macro/` existe con `macro-view.yaml` (si no: crear estructura desde
   template y avisar — no hay nada que actualizar todavía). El archivo debe
   parsear: correr `python tools/contract_check.py <raiz>/<TICKER>` sobre una
   cobertura del workspace y resolver K1/K3 ANTES de proponer cambios — editar
   un YAML roto lo rompe más.
```

En `commands/init-coverage.md`, reemplazar

```
Los pasos 3 y 4 pueden correr en paralelo (la máquina captura mientras el analista
piensa industria); sus gates cierran por separado y AMBOS deben estar cerrados antes
del paso 5.
```

por

```
Los pasos 3 y 4 pueden correr en paralelo (la máquina captura mientras el analista
piensa industria); sus gates cierran por separado y AMBOS deben estar cerrados antes
del paso 5.

Al cerrar los gates de los pasos 2 y 3, y antes del paso 5:
`python tools/contract_check.py <raiz>/<TICKER>` sin fallas (checks K). Antes
del gate del paso 2 el perfil no está confirmado y K2 falla por diseño.
```

- [ ] **Step 3: `AGENTS.md` y `CONTRIBUTING.md`**

En `AGENTS.md`, reemplazar

```
- Historia financiera masiva de emisoras SEC: `python tools/xbrl_fetch.py <TICKER>`
```

por

```
- Contratos de la cobertura (perfil, macro-view, comps, CSV canónicos):
  `python tools/contract_check.py <raiz>/<TICKER>` — checks K, solo lectura.
  Sin PyYAML los de YAML quedan `[pendiente]`, nunca verdes.
- Historia financiera masiva de emisoras SEC: `python tools/xbrl_fetch.py <TICKER>`
```

En `CONTRIBUTING.md`, reemplazar

```
- **stdlib por default.** Una dependencia nueva necesita justificación explícita
  en el PR; `openpyxl` es la única aceptada hoy.
```

por

```
- **stdlib por default.** Una dependencia nueva necesita justificación explícita
  en el PR. Aceptadas hoy: `openpyxl` (obligatoria) y `pyyaml` (opcional: solo
  la usa `tools/contract_check.py`, que sin ella reporta sus checks de YAML
  como `[pendiente]` en vez de tronar).
```

y reemplazar

```
Requisitos: Python 3.10+ y `openpyxl` (única dependencia de terceros; el resto es
stdlib a propósito, para que los tools corran en cualquier sandbox).
```

por

```
Requisitos: Python 3.10+, `openpyxl` y, opcional, `pyyaml` (solo para validar
los contratos YAML; el resto es stdlib a propósito, para que los tools corran
en cualquier sandbox).
```

- [ ] **Step 4: Suite completa**

Run: `python -m unittest discover -s tests`
Expected: `OK`. Incluye `test_documented_k_checks_are_exactly_the_emitted_ones` y la prueba de consola ASCII sobre `tools/*.py`.

- [ ] **Step 5: Commit (CLI + contrato del repo + doctrina)**

```bash
git add tools/contract_check.py tests/test_contract_check.py tests/test_repo_contracts.py \
        skills/model-standards/references/integrity-checks.md commands AGENTS.md CONTRIBUTING.md
git commit -m "feat: CLI de contract_check, contrato del repo y doctrina de los checks K"
```

---

### Task 7: Verificación final y dogfood

**Files:** ninguno se modifica.

- [ ] **Step 1: Suite, sintaxis y demo**

Run: `python -m unittest discover -s tests`
Expected: `OK`.

Run: `python -c "import ast,glob;[ast.parse(open(f,encoding='utf-8').read()) for f in glob.glob('tools/*.py')];print('[ok] sintaxis')"`
Expected: `[ok] sintaxis`.

Run: `python tools/xlsx_builder.py demo ./scratch/demo.xlsx`
Expected: `Resumen F: 21 ok, 0 fallas, 0 pendientes, 0 avisos`.

- [ ] **Step 2: Modo sin PyYAML, de verdad**

Con un intérprete que tenga `openpyxl` y NO tenga PyYAML, correr el tool sobre una cobertura.
Expected: `[pendiente] K1`…`K4` con la instrucción de instalar, `K5` evaluado, exit 3 o 1. Si no hay tal intérprete, registrarlo: el modo degradado queda cubierto solo por la prueba con `mock`.

- [ ] **Step 3: Dogfood sobre las coberturas reales**

Correr `python tools/contract_check.py <cobertura>` sobre cada cobertura del workspace del analista.
Expected (medición del 2026-11-26): K1 falla en todas por `macro-view.yaml` (línea 258); K4 falla donde hay comps sin los campos obligatorios; K5 falla en la cobertura con CSV en formato ancho y avisa en la de etiquetas con texto extra; K2 avisa por campos propios y por la periodicidad deprecada. Registrar el resultado real. No corregir ningún contrato: es trabajo de cada cobertura, con decisión del analista.

- [ ] **Step 4: Suite parado en cada commit**

Para cada commit de la rama: `git checkout <sha>`, `python -m unittest discover -s tests`, y volver a la rama.
Expected: `OK` en todos.

---

## Desviaciones al ejecutar (2026-11-26)

El dogfood de la Task 7 sobre tres coberturas reales encontró dos falsos
positivos del propio validador. Se corrigieron con prueba primero, en un
commit aparte:

- **K2, `fiscal_year_end`.** Una cobertura declara `"09-lastSat"` (año fiscal
  de 52/53 semanas). No cabe en `MM-DD` y es legítimo: pasó de falla a aviso.
- **K4, campos de texto dentro de `balance_items` / `income_items`.** El plan
  validaba como dato `{value, source}` toda llave salvo `period`; los
  snapshots reales traen `currency: "USD"` ahí. Ahora solo se validan como
  dato las llaves que el template nombra (`COMP_SECTION_DATA`); las demás son
  campo desconocido, aviso. La prueba de contrato del repo ata esas llaves al
  template.

El hallazgo que sí era real y se conserva como falla: datos `observado` sin
`source` (el origen estaba en un comentario del YAML, no en el campo).

## Fuera de este plan

- Corregir los contratos de las coberturas existentes.
- Bump de versión, en su propio commit.
- Migrar los lectores por línea de `model_check.py` a PyYAML.
