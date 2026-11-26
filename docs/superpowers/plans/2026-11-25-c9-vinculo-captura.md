# C9 por código: vínculo fila ↔ capa de captura — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que `python tools/model_check.py <modelo.xlsx>` verifique por código que toda cifra histórica tecleada del modelo coincide con `model/inputs/canonical_*.csv`, con un umbral de materialidad para las cifras sin respaldo.

**Architecture:** El builder declara el vínculo fila ↔ `line_item` como un nombre definido de Excel con alcance de hoja (`CAPTURE.<line_item>`, escala en el comentario). `model_check.py` lee esos nombres, los CSV canónicos y el umbral del `issuer-profile.yaml`, y emite un `Finding` C9 con los estados que ya existen (`[ok]`, `[x]`, `[aviso]`). Solo lectura.

**Tech Stack:** Python 3.10+, stdlib (`csv`, `re`, `dataclasses`, `pathlib`), `openpyxl>=3.1`, `unittest`.

**Spec:** `docs/superpowers/specs/2026-11-25-c9-vinculo-captura-design.md`

**Rama:** `feat/c9-vinculo-captura` (sale de `main` con los PR 43 y 44 fusionados).

**Convenciones del repo que aplican a cada tarea** (CONTRIBUTING.md):

- stdout ASCII-only: sin acentos en `print()`, en ayudas de argparse ni en el `detail` de un `Finding`.
- Commits: Conventional Commits, sujeto en español, ASCII, 72 caracteres o menos.
- Correr pruebas: `python -m unittest discover -s tests` (toda la suite) o `python -m unittest tests.<modulo>.<Clase>` (una clase).
- Verificado antes de este plan, con Excel COM: el comentario de un nombre definido sobrevive a abrir y guardar en Excel, y el rango se desplaza al insertar una fila arriba (`$10:$10` pasó a `$11:$11`). Excel reescribe la referencia SIN comillas (`Operating!$11:$11`): el lector debe aceptar ambas formas.

---

## Mapa de archivos

| Archivo | Responsabilidad en este cambio |
|---|---|
| `tools/xlsx_builder.py` | `ModelStyler.link_capture` (escribe el vínculo); el demo vincula sus filas y escribe su CSV de captura |
| `tools/model_check.py` | Leer vínculos, CSV y umbral; `audit_capture` (C9); banderas `--inputs` y `--profile` |
| `tests/test_xlsx_builder_audit.py` | Pruebas de `link_capture` y del fixture del demo |
| `tests/test_model_check.py` | Pruebas de la capa de lectura, de C9 en los dos sentidos y del CLI |
| `tests/test_repo_contracts.py` | C9 entra a los checks C de escaneo |
| `templates/issuer-profile.yaml`, `templates/coverage-tree.md` | Contrato: umbral, líneas reservadas, escala, `derivado` |
| `skills/model-standards/references/integrity-checks.md`, `skills/model-standards/SKILL.md`, `skills/xlsx-building/SKILL.md`, `skills/statement-mapper/SKILL.md`, `commands/model-check.md`, `AGENTS.md`, `CONTRIBUTING.md` | Doctrina sincronizada |

---

### Task 1: `link_capture` en el builder

**Files:**
- Modify: `tools/xlsx_builder.py` (constantes de contrato cerca de `BUILDER_STAMP_KEY`; método nuevo después de `define_constant`)
- Test: `tests/test_xlsx_builder_audit.py` (clase `BuilderTests`)

- [ ] **Step 1: Escribir las pruebas que fallan**

Agregar al final de la clase `BuilderTests` en `tests/test_xlsx_builder_audit.py`:

```python
    def test_link_capture_writes_a_sheet_scoped_name_with_row_and_scale(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.link_capture(ws, 91, "bs_ap", scale=1_000_000)
        defined = ws.defined_names["CAPTURE.bs_ap"]
        self.assertEqual(defined.attr_text, "'Operating'!$91:$91")
        self.assertEqual(defined.comment, "scale=1000000.0")
        # Alcance de hoja: el nombre no existe a nivel libro.
        self.assertNotIn("CAPTURE.bs_ap", styler.wb.defined_names)

    def test_link_capture_survives_save_and_load(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.link_capture(ws, 91, "bs_ap", scale=1_000_000)
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "t.xlsx")
            styler.save(path)
            reloaded = load_workbook(path)["Operating"].defined_names["CAPTURE.bs_ap"]
        self.assertEqual(reloaded.attr_text, "'Operating'!$91:$91")
        self.assertEqual(reloaded.comment, "scale=1000000.0")

    def test_link_capture_allows_the_same_line_once_per_sheet(self) -> None:
        styler = B.ModelStyler()
        operating = styler.new_sheet("Operating")
        annual = styler.new_sheet("Annual")
        styler.link_capture(operating, 91, "bs_ap")
        styler.link_capture(annual, 40, "bs_ap")        # otra hoja: permitido
        with self.assertRaises(ValueError):
            styler.link_capture(operating, 92, "bs_ap")  # misma hoja: rechazado

    def test_link_capture_rejects_bad_line_item_and_scale(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        for line_item in ("bs ap", "1bs", "bs-ap", ""):
            with self.subTest(line_item=line_item), self.assertRaises(ValueError):
                styler.link_capture(ws, 10, line_item)
        for scale in (0, -1, True, float("nan")):
            with self.subTest(scale=scale), self.assertRaises(ValueError):
                styler.link_capture(ws, 10, "bs_ap", scale=scale)
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_xlsx_builder_audit.BuilderTests`
Expected: 4 errores `AttributeError: 'ModelStyler' object has no attribute 'link_capture'`.

- [ ] **Step 3: Implementar**

En `tools/xlsx_builder.py`, después de la línea `FONT_NAME = "Aptos Narrow"`:

```python
# Vinculo fila <-> capa de captura (check C9). Un nombre definido con alcance
# de hoja por linea: Excel ajusta su rango al insertar o mover filas, asi que
# el vinculo sobrevive a una edicion manual del libro.
CAPTURE_NAME_PREFIX = "CAPTURE."
# Lineas que C9 usa como referencia de materialidad: ventas y activos totales
# (los nombres que emite xbrl_fetch.CONCEPT_MAP).
CAPTURE_REFERENCE_LINES = ("is_ns_total", "bs_ta")
```

En la clase `ModelStyler`, inmediatamente después del método `define_constant`:

```python
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
```

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_xlsx_builder_audit`
Expected: `OK` (todas, incluidas las 4 nuevas).

- [ ] **Step 5: Commit**

```bash
git add tools/xlsx_builder.py tests/test_xlsx_builder_audit.py
git commit -m "feat: link_capture declara el vinculo fila - linea de captura"
```

---

### Task 2: El demo vincula sus filas y escribe su capa de captura

El demo debe pasar cada check (CONTRIBUTING 4.2). Para C9 necesita vínculos y un CSV.

**Files:**
- Modify: `tools/xlsx_builder.py` (imports de stdlib; función `_demo`; función nueva `_write_demo_capture`)
- Test: `tests/test_xlsx_builder_audit.py` (clase `AuditTests`)

- [ ] **Step 1: Escribir la prueba que falla**

En `tests/test_xlsx_builder_audit.py`, agregar `import csv` al bloque de imports de stdlib (orden alfabético, antes de `import io`), y al final de la clase `AuditTests`:

```python
    def test_demo_writes_its_capture_layer_next_to_the_book(self) -> None:
        path = Path(self.tmp.name) / "demo.xlsx"
        with contextlib.redirect_stdout(io.StringIO()):
            B._demo(str(path))
        capture = path.parent / "inputs" / "canonical_quarterly.csv"
        with capture.open(encoding="utf-8", newline="") as fh:
            rows = list(csv.DictReader(fh))
        names = set(load_workbook(path)["Operating"].defined_names)
        self.assertEqual(len(rows), 16 * 8)     # 16 filas observadas x 8 trimestres A
        self.assertEqual({row["tag"] for row in rows}, {"observado"})
        self.assertEqual(rows[0]["period"], "1Q2023")
        lines = {row["line_item"] for row in rows}
        for reserved in B.CAPTURE_REFERENCE_LINES:
            self.assertIn(reserved, lines)
            self.assertIn(B.CAPTURE_NAME_PREFIX + reserved, names)
```

- [ ] **Step 2: Correr y verificar que falla**

Run: `python -m unittest tests.test_xlsx_builder_audit.AuditTests.test_demo_writes_its_capture_layer_next_to_the_book`
Expected: `FileNotFoundError` sobre `inputs/canonical_quarterly.csv`.

- [ ] **Step 3: Implementar**

En `tools/xlsx_builder.py`, bloque de imports de stdlib (hoy `import sys` seguido de `from dataclasses import dataclass`): dejarlo así:

```python
import csv
import sys
from dataclasses import dataclass
from enum import Enum, auto
from pathlib import Path
from typing import Iterable, Optional
```

Dentro de `_demo`, inmediatamente después del diccionario `rows = {...}`:

```python
    # Linea de captura de cada fila OBSERVADA del demo (contrato del C9).
    capture_lines = {"rev": "is_ns_total", "cogs": "is_cogs_total", "da": "cf_da",
                     "interest": "is_interest", "tax": "is_tax", "cash": "bs_cash",
                     "ar": "bs_ar", "inv": "bs_inv", "ca": "bs_ca", "ta": "bs_ta",
                     "ap": "bs_ap", "cl": "bs_cl", "debt": "bs_debt",
                     "equity": "bs_equity", "re": "bs_re", "cfo": "cf_cfo"}
    captured: list[tuple[str, str, float]] = []
```

Dentro de la función anidada `line`, después de `label(op, r, text)` y antes del `for`:

```python
        if not callable(hist):
            # Fila con historico observado: se vincula y su serie va al CSV.
            styler.link_capture(op, r, capture_lines[key])
            captured.extend((capture_lines[key], quarters[i][:-1], float(hist[i]))
                            for i in range(n_act))
```

Al final de `_demo`, reemplazar

```python
    styler.save(path)
    print(f"[ok] demo escrito: {path}")
```

por

```python
    styler.save(path)
    _write_demo_capture(Path(path).parent / "inputs", captured)
    print(f"[ok] demo escrito: {path}")
```

Y agregar, inmediatamente antes de `def _demo(path: str) -> None:`:

```python
def _write_demo_capture(folder: Path, rows: list[tuple[str, str, float]]) -> None:
    """Capa de captura minima del demo (line_item, periodo, valor): el fixture
    contra el que corre C9. No es un artefacto de cobertura."""
    folder.mkdir(parents=True, exist_ok=True)
    with (folder / "canonical_quarterly.csv").open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(("line_item", "period", "value", "source_doc", "source_ref", "tag"))
        for line_item, period, value in rows:
            writer.writerow((line_item, period, value, "demo", "xlsx_builder._demo", "observado"))
```

- [ ] **Step 4: Correr y verificar que pasa, y que el demo sigue verde**

Run: `python -m unittest discover -s tests`
Expected: `OK`.

Run: `python tools/xlsx_builder.py demo ./scratch/demo.xlsx`
Expected: última línea `Resumen F: 20 ok, 0 fallas, 0 pendientes, 0 avisos`, y existe `scratch/inputs/canonical_quarterly.csv`.

- [ ] **Step 5: Commit**

```bash
git add tools/xlsx_builder.py tests/test_xlsx_builder_audit.py
git commit -m "feat: el demo vincula sus filas y escribe su capa de captura"
```

---

### Task 3: Leer vínculos, CSV y umbral en `model_check.py`

**Files:**
- Modify: `tools/model_check.py` (imports; `PeriodColumn`; `_period_axis`; sección nueva "C9 - capa de captura" antes de la sección "Reporte y CLI")
- Test: `tests/test_model_check.py` (clase nueva `CaptureLayerTests`)

- [ ] **Step 1: Escribir las pruebas que fallan**

En `tests/test_model_check.py`: agregar `import csv` al bloque de imports de stdlib (antes de `import io`), agregar `from openpyxl.workbook.defined_name import DefinedName` junto a los otros imports de openpyxl, y agregar esta clase inmediatamente antes de `class DriverMapTests`:

```python
class CaptureLayerTests(unittest.TestCase):
    """Lectura de la capa de captura: vinculos del libro, CSV y umbral."""

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = Path(tmp.name)

    def _csv(self, text: str) -> Path:
        path = self.dir / "canonical_quarterly.csv"
        path.write_text(text, encoding="utf-8")
        return path

    def test_links_are_read_in_builder_and_excel_reference_forms(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        styler.link_capture(ws, 91, "bs_ap", scale=1_000_000)
        # Excel reescribe la referencia sin comillas al guardar.
        ws.defined_names.add(DefinedName("CAPTURE.bs_ta", attr_text="Operating!$27:$27",
                                         comment="scale=1000000"))
        ws.defined_names.add(DefinedName("OTRO_NOMBRE", attr_text="Operating!$5:$5"))
        self.assertEqual(M.capture_links(ws),
                         {91: M.CaptureLink("bs_ap", 1_000_000.0),
                          27: M.CaptureLink("bs_ta", 1_000_000.0)})

    def test_link_without_scale_is_a_capture_error(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Operating")
        ws.defined_names.add(DefinedName("CAPTURE.bs_ap", attr_text="Operating!$91:$91"))
        with self.assertRaises(M.CaptureError) as caught:
            M.capture_links(ws)
        self.assertIn("CAPTURE.bs_ap", str(caught.exception))

    def test_capture_values_are_keyed_by_line_and_period(self) -> None:
        path = self._csv("line_item,period,value,source_doc,source_ref,tag\n"
                         "bs_ap,3Q2016,26318000000.0,10-Q,\"XBRL, filed 2016-07-27\",observado\n"
                         "is_ni,4Q2016,9014000000,10-K,FY menos 3Q,derivado\n"
                         "is_ns_total,1Q2027,95000000000,call,p.3,guidance\n")
        self.assertEqual(M.load_capture(path),
                         {("bs_ap", "3Q2016"): 26318000000.0,
                          ("is_ni", "4Q2016"): 9014000000.0})   # guidance no respalda

    def test_capture_without_a_required_column_is_a_capture_error(self) -> None:
        path = self._csv("line_item,period,value\nbs_ap,3Q2016,1\n")
        with self.assertRaises(M.CaptureError) as caught:
            M.load_capture(path)
        self.assertIn("faltan columnas tag", str(caught.exception))

    def test_non_numeric_capture_value_names_file_and_line(self) -> None:
        for bad in ("abc", "", "nan"):
            with self.subTest(value=bad):
                path = self._csv("line_item,period,value,tag\n"
                                 "bs_ap,3Q2016,1,observado\n"
                                 f"bs_ar,3Q2016,{bad},observado\n")
                with self.assertRaises(M.CaptureError) as caught:
                    M.load_capture(path)
                self.assertIn("canonical_quarterly.csv linea 3", str(caught.exception))

    def test_threshold_comes_from_the_profile_or_the_default(self) -> None:
        profile = self.dir / "issuer-profile.yaml"
        profile.write_text("materiality:\n"
                           "  critical_pct_target_price: 10\n"
                           "  capture_unlinked_pct: 2.5   # C9\n", encoding="utf-8")
        self.assertEqual(M.materiality_pct_from_profile(profile), (2.5, "perfil"))
        profile.write_text("materiality:\n  critical_pct_target_price: 10\n", encoding="utf-8")
        self.assertEqual(M.materiality_pct_from_profile(profile), (0.5, "default"))
        self.assertEqual(M.materiality_pct_from_profile(self.dir / "no-existe.yaml"),
                         (0.5, "default"))
        self.assertEqual(M.materiality_pct_from_profile(None), (0.5, "default"))

    def test_period_axis_reports_the_header_row(self) -> None:
        styler = B.ModelStyler()
        ws = styler.new_sheet("Model")
        styler.period_header(ws, 3, 3, B.PeriodHeader(2020, 2025, 2022))
        self.assertEqual({period.row for period in M._period_axis(ws)}, {3})
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_model_check.CaptureLayerTests`
Expected: 7 errores `AttributeError` (`capture_links`, `CaptureLink`, `CaptureError`, `load_capture`, `materiality_pct_from_profile`, `row`).

- [ ] **Step 3: Implementar**

En `tools/model_check.py`, bloque de imports de stdlib: dejarlo así:

```python
import argparse
import csv
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, Sequence
```

En `PeriodColumn`, agregar el campo `row` con default (las pruebas existentes lo construyen con 4 posicionales):

```python
@dataclass(frozen=True)
class PeriodColumn:
    """Una columna de periodo del header: anio fiscal y, si aplica, trimestre."""

    column: int
    actual: bool
    year: int
    quarter: int | None
    row: int = 0
```

En `_period_axis`, pasar la fila en las dos construcciones:

```python
            if cell.number_format in (B.NumFmt.YEAR_A.value, B.NumFmt.YEAR_E.value):
                if isinstance(value, (int, float)):
                    axis.append(PeriodColumn(cell.column, cell.number_format == B.NumFmt.YEAR_A.value,
                                             int(value), None, cell.row))
            elif isinstance(value, str):
                match = re.fullmatch(r"([1-4])Q(20\d\d)([AE])", value.strip())
                if match:
                    axis.append(PeriodColumn(cell.column, match.group(3) == "A",
                                             int(match.group(2)), int(match.group(1)), cell.row))
```

Agregar esta sección inmediatamente antes del comentario de sección `# Reporte y CLI`:

```python
# ---------------------------------------------------------------------------
# C9 - historico del xlsx = capa de captura
# ---------------------------------------------------------------------------

# CSV canonico segun el header de la hoja: True = trimestral, False = anual.
CAPTURE_FILES = {True: "canonical_quarterly.csv", False: "canonical_annual.csv"}
REQUIRED_CAPTURE_COLUMNS = ("line_item", "period", "value", "tag")
# Una fila guidance es una afirmacion de management: no respalda un historico.
NON_BACKING_TAG = "guidance"
DEFAULT_CAPTURE_MATERIALITY_PCT = 0.5
# El historico se escribe sin redondear: la tolerancia solo absorbe el float.
MATCH_RELATIVE_TOLERANCE = 1e-6

_CAPTURE_ROW = re.compile(r"\$?(\d+):\$?\d+$")
_CAPTURE_SCALE = re.compile(r"scale=([0-9.eE+-]+)")
_PROFILE_PCT = re.compile(r"^\s*capture_unlinked_pct:\s*([0-9]+(?:\.[0-9]+)?)", re.MULTILINE)


class CaptureError(Exception):
    """Vinculo o capa de captura ilegible: C9 lo reporta como falla, sin traceback."""


@dataclass(frozen=True)
class CaptureLink:
    """Vinculo de una fila del modelo con una linea de la capa de captura."""

    line_item: str
    scale: float


def capture_links(ws: Worksheet) -> dict[int, CaptureLink]:
    """Fila -> vinculo, leido de los nombres ``CAPTURE.<line_item>`` de la hoja."""
    links: dict[int, CaptureLink] = {}
    for name, defined in ws.defined_names.items():
        if not name.startswith(B.CAPTURE_NAME_PREFIX):
            continue
        row = _CAPTURE_ROW.search(defined.attr_text or "")
        scale = _CAPTURE_SCALE.search(defined.comment or "")
        if row is None or scale is None:
            raise CaptureError(f"{ws.title}: el nombre {name} no trae fila o escala "
                               "(se escribe con ModelStyler.link_capture)")
        links[int(row.group(1))] = CaptureLink(name[len(B.CAPTURE_NAME_PREFIX):],
                                               float(scale.group(1)))
    return links


def load_capture(path: Path) -> dict[tuple[str, str], float]:
    """(line_item, periodo) -> valor de un CSV canonico. Ignora filas guidance."""
    values: dict[tuple[str, str], float] = {}
    with path.open(encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        missing = [c for c in REQUIRED_CAPTURE_COLUMNS if c not in (reader.fieldnames or [])]
        if missing:
            raise CaptureError(f"{path.name}: faltan columnas {', '.join(missing)}")
        for number, row in enumerate(reader, start=2):
            if (row["tag"] or "").strip().lower() == NON_BACKING_TAG:
                continue
            try:
                value = float(row["value"])
            except (TypeError, ValueError):
                value = math.nan
            if math.isnan(value):
                raise CaptureError(
                    f"{path.name} linea {number}: value no numerico ({row['value']!r})")
            values[((row["line_item"] or "").strip(), (row["period"] or "").strip())] = value
    return values


def materiality_pct_from_profile(path: Path | None) -> tuple[float, str]:
    """(umbral en %, origen) de ``materiality.capture_unlinked_pct`` del perfil.

    Lectura por linea, como load_brand lee DESIGN.md: sin dependencia de YAML.
    Perfil o campo ausente = default (un perfil viejo sigue siendo valido).
    """
    if path is not None and path.is_file():
        found = _PROFILE_PCT.search(path.read_text(encoding="utf-8"))
        if found:
            return float(found.group(1)), "perfil"
    return DEFAULT_CAPTURE_MATERIALITY_PCT, "default"
```

Nota: `field` se usa en la Task 4; se importa aquí para no tocar el bloque de imports dos veces.

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_model_check`
Expected: `OK`.

- [ ] **Step 5: Commit**

```bash
git add tools/model_check.py tests/test_model_check.py
git commit -m "feat: model_check lee vinculos, capa de captura y umbral del perfil"
```

---

### Task 4: `audit_capture` — el check C9

**Files:**
- Modify: `tools/model_check.py` (continuar la sección "C9", después de `materiality_pct_from_profile`)
- Test: `tests/test_model_check.py` (helper en `_DemoCase`; clase nueva `CaptureCheckTests`)

Datos del demo que usan las pruebas: `Operating` tiene 8 trimestres históricos en las columnas C–J (`1Q2023`…`4Q2024`); la fila 12 es Ventas (`is_ns_total`, 100, 103, …); la fila 24 es Cuentas por cobrar (`bs_ar`, 30, 31, …); la fila 27 es Activo total (`bs_ta`, 300, 301, …). Con el umbral default, el corte de `1Q2023` es 0.5% × min(100, 300) = 0.5.

- [ ] **Step 1: Escribir las pruebas que fallan**

En `tests/test_model_check.py`, agregar `from copy import copy` al bloque de imports de stdlib (después de `import contextlib`).

En la clase `_DemoCase`, después del método `_mutated`:

```python
    def _beside_demo(self, mutate: Callable[[Workbook], None] | None = None) -> str:
        """Guarda el demo (mutado) JUNTO al original: comparte su inputs/."""
        wb = load_workbook(self.demo)
        if mutate is not None:
            mutate(wb)
        out = self.dir / "caso_2026-01-02_v1.xlsx"
        wb.save(out)
        return str(out)
```

Y, a nivel de módulo, después de la constante `FORECAST_CELL`:

```python
CAPTURE_COLUMNS = ("line_item", "period", "value", "source_doc", "source_ref", "tag")


def _unlinked_row(value: float) -> Callable[[Workbook], None]:
    """Mutacion: una fila sin vinculo con una cifra historica tecleada (1Q2023).
    Copia la fuente de una celda del demo para no disparar F2."""
    def mutate(wb: Workbook) -> None:
        ws = wb["Operating"]
        ws["B200"] = "Otros activos"
        ws["B200"].font = copy(ws["B12"].font)
        ws["C200"] = value
        ws["C200"].font = copy(ws["C12"].font)
    return mutate
```

Agregar esta clase inmediatamente antes de `class DriverMapTests`:

```python
class CaptureCheckTests(_DemoCase):
    """C9 en los dos sentidos: el demo lo pasa, cada libro roto lo falla."""

    def _c9(self, mutate: Callable[[Workbook], None] | None = None,
            **kwargs: object) -> B.Finding:
        (finding,) = M.audit_capture(self._beside_demo(mutate), **kwargs)  # type: ignore[arg-type]
        return finding

    def _rewrite_capture(self, transform: Callable[[list[dict[str, str]]],
                                                   list[dict[str, str]]]) -> None:
        path = self.dir / "inputs" / "canonical_quarterly.csv"
        with path.open(encoding="utf-8", newline="") as fh:
            rows = list(csv.DictReader(fh))
        with path.open("w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=CAPTURE_COLUMNS)
            writer.writeheader()
            writer.writerows(transform(rows))

    @staticmethod
    def _is(row: dict[str, str], line_item: str, period: str | None = None) -> bool:
        return row["line_item"] == line_item and period in (None, row["period"])

    def test_demo_matches_its_capture_layer(self) -> None:
        finding = self._c9()
        self.assertTrue(finding.ok, finding.detail)
        self.assertFalse(finding.warning)
        self.assertIn("128 cifras en 16 filas coinciden", finding.detail)
        self.assertIn("umbral 0.5%", finding.detail)

    def test_linked_figure_that_differs_from_the_capture_always_fails(self) -> None:
        def mutate(wb: Workbook) -> None:
            wb["Operating"]["D12"] = 999
        # Ni un umbral del 100% perdona una divergencia.
        finding = self._c9(mutate, materiality_pct=100.0)
        self.assertTrue(finding.blocking)
        self.assertIn("DIVERGENCIAS", finding.detail)
        self.assertIn("Operating!D12 (is_ns_total 2Q2023)", finding.detail)

    def test_material_figure_without_a_link_fails(self) -> None:
        finding = self._c9(_unlinked_row(50))
        self.assertTrue(finding.blocking)
        self.assertIn("SIN RESPALDO sobre el umbral", finding.detail)
        self.assertIn("Operating!fila 200 (Otros activos): 1 cifras sin vinculo", finding.detail)

    def test_immaterial_figure_without_a_link_is_a_warning(self) -> None:
        finding = self._c9(_unlinked_row(0.2))
        self.assertFalse(finding.ok)
        self.assertTrue(finding.warning)
        self.assertFalse(finding.blocking)
        self.assertIn("bajo el umbral", finding.detail)
        self.assertIn("Operating!fila 200", finding.detail)

    def test_threshold_decides_between_failure_and_warning(self) -> None:
        self.assertTrue(self._c9(_unlinked_row(50)).blocking)                       # 50 > 0.5
        self.assertTrue(self._c9(_unlinked_row(50), materiality_pct=60.0).warning)  # 50 <= 60

    def test_linked_row_missing_a_period_in_the_capture_fails(self) -> None:
        self._rewrite_capture(lambda rows: [r for r in rows if not self._is(r, "bs_ar", "3Q2023")])
        finding = self._c9()
        self.assertTrue(finding.blocking)
        self.assertIn("Operating!fila 24", finding.detail)
        self.assertIn("bs_ar sin periodo en la captura (desde 3Q2023)", finding.detail)

    def test_guidance_row_does_not_back_a_historical_figure(self) -> None:
        def retag(tag: str) -> Callable[[list[dict[str, str]]], list[dict[str, str]]]:
            return lambda rows: [dict(r, tag=tag) if self._is(r, "bs_ar", "3Q2023") else r
                                 for r in rows]
        self._rewrite_capture(retag("guidance"))
        self.assertTrue(self._c9().blocking)
        self._rewrite_capture(retag("derivado"))
        self.assertTrue(self._c9().ok)

    def test_scale_converts_model_units_to_source_units(self) -> None:
        def rescale(wb: Workbook) -> None:
            wb["Operating"].defined_names["CAPTURE.bs_cash"].comment = "scale=1000.0"
        # Solo cambia la escala del vinculo: el libro ya no coincide con el CSV.
        self.assertIn("DIVERGENCIAS", self._c9(rescale).detail)
        # El CSV pasa a la unidad de la fuente (x1000): vuelve a coincidir.
        self._rewrite_capture(lambda rows: [
            dict(r, value=str(float(r["value"]) * 1000)) if self._is(r, "bs_cash") else r
            for r in rows])
        self.assertTrue(self._c9(rescale).ok)

    def test_unbacked_figure_without_materiality_reference_fails(self) -> None:
        self._rewrite_capture(lambda rows: [r for r in rows if not self._is(r, "bs_ta")])
        finding = self._c9(materiality_pct=100.0)
        self.assertTrue(finding.blocking)
        self.assertIn("Operating!fila 27", finding.detail)
        self.assertIn("sin referencia de materialidad", finding.detail)

    def test_assumptions_section_is_out_of_scope(self) -> None:
        def rename_section(wb: Workbook) -> None:
            wb["Operating"]["B5"] = "Supuestos"
        # Sin la banda "Assumptions", sus constantes tecleadas (365 dias) cuentan.
        finding = self._c9(rename_section)
        self.assertTrue(finding.blocking)
        self.assertIn("Operating!fila 7", finding.detail)

    def test_missing_capture_folder_fails_unless_the_model_is_external(self) -> None:
        elsewhere = self._mutated(lambda wb: None)       # roto/ no tiene inputs/
        (strict,) = M.audit_capture(elsewhere)
        self.assertTrue(strict.blocking)
        self.assertIn("sin capa de captura", strict.detail)
        (external,) = M.audit_capture(elsewhere, external=True)
        self.assertTrue(external.ok)
        self.assertIn("n/a", external.detail)
        (pointed,) = M.audit_capture(elsewhere, inputs_dir=str(self.dir / "inputs"))
        self.assertTrue(pointed.ok, pointed.detail)

    def test_unreadable_capture_is_reported_as_a_failure_not_a_traceback(self) -> None:
        capture = self.dir / "inputs" / "canonical_quarterly.csv"
        capture.write_text("line_item,period,value\nbs_ap,1Q2023,20\n", encoding="utf-8")
        finding = self._c9()
        self.assertTrue(finding.blocking)
        self.assertIn("faltan columnas tag", finding.detail)

    def test_detail_is_ascii(self) -> None:
        for finding in (self._c9(), self._c9(_unlinked_row(50)), self._c9(_unlinked_row(0.2))):
            self.assertTrue(finding.detail.isascii(), finding.detail)
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_model_check.CaptureCheckTests`
Expected: 13 errores `AttributeError: module 'model_check' has no attribute 'audit_capture'`.

- [ ] **Step 3: Implementar**

En `tools/model_check.py`, después de `materiality_pct_from_profile`:

```python
@dataclass
class _UnbackedRow:
    """Fila con cifras historicas que la capa de captura no respalda."""

    label: str
    reason: str
    cells: int = 0
    material: int = 0
    unmeasurable: bool = False

    def describe(self, where: str) -> str:
        note = "; sin referencia de materialidad" if self.unmeasurable else ""
        return f"{where} ({self.label}): {self.cells} cifras {self.reason}{note}"


@dataclass
class _CaptureScan:
    """Acumulado del escaneo de C9 sobre todas las hojas de periodo."""

    figures: int = 0
    matched: int = 0
    matched_rows: int = 0
    mismatches: list[str] = field(default_factory=list)
    unbacked: dict[str, _UnbackedRow] = field(default_factory=dict)


def _period_key(period: PeriodColumn) -> str:
    """Llave de periodo del CSV canonico: '3Q2016' o 'FY2016'."""
    if period.quarter is not None:
        return f"{period.quarter}Q{period.year}"
    return f"FY{period.year}"


def _matches(value: float, scale: float, captured: float) -> bool:
    """valor_xlsx * scale = valor_csv, salvo error de punto flotante."""
    if captured == 0:
        return abs(value * scale) <= 1e-9 * scale
    return abs(value * scale - captured) <= MATCH_RELATIVE_TOLERANCE * abs(captured)


def _materiality_threshold(captured: dict[tuple[str, str], float], period: str,
                           materiality_pct: float) -> float | None:
    """X% de la MENOR entre ventas y activos totales del periodo, leidas del
    CSV (no del libro: asi C9 no depende de que este recalculado). None si
    falta cualquiera de las dos: la cifra no se puede medir."""
    sales, assets = (captured.get((line, period)) for line in B.CAPTURE_REFERENCE_LINES)
    if sales is None or assets is None:
        return None
    return materiality_pct / 100 * min(abs(sales), abs(assets))


def _money_scale(links: dict[int, CaptureLink]) -> float | None:
    """Escala monetaria de la hoja: la del vinculo de ventas (o activos).
    Una fila sin vinculo no tiene escala propia y hereda esta."""
    by_line = {link.line_item: link.scale for link in links.values()}
    for line in B.CAPTURE_REFERENCE_LINES:
        if line in by_line:
            return by_line[line]
    return None


def _scan_capture_sheet(ws: Worksheet, book: Book, folder: Path,
                        materiality_pct: float, scan: _CaptureScan) -> None:
    axis = [period for period in _period_axis(ws) if period.actual]
    if not axis:
        return
    source = folder / CAPTURE_FILES[any(period.quarter is not None for period in axis)]
    captured = load_capture(source) if source.is_file() else {}
    links = capture_links(ws)
    money = _money_scale(links)
    assumptions = _section_rows(ws, ASSUMPTIONS_TITLE) if ws is book.core_sheet else None
    header_row = max(period.row for period in axis)
    for r in range(header_row + 1, min(ws.max_row, B._MAX_SCAN_ROWS) + 1):
        if assumptions and assumptions[0] <= r <= assumptions[1]:
            continue  # Assumptions: supuestos del analista, no historico de filing
        link = links.get(r)
        row_matched = False
        for period in axis:
            cell = ws.cell(row=r, column=period.column)
            value = cell.value
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                continue  # vacia, texto o formula: no es una cifra tecleada
            scan.figures += 1
            key = _period_key(period)
            backing = captured.get((link.line_item, key)) if link is not None else None
            if link is not None and backing is not None:
                if _matches(value, link.scale, backing):
                    scan.matched += 1
                    row_matched = True
                else:
                    scan.mismatches.append(
                        f"{ws.title}!{cell.coordinate} ({link.line_item} {key}): "
                        f"libro {value * link.scale:,.2f} vs captura {backing:,.2f}")
                continue
            where = f"{ws.title}!fila {r}"
            entry = scan.unbacked.get(where)
            if entry is None:
                label = str(ws.cell(row=r, column=2).value or "").strip()[:30]
                reason = ("sin vinculo" if link is None else
                          f"de {link.line_item} sin periodo en la captura (desde {key})")
                entry = scan.unbacked[where] = _UnbackedRow(label, reason)
            entry.cells += 1
            threshold = _materiality_threshold(captured, key, materiality_pct)
            scale = link.scale if link is not None else money
            if threshold is None or scale is None:
                entry.unmeasurable = True
                entry.material += 1
            elif abs(value) * scale > threshold:
                entry.material += 1
        if row_matched:
            scan.matched_rows += 1


def audit_capture(path: str, inputs_dir: str | None = None,
                  materiality_pct: float = DEFAULT_CAPTURE_MATERIALITY_PCT,
                  external: bool = False, threshold_source: str = "default") -> list[Finding]:
    """C9: toda cifra historica tecleada del modelo coincide con la capa de
    captura (model/inputs/canonical_*.csv). Lectura pura.

    Divergencia en una fila vinculada = falla siempre. Cifra sin respaldo
    (fila sin vinculo, o periodo ausente en el CSV) = falla si es material,
    [aviso] si no. ``external``: modelo que no construyo el plugin; sin capa
    de captura el check no aplica.
    """
    name = "C9 historico = capa de captura"
    folder = Path(inputs_dir) if inputs_dir else Path(path).parent / "inputs"
    if not folder.is_dir():
        if external:
            return [Finding(name, True, "n/a (modelo externo sin capa de captura)")]
        return [Finding(name, False,
                        f"sin capa de captura: no existe {folder} (usa --inputs)")]
    book = load_book(path)
    scan = _CaptureScan()
    try:
        for ws in book.formulas.worksheets:
            if ws.title in PERIOD_GRID_SHEETS:
                _scan_capture_sheet(ws, book, folder, materiality_pct, scan)
    except CaptureError as exc:
        return [Finding(name, False, str(exc))]
    rule = f"umbral {materiality_pct:g}% de min(ventas, activos), {threshold_source}"
    material = [row.describe(where) for where, row in scan.unbacked.items() if row.material]
    minor = [row.describe(where) for where, row in scan.unbacked.items() if not row.material]
    if scan.mismatches or material:
        parts: list[str] = []
        if scan.mismatches:
            parts.append("DIVERGENCIAS: " + _examples(scan.mismatches))
        if material:
            parts.append("SIN RESPALDO sobre el umbral: " + _examples(material))
        return [Finding(name, False, "; ".join(parts) + f" ({rule})")]
    if minor:
        return [Finding(name, False,
                        f"{len(minor)} filas con cifras sin respaldo bajo el umbral ({rule}): "
                        + _examples(minor), warning=True)]
    if not scan.figures:
        return [Finding(name, True, "n/a (sin cifras historicas tecleadas)")]
    return [Finding(name, True,
                    f"{scan.matched} cifras en {scan.matched_rows} filas coinciden "
                    f"con la captura ({rule})")]
```

- [ ] **Step 4: Correr y verificar que pasan**

Run: `python -m unittest tests.test_model_check`
Expected: `OK`.

- [ ] **Step 5: Commit**

```bash
git add tools/model_check.py tests/test_model_check.py
git commit -m "feat: check C9 por codigo con umbral de materialidad"
```

---

### Task 5: C9 en el CLI y en el contrato del repo

**Files:**
- Modify: `tools/model_check.py` (docstring del módulo; `main`)
- Modify: `tests/test_repo_contracts.py` (`SCANNED_C_CHECKS`)
- Test: `tests/test_model_check.py` (clase `CliTests`)

- [ ] **Step 1: Escribir las pruebas que fallan**

En `tests/test_model_check.py`, al final de la clase `CliTests`:

```python
    def test_c9_is_part_of_the_single_report(self) -> None:
        _, out = self._run(str(self.demo))
        self.assertIn("[ok] C9 historico = capa de captura", out)

    def test_inputs_flag_points_at_a_capture_layer_elsewhere(self) -> None:
        elsewhere = self._mutated(lambda wb: None)       # roto/ no tiene inputs/
        code, out = self._run(elsewhere)
        self.assertEqual(code, 1)
        self.assertIn("[x] C9", out)
        _, out = self._run(elsewhere, "--inputs", str(self.dir / "inputs"))
        self.assertIn("[ok] C9", out)

    def test_profile_threshold_turns_a_material_row_into_a_warning(self) -> None:
        book = self._beside_demo(_unlinked_row(50))
        code, out = self._run(book)                       # sin perfil: default 0.5%
        self.assertEqual(code, 1)
        self.assertIn("[x] C9", out)
        profile = self.dir / "issuer-profile.yaml"
        profile.write_text("materiality:\n  capture_unlinked_pct: 60\n", encoding="utf-8")
        code, out = self._run(book, "--profile", str(profile))
        # 3 y no 1: el aviso no bloquea; queda el pendiente de recalculo (S10).
        self.assertEqual(code, 3)
        self.assertIn("[aviso] C9", out)
        self.assertIn("umbral 60%", out)
        self.assertIn("perfil", out)
```

En `tests/test_repo_contracts.py`, reemplazar

```python
# Checks C que son escaneo por codigo (el resto son formulas de la tab Checks).
# C9 esta documentado como escaneo pero aun no tiene implementacion unica.
SCANNED_C_CHECKS = {"C10"}
```

por

```python
# Checks C que son escaneo por codigo (el resto son formulas de la tab Checks).
SCANNED_C_CHECKS = {"C9", "C10"}
```

- [ ] **Step 2: Correr y verificar que fallan**

Run: `python -m unittest tests.test_model_check.CliTests tests.test_repo_contracts`
Expected: fallan las 3 pruebas nuevas de `CliTests` (el reporte no trae `C9`; `--inputs` y `--profile` no existen: `SystemExit: 2`). `test_repo_contracts` pasa (el Task 4 ya emite `"C9 `).

- [ ] **Step 3: Implementar**

En `tools/model_check.py`, reemplazar el docstring del módulo, desde la línea `Uso:` hasta la línea que termina en `no se reportan aqui.`, por:

```python
Uso:
    python tools/model_check.py <modelo.xlsx> [--driver-map driver-map.md]
                                [--brand brand/DESIGN.md] [--external]
                                [--inputs DIR] [--profile issuer-profile.yaml]

Corre S (estructurales) -> C por escaneo (C9, C10) -> F (formato, via
``xlsx_builder.audit_format``). Los C de formula viva (tab Checks) y los D
(doctrina) no son escaneo de codigo y no se reportan aqui.

C9 busca la capa de captura en ``inputs/`` junto al modelo y el umbral de
materialidad en el ``issuer-profile.yaml`` de la carpeta padre de ``model/``
(arbol de cobertura). ``--inputs`` y ``--profile`` cambian esas rutas.
```

En `main`, reemplazar desde `parser = argparse.ArgumentParser(` hasta el final de la función por:

```python
    parser = argparse.ArgumentParser(
        description="Checks S, C9, C10 y F del modelo por codigo (solo lectura).")
    parser.add_argument("model", help="ruta del modelo .xlsx")
    parser.add_argument("--driver-map", help="driver-map.md: exige un bloque Sch por schedule (S1)")
    parser.add_argument("--brand", help="brand/DESIGN.md: colores de marca aceptados por F4")
    parser.add_argument("--external", action="store_true",
                        help="modelo que no construyo el plugin: la falta de sello (F10) es aviso")
    parser.add_argument("--inputs",
                        help="carpeta de la capa de captura para C9 (default: inputs/ junto al modelo)")
    parser.add_argument("--profile",
                        help="issuer-profile.yaml con el umbral de C9 (default: el de la cobertura)")
    args = parser.parse_args(list(argv[1:]))
    model = Path(args.model)
    if not model.is_file():
        print(f"[x] no existe el modelo: {args.model}")
        return 2
    schedules = schedules_from_driver_map(args.driver_map) if args.driver_map else None
    brand = B.load_brand(args.brand) if args.brand else None
    profile = (Path(args.profile) if args.profile
               else model.resolve().parent.parent / "issuer-profile.yaml")
    materiality_pct, threshold_source = materiality_pct_from_profile(profile)
    return print_report([
        ("Estructurales (S)", audit_structure(args.model, schedules)),
        ("Contables por escaneo (C)", [
            *audit_capture(args.model, inputs_dir=args.inputs, materiality_pct=materiality_pct,
                           external=args.external, threshold_source=threshold_source),
            *audit_cash_roll_formulas(args.model),
        ]),
        ("Formato (F)", B.audit_format(args.model, brand=brand, external=args.external)),
    ])
```

- [ ] **Step 4: Correr toda la suite**

Run: `python -m unittest discover -s tests`
Expected: `OK`. Incluye `tests/test_cli_ascii.py` (consola ASCII) y `tests/test_repo_contracts.py` (C9 documentado = emitido).

- [ ] **Step 5: Commit**

```bash
git add tools/model_check.py tests/test_model_check.py tests/test_repo_contracts.py
git commit -m "feat: C9 en la corrida unica de model_check (--inputs, --profile)"
```

---

### Task 6: Contratos y doctrina

Sin código. Cada reemplazo es texto exacto: si el texto "viejo" no aparece una sola vez, detente y revisa.

**Files:**
- Modify: `templates/issuer-profile.yaml`, `templates/coverage-tree.md`
- Modify: `skills/model-standards/references/integrity-checks.md`, `skills/model-standards/SKILL.md`, `skills/xlsx-building/SKILL.md`, `skills/statement-mapper/SKILL.md`
- Modify: `commands/model-check.md`, `AGENTS.md`, `CONTRIBUTING.md`

- [ ] **Step 1: `templates/issuer-profile.yaml` — el umbral**

Reemplazar

```yaml
  relevant_pct_target_price: 3    # entre relevant y critical -> relevante
                                  # debajo -> cosmético; el analista puede reclasificar
```

por

```yaml
  relevant_pct_target_price: 3    # entre relevant y critical -> relevante
                                  # debajo -> cosmético; el analista puede reclasificar
  capture_unlinked_pct: 0.5       # check C9: cifra histórica SIN respaldo en la capa de
                                  # captura que pese más de X% de min(ventas, activos
                                  # totales) del periodo -> falla; debajo -> aviso.
                                  # Opcional: ausente = 0.5
```

- [ ] **Step 2: `templates/coverage-tree.md` — contrato de la capa de captura**

Reemplazar

```
documento, página/nota, fecha). Columnas mínimas de los CSV: `line_item`,
`period`, `value`, `source_doc`, `source_ref`, `tag` (observado/guidance).
El histórico del xlsx se puebla DESDE los canónicos — check C9 verifica que
coincidan. Regenerables desde los JSON; los JSON, desde los filings. Nada se
teclea dos veces.
```

por

```
documento, página/nota, fecha). Columnas mínimas de los CSV: `line_item`,
`period`, `value`, `source_doc`, `source_ref`, `tag` (observado / derivado /
guidance). `period` es `3Q2016` en el trimestral y `FY2016` en el anual.
El histórico del xlsx se puebla DESDE los canónicos — check C9 verifica por
código que coincidan. Regenerables desde los JSON; los JSON, desde los
filings. Nada se teclea dos veces.

Reglas que C9 exige de esta capa:

- **`derivado`**: un subtotal o residual que el modelo necesita como cifra
  (el 4Q = FY − 3Q; "otros activos" = total − partidas) entra al CSV con
  `tag = derivado` y su fórmula en `source_ref`. No se calcula en el script
  de build: una cifra que no está en el CSV no tiene respaldo.
- **`guidance`** no respalda una cifra histórica.
- **Líneas reservadas**: `is_ns_total` (ventas) y `bs_ta` (activos totales).
  Son la referencia de materialidad de C9; sin ellas en el CSV y vinculadas
  en el modelo, toda cifra sin respaldo falla.
- **Unidades**: el CSV guarda la unidad de la fuente. El modelo declara la
  suya en cada vínculo (`link_capture(..., scale=1_000_000)`):
  `valor_csv = valor_xlsx × scale`. El histórico se escribe sin redondear;
  el formato de celda redondea la vista.
```

- [ ] **Step 3: `integrity-checks.md` — fila C9 e introducción**

Reemplazar la fila completa que empieza con `| C9 | Histórico del xlsx = capa de captura | Escaneo por código: celdas históricas` (una sola línea, termina en `de cada cobertura |`) por esta línea:

```
| C9 | Histórico del xlsx = capa de captura | Escaneo por código (`tools/model_check.py`). **Qué mira:** toda cifra histórica TECLEADA (número, no fórmula) en columnas `A` de `Operating`/`Annual`/`Model`, fuera de la sección Assumptions. **Contra qué:** `model/inputs/canonical_quarterly.csv` (hoja trimestral) o `canonical_annual.csv` (hoja anual). **Vínculo:** el builder lo declara con `link_capture(ws, fila, line_item, scale)` — nombre definido `CAPTURE.<line_item>` con alcance de hoja y la escala en su comentario (`valor_csv = valor_xlsx × scale`); Excel ajusta el rango si alguien inserta filas. **Divergencia** en una fila vinculada = falla SIEMPRE, sin umbral (tolerancia relativa 1e-6: el histórico se escribe sin redondear). **Sin respaldo** (fila sin vínculo, o periodo ausente en el CSV; `tag = guidance` no respalda) = falla si la cifra pesa más de `materiality.capture_unlinked_pct` (issuer-profile, default 0.5%) de la MENOR entre ventas (`is_ns_total`) y activos totales (`bs_ta`) del periodo; bajo el umbral = `[aviso]`, se lista siempre. Sin referencia para medir = falla. Sin carpeta de captura = falla (con `--external`: n/a) |
```

Reemplazar

```
Los checks S y el C10 corren con `python tools/model_check.py <modelo.xlsx>
[--driver-map driver-map.md] [--brand brand/DESIGN.md]` — misma doctrina y
```

por

```
Los checks S, C9 y C10 corren con `python tools/model_check.py <modelo.xlsx>
[--driver-map driver-map.md] [--brand brand/DESIGN.md]` — misma doctrina y
```

- [ ] **Step 4: `commands/model-check.md`**

Reemplazar

```
en orden S (estructurales) → C (contables) → D (contenido) → F (formato). Los S,
el C10 y los F corren en UNA corrida con `python tools/model_check.py
```

por

```
en orden S (estructurales) → C (contables) → D (contenido) → F (formato). Los S,
C9, C10 y los F corren en UNA corrida con `python tools/model_check.py
```

Reemplazar

```
cobertura. Los C de fórmula viva se leen de la tab `Checks`; C9 y los D los
evalúa el agente con la regla escrita. Solo lectura: este comando NUNCA
modifica el modelo.
```

por

```
cobertura. Los C de fórmula viva se leen de la tab `Checks`; los D los
evalúa el agente con la regla escrita. C9 toma la capa de captura de
`model/inputs/` junto al modelo y el umbral del `issuer-profile.yaml` de la
cobertura (`--inputs` / `--profile` cambian esas rutas). Solo lectura: este
comando NUNCA modifica el modelo.
```

- [ ] **Step 5: skills**

En `skills/model-standards/SKILL.md`, reemplazar

```
6. Corre TODOS los checks de `references/integrity-checks.md`: los S, el C10 y
   los F por código (`python tools/model_check.py <modelo> --driver-map
```

por

```
6. Corre TODOS los checks de `references/integrity-checks.md`: los S, C9, C10 y
   los F por código (`python tools/model_check.py <modelo> --driver-map
```

En `skills/xlsx-building/SKILL.md`, reemplazar

```
   `series_row` — histórico calculado y forecast en LA MISMA fila (check F11),
   jamás filas hist/forecast separadas ni columnas históricas vacías.
```

por

```
   `series_row` — histórico calculado y forecast en LA MISMA fila (check F11),
   jamás filas hist/forecast separadas ni columnas históricas vacías.
   Toda fila con histórico OBSERVADO se vincula a su línea de la capa de
   captura: `styler.link_capture(ws, fila, "bs_ap", scale=1_000_000)`
   (`valor_csv = valor_xlsx × scale`; el histórico va sin redondear). Sin
   vínculo, C9 trata sus cifras como no respaldadas.
```

En `skills/statement-mapper/SKILL.md`, reemplazar

```
  `templates/coverage-tree.md` §model/inputs). El histórico del xlsx se puebla
  desde estos CSV — jamás re-tecleado (check C9 verifica coincidencia).
```

por

```
  `templates/coverage-tree.md` §model/inputs). El histórico del xlsx se puebla
  desde estos CSV — jamás re-tecleado (check C9 lo verifica por código, cifra
  por cifra). Un subtotal o residual que el modelo necesite como cifra entra
  aquí con `tag = derivado` y su fórmula en `source_ref`: si no está en el
  CSV, C9 lo reporta sin respaldo.
```

- [ ] **Step 6: `AGENTS.md` y `CONTRIBUTING.md`**

En `AGENTS.md`, reemplazar

```
  checks S + C10 + F en una corrida: `python tools/model_check.py <modelo>` (nunca
```

por

```
  checks S + C9 + C10 + F en una corrida: `python tools/model_check.py <modelo>` (nunca
```

En `CONTRIBUTING.md`, reemplazar

```
| **C** (contables) | fórmulas en la tab `Checks` del modelo; los de escaneo (C10) en `tools/model_check.py` | la fórmula la escribe el builder, no el analista |
```

por

```
| **C** (contables) | fórmulas en la tab `Checks` del modelo; los de escaneo (C9, C10) en `tools/model_check.py` | la fórmula la escribe el builder, no el analista |
```

- [ ] **Step 7: Verificar que la doctrina no dejó referencias colgadas**

Run: `python -m unittest tests.test_repo_contracts`
Expected: `OK` (rutas citadas existen; C9 documentado y emitido).

Run: `grep -rn "Sin implementación única" skills commands templates AGENTS.md CONTRIBUTING.md README.md`
Expected: sin resultados.

- [ ] **Step 8: Commit**

```bash
git add templates skills commands AGENTS.md CONTRIBUTING.md
git commit -m "docs: contrato y doctrina del C9 por codigo"
```

---

### Task 7: Verificación final

**Files:** ninguno se modifica.

- [ ] **Step 1: Suite completa**

Run: `python -m unittest discover -s tests`
Expected: `OK`.

- [ ] **Step 2: Sintaxis y demo, como los corre la CI**

Run: `python -c "import ast,glob;[ast.parse(open(f,encoding='utf-8').read()) for f in glob.glob('tools/*.py')];print('[ok] sintaxis')"`
Expected: `[ok] sintaxis`.

Run: `python tools/xlsx_builder.py demo ./scratch/demo_2026-01-01_v1.xlsx`
Expected: `Resumen F: 20 ok, 0 fallas, 0 pendientes, 0 avisos`.

- [ ] **Step 3: Corrida única sobre el demo**

Run: `python tools/model_check.py ./scratch/demo_2026-01-01_v1.xlsx`
Expected: `[ok] C9 historico = capa de captura: 128 cifras en 16 filas coinciden con la captura (umbral 0.5% de min(ventas, activos), default)`; exit 3 por `[pendiente] S10` (el demo sale de openpyxl sin valores calculados).

- [ ] **Step 4: Dogfood — el modelo vigente de AAPL debe FALLAR C9**

AAPL v4 se construyó antes de este cambio y no trae vínculos. Correr `model_check.py` sobre el modelo vigente de la cobertura de AAPL con su `--driver-map`.
Expected: `[x] C9` con "SIN RESPALDO sobre el umbral" listando filas de `Operating`; los demás checks sin cambio (31 ok antes de C9). Es el resultado correcto: confirma que un modelo sin vínculos no pasa. La migración es la sección D6 del spec, con gate del analista, y no forma parte de este plan.

- [ ] **Step 5: Python 3.10 (compatibilidad declarada)**

Si hay un intérprete 3.10–3.13 con `openpyxl`: correr la suite ahí. Si no lo hay, registrar "no verificado en 3.10–3.13; lo cubre la CI" en la descripción del PR. No instalar paquetes en un intérprete del usuario sin preguntar.

---

## Desviaciones al ejecutar (2026-11-25)

- **Rebase a media ejecución.** Entre las tareas 4 y 5, `main` recibió el PR 45
  (check F21, `recalc_validate.py`, `--recalc`). La rama se rebasó ahí; los dos
  conflictos fueron de imports.
- **Tarea 4 absorbió el cambio de `SCANNED_C_CHECKS`.** El plan lo ponía en la
  tarea 5, pero la prueba de contrato falla en cuanto C9 se emite: va en el
  mismo commit, para que cada commit quede en verde.
- **Tarea 5: `CaptureOptions`.** Con `--recalc`, `model_check.py` audita una
  copia en un directorio temporal. C9 resuelve `inputs/` y el perfil sobre la
  ruta ORIGINAL del modelo y los pasa a `audit_groups(..., capture=...)`. Sin
  esto, todo modelo recalculado habría fallado por "sin capa de captura".
- **Tarea 6.** Dos anclas de texto cambiaron con el PR 45
  (`commands/model-check.md` y `skills/model-standards/SKILL.md`); el
  contenido nuevo es el mismo.

## Fuera de este plan

- **Migración de AAPL a v5** (spec D6): se hace en el workspace de la cobertura, después de fusionar, con decisión del analista fila por fila para las 4 filas de "Otros" y las 3 semillas de apertura.
- **Bump de versión** (0.6.0 → 0.7.0): en su propio commit, cuando el analista lo pida.
- Escribir el resultado de C9 en la tab `Checks` al construir.
