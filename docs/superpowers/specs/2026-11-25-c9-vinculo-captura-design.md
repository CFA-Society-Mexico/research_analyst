# Diseño — C9 por código: vínculo fila ↔ capa de captura y umbral de materialidad

Fecha: 2026-11-25 · Aprobado por Alan (brainstorming P1 = cobertura total con
umbral · P2 = materialidad · P3 = nombres definidos).
Origen: auditoría del repo. C9 era el último check declarado "escaneo por
código" sin implementación única (ver `integrity-checks.md`, fila C9).

## El problema

C9 dice: el histórico del xlsx coincide con `model/inputs/canonical_*.csv`.
Hoy no se puede verificar desde fuera del script que construyó el libro:

- El xlsx no guarda qué `line_item` del CSV corresponde a qué fila. Ese mapa
  vive solo dentro del `build_model.py` de cada cobertura.
- En el dogfood de AAPL, C9 lo corre el mismo script que escribe el libro,
  sobre 9 líneas elegidas a mano, y estampa `OK` con fecha en la tab `Checks`.
- Medición sobre AAPL v4 (`Operating`, 72 filas con cifras históricas
  tecleadas): 51 coinciden con la captura y solo les falta declarar el
  vínculo; 14 son supuestos de la sección Assumptions; 4 son residuales
  ("Otros activos/pasivos") calculados en Python que no están en el CSV; 3 son
  saldos de apertura (una celda semilla cada uno).

Una cifra histórica mal es uno de los dos bugs que llegan al reporte
publicado (CONTRIBUTING §9). C9 es el check que lo vigila.

## Decisiones de brainstorming

| # | Decisión | Elección |
|---|---|---|
| P1 | Qué garantiza C9 | **Cobertura total**: toda cifra histórica tecleada en `Operating`/`Annual`/`Model`, fuera de Assumptions, debe estar respaldada por la captura. Con flexibilidad vía umbral |
| P2 | Qué mide el umbral | **Materialidad** de la cifra sin respaldo, no conteo. El conteo no ve el tamaño: en AAPL, 85% de filas vinculadas dejaba fuera justo los residuales de decenas de miles de millones |
| P3 | Dónde vive el vínculo | **Nombres definidos de Excel con alcance de hoja**. Excel ajusta el rango cuando alguien inserta o mueve filas; el libro se describe a sí mismo; no agrega hojas |

Descartadas: hoja `Capture_Map` (la fila escrita como número se desfasa al
insertar filas, y agrega una tab) y archivo lateral `row-map.csv` (dos
archivos que se desincronizan por versión del modelo).

## D1 — El vínculo (lo escribe el builder)

Método nuevo en `ModelStyler` (`tools/xlsx_builder.py`):

```python
styler.link_capture(ws, row, "bs_ap", scale=1_000_000)
```

- Crea el nombre definido `CAPTURE.<line_item>` con **alcance de la hoja**
  `ws`, apuntando a la fila completa: `'Operating'!$91:$91`. El mismo
  `line_item` puede vincularse una vez en cada hoja (`Operating` y `Annual`).
- Guarda la escala en el comentario del nombre: `scale=1000000`.
  Regla: `valor_csv = valor_xlsx × scale`. El CSV guarda la unidad de la
  fuente; el modelo, la unidad de trabajo (USD millones en AAPL).
- `line_item` debe cumplir `[A-Za-z][A-Za-z0-9_]*`. `scale` debe ser > 0.
- Falla al construir (`ValueError`) si el `line_item` ya está vinculado en esa
  hoja, o si el nombre o la escala son inválidos. Una fila, una línea.
- **Líneas reservadas** como referencia de materialidad: `is_ns_total`
  (ventas) y `bs_ta` (activos totales). Son los nombres que ya emite
  `xbrl_fetch.py` (`CONCEPT_MAP`).

Regla de unidades que acompaña al vínculo: **el histórico se escribe sin
redondear**; el formato de celda redondea la vista. Es lo que AAPL ya hace.

## D2 — El check (`tools/model_check.py`)

Función pública nueva:

```python
audit_capture(path, inputs_dir=None, materiality_pct=0.5, external=False,
              threshold_source="default") -> list[Finding]
```

`threshold_source` solo alimenta el detalle del check ("perfil" o "default").

Devuelve un solo `Finding`: `C9 historico = capa de captura`.

### Qué cuenta como cifra histórica

Celda con número tecleado (`int`/`float`, no fórmula, no booleano) que cumple
todo esto:

1. Está en una hoja de `PERIOD_GRID_SHEETS` (`Operating`, `Annual`, `Model`).
2. Está en una columna histórica del header (`A`), en una fila posterior a la
   fila del header.
3. No está en la sección Assumptions de la hoja núcleo (`Operating`, o
   `Model` en modo anual).

Las fórmulas no son cifras tecleadas y no entran: las gobiernan S5 y los F.

### Contra qué se compara

- Hoja con header trimestral → `canonical_quarterly.csv`, llave
  `(line_item, "3Q2016")`.
- Hoja con header anual → `canonical_annual.csv`, llave `(line_item, "FY2016")`.
- La carpeta es `inputs/` junto al xlsx (árbol de cobertura:
  `<TICKER>/model/<modelo>.xlsx` y `<TICKER>/model/inputs/`). `--inputs` la
  cambia.
- Columnas leídas: `line_item`, `period`, `value`, `tag`. Las filas con
  `tag = guidance` no respaldan una cifra histórica y se ignoran; `observado`
  y `derivado` sí.

### Veredicto por cifra

| Caso | Clasificación |
|---|---|
| Fila vinculada, el CSV trae esa `(line_item, period)` y coincide | coincide |
| Fila vinculada, el CSV la trae y NO coincide | **divergencia** — falla siempre, sin umbral |
| Fila vinculada, el CSV no trae ese periodo | sin respaldo |
| Fila sin vínculo | sin respaldo |

Coincide = `|valor_xlsx × scale − valor_csv| ≤ 1e-6 × |valor_csv|`
(con `valor_csv = 0`: diferencia absoluta ≤ `1e-9 × scale`).

### Materialidad de una cifra sin respaldo

Una cifra sin respaldo es **material** si

```
|valor_xlsx| × escala_monetaria  >  (materiality_pct / 100) × min(|ventas_p|, |activos_p|)
```

- `ventas_p` y `activos_p` son `is_ns_total` y `bs_ta` del periodo `p`,
  **leídos del CSV** de esa hoja, no del libro: así el check no depende de
  que el libro esté recalculado.
- `escala_monetaria` es la escala del vínculo `CAPTURE.is_ns_total` de la hoja
  (o la de `CAPTURE.bs_ta` si falta el primero). Una fila sin vínculo no tiene
  escala propia; se asume la unidad monetaria del modelo. Una fila vinculada
  a la que le falta un periodo en el CSV usa la escala de su propio vínculo.
- Se usa la **menor** de las dos referencias porque una cifra sin vínculo no
  tiene metadatos: el check no puede saber si es de resultados o de balance.
  Es la lectura estricta; el umbral compensa.
- Si falta cualquiera de las dos referencias para el periodo, o la hoja no
  tiene ningún vínculo de referencia, la cifra **no se puede medir y cuenta
  como material**. El detalle lo dice.

### Veredicto del check

| Situación | Resultado |
|---|---|
| Alguna divergencia, o alguna cifra sin respaldo material | `[x]`. El detalle lista primero divergencias, luego materiales |
| Ninguna de las anteriores, pero hay cifras sin respaldo bajo el umbral | `[aviso]` (`Finding.warning`). Se listan siempre; no cambia el exit code |
| Todo coincide | `[ok]` con el conteo de cifras y filas comparadas |
| No existe la carpeta de captura, sin `--external` | `[x]`: "sin capa de captura" |
| No existe la carpeta de captura, con `--external` | `[ok]` n/a |
| Hoja sin columnas históricas o libro sin hojas de periodo | `[ok]` n/a |

Si la carpeta existe pero falta el CSV de una hoja (p. ej. no hay
`canonical_annual.csv`), las cifras de esa hoja quedan sin respaldo y sin
referencia de materialidad: fallan.

Errores de entrada (CSV sin una columna requerida, `value` no numérico) se
reportan como `[x]` con archivo y línea, sin traceback.

## D3 — El umbral

- Campo nuevo y opcional en `templates/issuer-profile.yaml`, dentro del
  bloque `materiality` que ya existe:

  ```yaml
  materiality:
    capture_unlinked_pct: 0.5   # C9: cifra historica sin respaldo en la captura
                                # que pese mas de X% de min(ventas, activos) -> falla
  ```

- Default 0.5 cuando el campo o el archivo no existen: un perfil de una
  cobertura existente sigue siendo válido (CONTRIBUTING 4.4).
- `model_check.py` lo lee con una expresión regular de una línea, como
  `load_brand` lee `DESIGN.md`. No agrega dependencia de YAML.
- Ubicación del perfil: `issuer-profile.yaml` en la carpeta padre de `model/`
  (árbol de cobertura). `--profile` la cambia.
- El detalle de C9 nombra siempre el umbral usado y de dónde salió (perfil o
  default).

## D4 — CLI

```
python tools/model_check.py <modelo.xlsx> [--driver-map ...] [--brand ...]
                            [--external] [--inputs DIR] [--profile YAML]
```

`C9` se emite en el grupo "Contables por escaneo (C)", junto a `C10`.

## D5 — Demo, pruebas y doctrina

- **Demo.** `_demo` vincula sus filas observadas (`scale=1`), incluidas
  `is_ns_total` y `bs_ta`, y escribe `inputs/canonical_quarterly.csv` junto
  al libro. Es un fixture de autoprueba, no un artefacto de cobertura.
- **Pruebas** (`tests/test_model_check.py`), en los dos sentidos:
  demo en verde; divergencia en una celda vinculada; fila sin vínculo
  material; fila sin vínculo bajo el umbral (aviso, exit sin cambio); fila
  vinculada sin periodo en el CSV; sin carpeta de captura con y sin
  `--external`; escala distinta de 1; umbral leído del perfil; fila
  `guidance` que no respalda; CSV mal formado.
  `link_capture` (`tests/test_xlsx_builder_audit.py`): nombre, alcance,
  escala, y los tres rechazos.
- **Contrato del repo.** `SCANNED_C_CHECKS` en `tests/test_repo_contracts.py`
  pasa a `{"C9", "C10"}`.
- **Doctrina.**
  - `integrity-checks.md`: fila C9 reescrita (cómo detecta, el umbral, qué es
    aviso). Fuera la nota "sin implementación única".
  - `templates/coverage-tree.md` §model/inputs: líneas reservadas, regla de
    escala, y `derivado` en la lista de valores de `tag` (el tool ya lo
    emite; el contrato solo nombraba observado/guidance).
  - `skills/xlsx-building/SKILL.md`: `link_capture` en el procedimiento.
  - `skills/statement-mapper/SKILL.md`: un residual o subtotal que el modelo
    necesite como cifra entra al CSV con `tag = derivado` y su fórmula en
    `source_ref`; no se calcula en el script de build.
  - `commands/model-check.md`: C9 sale de la lista de checks que evalúa el
    agente.

## D6 — Migración de AAPL (fase aparte, fuera del repo)

El modelo vigente (v4) falla C9 hasta reconstruirse. Se hace después de
fusionar el cambio del repo, en el workspace de la cobertura:

1. `build_model.py`: `op_row` ya recibe `canon=`; se agrega la llamada a
   `link_capture` ahí y en las filas de `Annual`. BUILD_DATE y VERSION nuevos
   (v5); v4 queda intacto.
2. Sale el C9 de 9 líneas que el script calcula al construir; lo reemplaza
   `model_check.py`.
3. **Gate del analista, fila por fila**, para las 4 de "Otros" y las 3
   semillas de apertura: entrar al CSV como `derivado`, o volverse fórmula.
4. Recalcular con Excel, correr `model_check.py`, registrar en
   `journal/decisions.md`.

## Verificado con Excel (2026-11-25)

Con Excel COM, sobre un libro escrito con openpyxl que lleva un nombre con
alcance de hoja `CAPTURE.bs_ap` → `$10:$10` y comentario `scale=1000000`:

- (a) El comentario del nombre sobrevive a abrir y guardar en Excel.
- (b) Insertar una fila arriba desplaza el rango a `$11:$11`, y openpyxl lo
  relee así.

Consecuencia para el lector: Excel reescribe la referencia sin comillas
(`Operating!$11:$11`), mientras el builder la escribe con ellas
(`'Operating'!$91:$91`). `capture_links` acepta las dos formas.

## Invariantes

- **R1**: quita al agente (y al script de cada cobertura) la verificación de
  C9. Los residuales dejan de calcularse fuera de la capa de captura.
- **R2**: una cifra histórica sin respaldo deja de pasar en silencio.
- **R5**: `model_check.py` sigue siendo solo lectura. El demo escribe su
  fixture en la carpeta del demo, nunca en una cobertura.
- **R6**: el vínculo es parte del xlsx (dueña: `model-standards`, vía
  `xlsx-building`); los CSV siguen siendo de `statement-mapper`.
- **R7**: sin dependencias nuevas.

## Fuera de alcance

- Guidance en la capa de captura y su contraste contra supuestos.
- Tab `Macro` contra `macro/series/` (otra capa de captura, otro check).
- `Rev_Reconcile`, `Val_Comps` y demás hojas auxiliares.
- Los avisos D3, D6 y D10, que siguen sin código.
- Escribir el resultado de C9 en la tab `Checks` al construir.
