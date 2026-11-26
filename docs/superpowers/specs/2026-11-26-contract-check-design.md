# Diseño — validador de contratos de la cobertura (checks K)

Fecha: 2026-11-26 · Aprobado por Alan (brainstorming P1 = PyYAML opcional ·
P2 = falla lo estructural, avisa lo tolerable).
Origen: el dogfood de `model_check.py` sobre tres coberturas reales. Los
contratos se rompen en la captura y se descubren tarde, cuando un check del
modelo tropieza con ellos.

## El problema

Los contratos de `templates/` los escribe el agente y nada los verifica.
Medición sobre tres coberturas reales (2026-11-26, solo lectura):

| Contrato | Hallazgo |
|---|---|
| `macro/macro-view.yaml` | No es YAML válido: un parser truena en la línea 258 (bloque `note:` mal indentado). Es el archivo compartido por todas las coberturas. El check D6 lo lee por línea, encontró `updated_at` y lo reportó vigente |
| `comps/*.yaml` | 13 de 13 snapshots vigentes (dos coberturas) sin `price_currency`, `shares_per_ads` ni `fx_to_reporting`, obligatorios desde v0.7.0. Una usa la llave vieja `currency` |
| `issuer-profile.yaml` | Una cobertura usa `model_periodicity: annual_plus_quarterly`, valor deprecado |
| `canonical_*.csv` | Una cobertura con 41 filas de etiqueta fuera de la lista (`observado (signo normalizado…)`); otra con el CSV en formato ancho (una columna por periodo) |

## Decisiones de brainstorming

| # | Decisión | Elección |
|---|---|---|
| P1 | Cómo leer YAML | **PyYAML como dependencia opcional.** Con PyYAML el validador revisa todo; sin él, los checks de YAML salen `[pendiente]` y el de CSV corre igual. Solo un parser real dice si un archivo es YAML válido, y los campos obligatorios de comps están anidados |
| P2 | Qué pasa cuando un contrato no cumple | **Falla lo estructural, avisa lo tolerable.** Falla: no parsea, obligatorio vacío, valor fuera de lista, CSV sin columnas, llave duplicada. Aviso: campo desconocido, valor deprecado con equivalente, etiqueta válida con texto extra |

Descartadas: un lector propio de YAML (medio parser por escribir; no detecta
el archivo roto y sus bugs serían falsos positivos) y validar solo lo legible
por línea (deja fuera los dos hallazgos más graves).

## D1 — La herramienta

```
python tools/contract_check.py <carpeta de la cobertura> [--macro macro-view.yaml]
```

- Recibe la carpeta de UNA cobertura (`<raiz>/<TICKER>`). Solo lee.
- `--macro`: default `<raiz>/macro/macro-view.yaml` (la carpeta hermana).
- Reusa `Finding` de `xlsx_builder` (`[ok]` / `[x]` / `[aviso]` / `[pendiente]`).
  Exit codes de la casa: 0 verde, 1 alguna falla, 3 sin fallas pero con
  pendientes, 2 carpeta inexistente.
- Tool aparte de `model_check.py`: ese audita el xlsx; este audita los
  insumos y corre antes de que exista un modelo.
- Sin PyYAML: K1–K4 `[pendiente]` con la instrucción
  `python -m pip install pyyaml`; K5 corre.

## D2 — Los checks

Un `Finding` por check. Cada uno junta problemas duros (falla) y blandos
(aviso): con algún duro es `[x]`; solo blandos, `[aviso]`; ninguno, `[ok]`.
El detalle nombra archivo y campo.

### K1 — todo YAML parsea

`issuer-profile.yaml`, `macro-view.yaml` y cada `comps/*.yaml`. Falla con
archivo, línea, columna y el problema que reporta el parser. Un archivo que no
parsea deja su check de contenido (K2/K3/K4) como `[pendiente]`: no se puede
evaluar.

### K2 — `issuer-profile.yaml`

No existe el archivo: falla (una cobertura sin perfil no es una cobertura).

| Regla | Campos | Severidad |
|---|---|---|
| Obligatorio, no vacío | `ticker`, `name`, `as_of`, `listing`, `issuer_type`, `framework_governing`, `currency_reporting`, `units`, `fiscal_year_end`, `model_periodicity` | falla |
| Valor de su lista | `listing` (bmv, sec, dual, private) · `issuer_type` (non_financial, bank, insurer, fibra, private) · `framework_governing` (ifrs, us_gaap, nif, cnbv_criteria) · `model_periodicity` (annual, quarterly) · `life_cycle_stage` si no está vacío (embryonic, growth, mature, decline) | falla |
| Valor deprecado | `model_periodicity: annual_plus_quarterly` (se trata como `quarterly`) | aviso |
| Formato | `as_of` con fecha `YYYY-MM-DD` | falla |
| Formato | `fiscal_year_end` `MM-DD` | aviso — un año fiscal de 52/53 semanas ("último sábado de septiembre") es legítimo y no cabe en `MM-DD` |
| Confirmación del analista | `framework_confirmed_by_user`, `periodicity_confirmed_by_user`, `mnpi_confirmed_public` deben ser `true` | falla |
| Tipo, si el campo está | `adr` booleano; `quarterly_history_years` e `industry_staleness_months` enteros > 0; cada valor de `materiality` número ≥ 0; cada valor de `valuation_methods` booleano | falla |
| Campo de primer nivel que el contrato no conoce | — | aviso |

Un campo opcional ausente no es hallazgo: aplica el default del contrato.

### K3 — `macro/macro-view.yaml`

No existe: aviso (igual que D6). `updated_at` vacío o sin fecha: falla.
`staleness_warn_months`, si está, entero > 0: falla. Campo de primer nivel
desconocido: aviso.

Bloques por país (agregado 2026-11-26, decisión del analista: mapa
`countries:`, bloque abierto, `null` avisa, `gate_status` en el contrato;
componentes heredan el source; el primer nivel sigue la misma regla):

| Regla | Severidad |
|---|---|
| `countries` no es mapa; un bloque no es mapa; `as_of` del bloque sin fecha | falla |
| Todo mapa con `value` (primer nivel y bloques) sin `source` propio ni de un mapa que lo contenga, o con `value` no numérico | falla |
| `value: null` con source (pendiente declarado) | aviso |
| Bloque sin `as_of`, sin `rates.risk_free_10y` o sin `equity_risk_premium` | aviso |
| Bloque de país suelto en el primer nivel (workspace anterior) | aviso con pista de migración; sus datos se validan igual |
| Sección propia de un país (labor, tax) | sin hallazgo: el bloque es abierto |

### K4 — `comps/*.yaml`

Sin snapshots: n/a. Se valida el snapshot **vigente** de cada comparable (el
de `as_of` más reciente; mismo criterio que D3): los anteriores son historia,
y el contrato dice que siguen legibles.

| Regla | Severidad |
|---|---|
| `ticker` y `as_of` no vacíos; `as_of` contiene una fecha | falla |
| `stale_days`, si está, entero > 0 | falla |
| `market_data` trae `price`, `price_currency`, `shares_per_ads`, `fx_to_reporting`, `shares_outstanding` | falla |
| `price_currency` es ISO 4217 (tres letras mayúsculas) | falla |
| Cada dato capturado (`price`, `shares_per_ads`, `fx_to_reporting`, `shares_outstanding`, `market_cap`, y cada entrada no vacía de `balance_items` / `income_items`) trae `value` numérico y `source` no vacío | falla |
| `fx_to_reporting` trae además `as_of` | falla |
| `framework` fuera de (ifrs, us_gaap, nif) | aviso — un comparable extranjero puede reportar bajo otro marco; es una pista de comparabilidad, no estructura |
| Campo desconocido, de primer nivel o dentro de `market_data`, `balance_items` o `income_items` | aviso. Solo los datos que el template nombra en esas secciones se validan como `{value, source}`; un campo de texto ajeno (p. ej. `currency`) no es un dato roto |

### K5 — `model/inputs/canonical_*.csv`

Sin ninguno de los dos archivos: n/a (cobertura sin captura todavía).

| Regla | Severidad |
|---|---|
| Columnas `line_item`, `period`, `value`, `source_doc`, `source_ref`, `tag` | falla (y no se evalúa el resto del archivo) |
| `period`: `FY2016` en el anual, `3Q2016` en el trimestral | falla |
| `line_item` no vacío; `value` numérico | falla |
| Llave `(line_item, period)` repetida | falla |
| `tag` fuera de (observado, derivado, guidance) | falla |
| `tag` que empieza con una etiqueta válida y trae texto extra | aviso |
| Fila `observado` sin `source_doc` | falla (R2: un observado lleva su documento) |

El signo de los pagos (C11) y la coincidencia con el modelo (C9) siguen en
`model_check.py`.

## D3 — De dónde salen las reglas

Tablas dentro de `tools/contract_check.py`. No se leen de los comentarios de
los templates: frágil.

Para que no se desfasen, prueba de contrato del repo
(`tests/test_repo_contracts.py`): el conjunto de campos de primer nivel de
cada `templates/*.yaml`, y el de `market_data` del snapshot, es exactamente el
que el validador conoce. Quien agrega un campo al template y olvida el
validador rompe la CI.

Consecuencia: `pyyaml` entra a `requirements.txt` (la CI lo necesita).
"Opcional" significa que el tool no truena sin él.

## D4 — Dónde se conecta

- Preflight de `/update-quarter`, `/model-check`, `/update-industry` y
  `/update-macro`.
- `/init-coverage`: después de los gates de `framework-mapper` (perfil) y
  `statement-mapper` (captura y comps). Antes de esos gates el perfil no está
  confirmado y K2 fallaría por diseño.
- Doctrina: sección K en
  `skills/model-standards/references/integrity-checks.md`, junto a las demás
  familias; `AGENTS.md` y `CONTRIBUTING.md` nombran el tool y la dependencia.

## Coberturas existentes

Sin cambio de contrato: el validador solo lee. Pero **hace visible lo que ya
estaba roto**: con las reglas de arriba, las tres coberturas medidas fallan
(macro-view compartido; comps sin los campos obligatorios; un CSV en formato
ancho). Corregirlas es trabajo de cada cobertura, con decisión del analista,
y no forma parte de este cambio.

## Invariantes

- **R1**: quita al agente la revisión a ojo de sus propios contratos.
- **R2**: K5 exige documento en cada `observado`.
- **R5**: solo lectura.
- **R6**: no escribe ningún artefacto; cada contrato sigue con su skill dueña.
- **R7**: una dependencia nueva, opcional y justificada (CONTRIBUTING 4.3):
  tres de los cuatro contratos son YAML.

## Fuera de alcance

- Corregir los contratos de las coberturas existentes.
- `driver-map.md`, reportes de industria y journals (Markdown).
- Reglas que el contrato no declara hoy.
- Migrar los lectores por línea de `model_check.py` (C9, D3, D6, D10) a
  PyYAML: leen un campo suelto y no necesitan la dependencia.
