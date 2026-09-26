---
name: model-standards
description: Estándar del modelo 3 estados del plugin — construye el xlsx completo (Operating trimestral + Annual agregada, o Model en modo anual; Schedules como bloques y Ratios con DuPont/ROIC/CCC como secciones, valuación determinista con terminal dual, reverse DCF y múltiplos justificados, históricos poblados desde la capa de captura model/inputs/) vía el builder determinista, corre la lista completa de ~50 checks S/C/D/F, y audita o RECONSTRUYE cualquier modelo existente al estándar con paridad de números. Usa esta skill siempre que haya que construir un modelo financiero en Excel, correr o verificar checks de integridad (balance, tie-out, reconciliación de revenue, formato), armar pestañas de valuación (DCF, comps, DDM, NAV/AFFO, SOTP, Ratios), auditar o migrar un modelo existente, o cuando el usuario pida "construye el modelo", "revisa mi modelo", "corre los checks", "arma la valuación", "reconstruye mi Excel al estándar" — aunque el modelo no lo haya creado este plugin.
---

# model-standards

Dueña única del xlsx. Nadie más escribe en el modelo. Construye contra dos contratos:
`templates/model-spec.md` (estructura) y `driver-map.md` (raíz del ticker — qué schedules
existen). **Toda cifra sale de fórmula o de un input etiquetado del analista — esta
skill jamás genera un número.**

## Cuándo corre

- `/init-coverage` paso 6: construir el modelo (requiere driver-map aprobado).
- `/update-quarter` paso 4: re-correr checks tras capturar el trimestre.
- `/model-check`: solo auditoría, standalone.
- Standalone: auditar cualquier modelo contra el estándar (ver abajo).

## Inputs (contratos de entrada)

| Input | De quién | Qué trae |
|---|---|---|
| `model/inputs/canonical_*.csv` | `statement-mapper` | cifras `observado` por línea, con cita en los `extract_*.json`; el histórico del xlsx se puebla desde aquí (check C9) |
| `driver-map.md` (raíz) | `driver-inventory` (design) | drivers clave + spec de schedules + doble ruta de revenue |
| `issuer-profile.yaml` (raíz) | `framework-mapper` | marco, convenciones, métodos de valuación activos |
| `macro/macro-view.yaml` (workspace) | analista | rf, ERP, FX, decks — para tab Macro |
| `comps/*.yaml` | `statement-mapper` | snapshots para Val_Comps |

Falta un contrato → detente y repórtalo; no construyas parcial ni inventes el insumo.

## Procedimiento de construcción

1. Lee los cinco contratos. Confirma con el usuario: periodos históricos, horizonte,
   periodicidad (`model_periodicity` del perfil — si el campo está vacío o sin
   confirmar, DETENTE y pregunta; anual nunca se asume) y métodos activos (del
   perfil — puede sobreescribir).
2. Genera el xlsx VÍA `tools/xlsx_builder.py` (skill `xlsx-building` —
   obligatoria, nunca openpyxl crudo para estructura/formato) siguiendo
   `model-spec.md`: tabs en orden, formato 100% por código
   (`references/excel-practices.md` es el contrato legible).
3. Modo `quarterly` (model-spec §Operating/§Annual): **dos hojas** —
   `Operating` (trimestral puro, secciones Assumptions → IS → BS → CF →
   Ratios → Schedules; única zona de input) y `Annual` (agregados por fórmula
   + sección DCF línea por línea; cero tecleo — C8/F14). Schedules: bloques
   `Sch: <nombre>` + cores, jamás hojas sueltas (S9). Modo `annual`: una hoja
   `Model` anual.
4. Sección `Ratios` generada por `ModelStyler.build_ratios` con el registro
   canon→fila del build, UNA llamada por hoja sobre el horizonte completo
   (completitud Y unicidad = check F13; jamás a mano, jamás seccionada en
   histórico/forecast).
5. Pestañas de valuación según métodos activos — fórmulas y convenciones exactas en
   `references/valuation-conventions.md` (incluye bloque Hamada, reverse DCF y
   múltiplos justificados + PVGO). Usuario en el loop en cada supuesto.
6. Corre TODOS los checks de `references/integrity-checks.md` — S/C/D más los F
   de formato (`python tools/xlsx_builder.py audit <modelo>`, exit 0 exigido).
   Cualquier rojo: arregla estructura/fórmulas (nunca "ajustes" a cifras
   observadas) y re-corre.
7. Entrega solo con checks verdes. Check rojo irresoluble = reporta FALLA con celdas
   afectadas. Un modelo que "casi cuadra" no existe.

## Modo auditoría (standalone y /model-check)

Sobre cualquier xlsx (aunque no lo haya creado el plugin): corre la lista completa de
`integrity-checks.md`, reporta `[ok]/[x]` por check con celdas afectadas. No edites
el modelo ajeno sin instrucción explícita del usuario. Si el usuario pide
RECONSTRUIRLO al estándar: modo rebuild de `xlsx-building` (audit primero →
gate → rebuild con contenido preservado → paridad obligatoria).

Precondición standalone: ninguna. Es el punto de entrada más barato del plugin.

## Trazabilidad en el xlsx

- `observado`: input con comentario de celda citando fuente y página.
- `guidance`: input marcado como afirmación de management — nunca alimenta una
  fórmula sin que el analista lo convierta explícitamente en supuesto propio.
- `supuesto`: azul, en Assumptions, puesto por el analista.
- `output`: fórmula negra. Punto.

## Entrevista de cierre

Objetivo (preguntas adaptativas — generarlas desde lo construido y el journal, según
`templates/debate-protocol.md`): **¿el modelo captura el negocio real o solo la
contabilidad?** Sondear: qué bloque de Schedules falta, qué línea del build-up quedó
floja, si la estructura refleja cómo la emisora gana dinero. Confrontar el spread
ROIC−WACC histórico (tab Ratios) contra el moat que afirma industry-report, y la g
implícita de mercado (reverse DCF, check D4b) contra la g del analista. Registrar
en thesis-journal.

## Referencias

- `references/excel-practices.md` — convenciones CFI: leer antes de construir.
- `references/integrity-checks.md` — la lista S/C/D completa: leer antes de auditar.
- `references/valuation-conventions.md` — leer antes de armar cualquier Val_*.
- `references/ratios-analytics.md` — spec de la tab Ratios y del bloque `Sch: WC`:
  leer antes de construir Ratios o Schedules.
