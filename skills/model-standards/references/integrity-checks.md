# Integrity Checks — la lista que corre /model-check

Cada check es determinista: fórmula en la tab `Checks` o escaneo por código (openpyxl).
La celda de error del Cover agrega TODOS. Un check rojo = el comando reporta FALLA,
nunca éxito. Orden: estructura → contabilidad → contenido → formato.

Los checks F corren con `python tools/xlsx_builder.py audit <modelo.xlsx>` —
implementación única, cero interpretación del agente. Exit 0 = verde; exit 1 =
alguna falla; exit 3 = sin fallas pero con checks `[pendiente]` (no se pudieron
evaluar, p. ej. libro sin valores calculados). Un pendiente NUNCA se reporta
como verde.

## Estructurales (escaneo por código)

| # | Check | Cómo |
|---|---|---|
| S1 | Todas las tabs del model-spec vigente presentes (Cover, Checks, `Operating` + `Annual` en modo quarterly o `Model` en modo annual, Macro, condicionales); en la sección Schedules de `Operating` (o `Model`), un bloque `Sch: <nombre>` por schedule del driver-map + cores (PPE, Debt, WC) | openpyxl: nombres de hoja vs contrato; escaneo de col. A por headers `Sch: ` vs driver-map |
| S9 | Schedules como bloques, no como tabs: ninguna hoja con nombre `Sch_*` o `Sch *`; tampoco hojas `IS`/`BS`/`CF`/`Val_DCF` sueltas (viven como secciones de `Operating`/`Annual`, o de `Model` en modo annual) | escaneo de nombres de hoja |
| S10 | Sin ciclos ni cadenas rotas: tras recalcular (Excel COM), ninguna celda de FÓRMULA queda sin valor cacheado | el síntoma del smoke #5: circularidad caja↔rendimiento↔utilidad dejó el forecast entero sin calcular; regla preventiva: rendimientos/intereses sobre saldo de APERTURA |
| S2 | Cero links externos a otros libros | escaneo de fórmulas por `[` |
| S3 | Sin funciones volátiles (OFFSET, INDIRECT, NOW, TODAY) | escaneo de fórmulas |
| S4 | Sin números hard-coded dentro de fórmulas (fuera de Assumptions) | escaneo: constantes en fórmulas de tabs de cálculo |
| S5 | Una fórmula por fila: fórmula idéntica (relativa) en todos los periodos | comparación de R1C1 por fila |
| S6 | Inputs solo en la SECCIÓN Assumptions de `Operating` (o `Model` en modo annual); resto del libro — `Annual` incluida — sin celdas de input | escaneo: fill de input fuera del rango de la sección = falla |
| S7 | Modo de cálculo automático, no manual | propiedad del libro |
| S8 | Versionado del archivo cumple `_YYYY-MM-DD_v#` | nombre de archivo |

## Contables (fórmulas en tab Checks)

| # | Check | Fórmula |
|---|---|---|
| C1 | Balance cuadra cada periodo | A − (P + C) = 0, todos los periodos |
| C2 | Tie-out de caja | caja final CF = caja BS, todos los periodos |
| C3 | Roll de utilidades retenidas | RE₁ = RE₀ + NI − dividendos |
| C4 | Depreciación acumulada consistente con schedule | BS vs bloque `Sch: PPE` de Schedules |
| C5 | Interés consistente con schedule de deuda | IS vs bloque `Sch: Debt` (documentar switch si hay circularidad) |
| C6 | Identidad DuPont | ROE directo (NI/capital prom.) − ROE DuPont 5 factores = 0, todos los periodos (tab Ratios) |
| C7 | CCC del forecast consistente con schedule de WC | Ratios (forecast) vs días DIO/DSO/DPO del bloque `Sch: WC` |
| C8 | Agregado anual estructural (modo `quarterly`) | La hoja `Annual` es 100% fórmulas: cada FY del tramo trimestral agrega los 4 trimestres de `Operating` (flujos = Σ4Q; stocks = 4Q; ratios recalculados); pre-corte, links a canonical_annual. Un número tecleado en `Annual` = falla (F14 lo detecta por fill de input) |
| C10 | **Desfase del roll de caja** | Escaneo de fórmulas: (i) cada celda de "Efectivo al inicio" referencia la celda de CIERRE de la columna previa; (ii) cada celda de caja del BS referencia el cierre del CF de su misma columna; (iii) rendimientos/intereses referencian saldos de la columna PREVIA. Una referencia a cierre de la misma columna en (iii) = ciclo = falla |
| C9 | Histórico del xlsx = capa de captura | Escaneo por código: celdas históricas de IS/BS/CF coinciden con `model/inputs/canonical_annual.csv` (y `canonical_quarterly.csv` si aplica), tolerancia de redondeo; discrepancia = falla de captura o edición manual del histórico |

## De contenido (doctrina del plugin)

| # | Check | Regla |
|---|---|---|
| D1 | **Cobertura de drivers**: cada línea de forecast referencia un input de Assumptions | línea sin driver = flag (driver-inventory la reporta) |
| D2 | **Reconciliación de revenue**: ruta bottom-up (schedules) vs top-down (PIB × industria × participación) | divergencia > umbral (default 10%) = flag: supuesto roto. **Solo aplica a años PLENAMENTE ESTIMADOS** (0 trimestres observados): la reconciliación contrasta dos rutas de SUPUESTOS, y un año ya ocurrido no se reconcilia contra una fórmula de PIB × prima — se observa. Detección genérica con `COUNTIF` de headers `*Q<año>A` en Operating: 4 observados ⇒ `n/a (observado)`; 1–3 ⇒ `n/a (transición)`; 0 ⇒ reconcilia de verdad. Si un año estimado diverge, ahí sí es supuesto roto |
| D3 | **Staleness de comps**: `as_of` de cada snapshot dentro de `stale_days` | snapshot viejo = flag en Val_Comps |
| D4 | **Cruce de terminales**: múltiplo implícito por Gordon y g implícita por exit multiple, ambos visibles | divergencia grande = revisar g o múltiplo con el analista |
| D5 | **Convergencia de métodos**: si DDM/FCFE/RI activos junto a DCF, divergencia extrema entre métodos se reporta como supuesto inconsistente | nunca se "promedia" para ocultarla |
| D6 | **Staleness de macro-view**: `updated_at` dentro de `staleness_warn_months` | aviso, no bloqueo |
| D7 | **Monitor NIF B-10**: si emisora NIF y acumulado trienal en macro-view ≥ 22%, exigir escenario de reconexión pre-modelado | ver framework-mapper reference |
| D8 | Etiquetas de trazabilidad presentes: inputs con comentario de fuente (observado) o marca guidance/supuesto | muestreo por tab |
| D9 | **Calidad de utilidades**: CFO/NI < 1 en ≥ 2 periodos consecutivos, o accruals ratio fuera de la banda histórica de la emisora | aviso, no bloqueo; material de entrevista |
| D10 | **Staleness de industry-report**: fecha del reporte más reciente en `research/industry/` dentro de `industry_staleness_months` (issuer-profile, default 6) | aviso, no bloqueo — recomendar `/update-industry` |
| D4b | **g implícita de mercado** (reverse DCF en Val_DCF) vs g terminal del analista | divergencia grande = pregunta obligada en entrevista de cierre; nunca bloquea |

## De formato (checks F — `tools/xlsx_builder.py audit`)

Whitelists exactas viven en el código (`tools/xlsx_builder.py`), derivadas del
corpus CFI (2026-08-30). El builder pasa estos checks por construcción; un modelo
externo se audita igual.

| # | Check | Regla |
|---|---|---|
| F1 | Gridlines ocultas | `showGridLines = False` en TODAS las hojas visibles |
| F2 | Fuente estándar | Una sola familia (Aptos Narrow, `FONT_NAME` del builder) en celdas usadas |
| F3 | Colores de fuente | Whitelist: negro, azul input `FF0000FF`, verde link `FF00CC00`, rojo warn, blanco |
| F4 | Paleta de fills | Whitelist: azul oscuro `FF1F4E79` (barra y bandas de sección), azul claro `FFBDD7EE` (sub-secciones), amarillo input `FFFFF2CC`, gris escenario `FFF2F2F2`. La paleta CFI navy/naranja/teal quedó FUERA (identidad 2026-08-31). Con `brand/DESIGN.md` presente, sus 3 slots se suman a la whitelist (pasar el archivo al audit) |
| F5 | Formatos numéricos | Whitelist literal (miles con paréntesis y guion-cero, %, 0.0x, USD, fecha, A/E, `;;;`) |
| F6 | Freeze panes | Presente en hojas de datos (Assumptions/IS/BS/CF/Ratios/Schedules/Rev_Reconcile/Val_*) |
| F7 | Outline POR SECCIÓN | En `Operating` y `Annual` (o `Model`/`Schedules` legacy): CADA sección con contenido (marcador `x` + header bold) tiene filas agrupadas — agrupar solo algunas secciones FALLA |
| F8 | Sin hojas basura | Ninguna `Sch_*`, `Hoja1`, `Sheet1` (refuerza S9) |
| F9 | Headers de periodo A/E | Formatos `0"A"` / `0"E"` presentes en la fila de años |
| F10 | Sello del builder | Custom property `research_analyst_builder` presente — el modelo se construyó vía `tools/xlsx_builder.py`. En modelo externo: `[aviso]`, no falla |
| F11 | **Continuidad de series** | Una serie = UNA fila continua en todo el horizonte: histórico calculado/observado y forecast en la misma fila (el rol cambia por columna), jamás columnas históricas vacías. Detección: toda fila con ≥3 celdas de input (fill amarillo) en columnas de periodo debe tener TODAS las columnas de periodo pobladas. Requiere headers A/E (F9) para ubicar las columnas — F9 rojo deja F11 sin efecto |
| F12 | **Sin series partidas** | Fila cuyo label contiene "forecast" con la mitad histórica (columnas A) vacía, o "histórico" con la mitad estimada (columnas E) vacía = la serie se partió en dos filas. Complementa F11 (que solo vigila filas de input); cacha el patrón exacto del smoke AAPL. Derivables hacia atrás (índices, ratios implícitos) se POBLAN por fórmula |
| F13 | **Ratios completa y ÚNICA** | (i) Completitud: el set completo de `REQUIRED_RATIO_LABELS` (~25) presente; (ii) UNICIDAD: cada razón UNA sola vez por hoja — un label duplicado delata secciones "Ratios histórico"/"Ratios forecast" partidas (bug del smoke #3); la serie completa vive en una fila. `build_ratios` cumple por construcción (iii) Cada razón con FÓRMULA en sus columnas: una etiqueta sin fórmulas no cuenta como presente |
| F14 | **Modelo trimestral-nativo** | Con sello `periodicity=quarterly`: (i) `Operating` trae ≥4 trimestres `A` y ≥4 `E` en el header — el modelo se CONSTRUYE sobre trimestres; (ii) `Annual` es FY-solo (cero columnas `#Q`) y sin UNA sola celda de input. Sin sello: n/a |
| F19 | **Roll de caja cerrado** | En TODAS las columnas, incluido el histórico: (i) `inicio(t) = cierre(t−1)` y (ii) **`inicio(t) + cambio neto(t) = cierre(t)`**. El (ii) es el que caza un CF INCOMPLETO: si falta una sección del flujo (p. ej. el movimiento de valores negociables — el mayor flujo después del operativo en emisoras con tesorería grande), C1 y C2 siguen en verde porque ambos leen el MISMO efectivo observado, mientras el roll no cierra en silencio. Bug del smoke #5: 38 trimestres históricos con el roll roto y todos los demás checks verdes Requiere valores calculados: si las filas del roll tienen fórmulas sin valor (libro recién guardado por openpyxl), reporta `[pendiente]` y exit 3 — jamás `[ok]`. Recalcular (Excel COM) y re-auditar |
| F20 | **Cobertura del audit** | Los checks F escanean hasta 5,000 filas × 400 columnas por hoja (`_MAX_SCAN_ROWS` / `_MAX_SCAN_COLS`). Una hoja más grande FALLA y se nombra: antes el tope era 400 × 40 y lo de abajo (los schedules de `Operating`) pasaba sin auditar y en verde. F18/F19 que truenan por excepción se reportan `[pendiente]`, no `[ok]` |
| F18 | **Sin referencias circulares** | Grafo de dependencias de fórmulas + DFS, SIN necesitar Excel: un ciclo = el forecast no calcula (causa raíz del smoke #5: `caja → otros ingresos → utilidad → CFO → caja`). Reporta la cadena completa del ciclo. Ratios legítimos de la misma columna (EBT/EBIT) NO son ciclos y no aparecen. Complementa S10 (que sí requiere recalc) |
| F17 | **FORECAST COMPLETO** (mandato duro) | Toda fila con histórico (≥3 celdas en columnas A) tiene el tramo E COMPLETO — una sola celda de forecast vacía = FALLA con fila y conteo. "Los forecast deben tener todas las fórmulas completas" — regla de primera clase, no negociable |
| F16 | **Respiro antes de headers** | Fila EN BLANCO antes de cada header/sub-header (salvo headers consecutivos y el tope de la hoja); tras un header el contenido empieza sin blanco. Elegancia auditada, no opcional. **Prevención en el builder:** `section_header`/`subsection` FALLAN al construir si la fila previa tiene contenido y no es header (v0.6.0) — el respiro deja de depender de que el agente se acuerde |
| F15 | **Series sin huecos ni texto** (S5 hecho código) | Dos venenos: (a) fila con contenido en ≥ mitad de las columnas de periodo pero con celdas VACÍAS = fórmula faltante en un tramo (la NOPAT del smoke #4); (b) TEXTO literal (`n/d`, `n/a`) en columnas de periodo de una fila de serie = rompe la cadena de cálculo (#VALUE! aguas abajo). El hueco de dato es DECISIÓN del analista con gate: 0 explícito con comentario, carry-forward del último disponible (fórmula marcada supuesto), o exclusión documentada — jamás texto. Bloques de valor único no disparan; detección reconoce años A/E Y trimestres de texto |

## Reporte de /model-check

```
[ok] / [x] por check, agrupado S / C / D / F
Resumen: N ok, M fallas, K avisos
Falla => exit report "FALLA", lista de celdas/hojas afectadas, siguiente acción sugerida
```

## Resultados en la tab Checks: fórmula viva o escaneo FECHADO

Un check de la tab `Checks` es una **fórmula viva** sobre el libro (C1–C7, D2,
D4, D9) o el **resultado de un escaneo por código** al construir (C8, C9, D3,
D6, D10, D1). Un literal `"OK"` sin fecha ni nota es un check congelado: queda
verde aunque el hecho cambie (bug de la auditoría AAPL 2026-09-02: D3 decía
`OK` con snapshots vencidos). Regla: todo resultado que no sea fórmula se
escribe con `ModelStyler.check_result(ws, row, col, resultado, note=..., computed_at=YYYY-MM-DD)`
— el builder rechaza el literal sin `computed_at` y sin `note`, y deja la
fecha de cálculo en la celda contigua para que el lector sepa cuándo dejó de
ser verdad.
