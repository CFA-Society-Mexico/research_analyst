# Model Spec — estructura del xlsx estándar (v3)

Contrato que `model-standards` implementa y `/model-check` audita. Las convenciones de
formato y fórmulas viven en `skills/model-standards/references/excel-practices.md`
(base CFI); los checks en `references/integrity-checks.md`; las pestañas de valuación
en `references/valuation-conventions.md`.

Doctrina de forma (diseño 2026-08-31, reemplaza la hoja única y las columnas
intercaladas): **dos hojas por granularidad y propósito** — `Operating`
(trimestral puro: ahí se CONSTRUYE el modelo) y `Annual` (agregados por
fórmula + valuación). Menos tabs, mejor; secciones colapsables en ambas.

## Pestañas (en este orden, modo `quarterly`)

| # | Tab | Contenido | Regla |
|---|---|---|---|
| 1 | `Cover` | Propósito, versión, autor, fecha, leyenda de colores, **celda única de error check** | La celda agrega TODOS los checks del libro |
| 2 | `Checks` | Todos los checks de integridad, uno por fila, con estado | Solo fórmulas; ver integrity-checks.md |
| 3 | `Operating` | **El operating model, trimestral puro** — secciones: Assumptions → IS → BS → CF → Ratios → Schedules | Ver §Operating. ÚNICA hoja con inputs (S6). Sin columnas FY |
| 4 | `Annual` | **Agregados anuales + valuación** — misma estructura (sin Assumptions) + sección DCF | Ver §Annual. CERO números tecleados: fórmulas y links (C8/F14) |
| 5 | `Macro` | Históricos observados de `macro/series/` (última obs + serie anual, con fuente y fecha) + valores de macro/macro-view.yaml: los de primer nivel, o los del bloque `countries.<pais>` que el analista asignó a la emisora (moneda y mercado; se pregunta, jamás se asume). Cada celda cita su ruta (`countries.mexico.rates.risk_free_10y`) | Sin inputs propios; yaml vacío = AVISO D6, la tab nunca queda muerta |
| — | `Rev_Reconcile` | Doble ruta de revenue: bottom-up (bloques Sch) vs top-down (Macro × industria × participación) | Divergencia > umbral → flag en error check |
| — | `Val_Comps` | Múltiplos POR FÓRMULA desde comps/*.yaml; market cap = precio ÷ acciones por ADS × acciones × FX a la moneda de reporte; media armónica; justificados + PVGO | Staleness de snapshots flaggeada (D3); múltiplo fuera de rango marcado `REVISAR` (D12) |
| — | `Val_<método>` | Solo métodos activos del perfil (NAV_AFFO / SOTP). DDM y FCFE son secciones de `Annual` | RI: especificado, inactivo v1 |
| — | `Sensitivity` | Cuadrículas de 2 variables con **fórmula cerrada por celda** (`sensitivity_grid` del builder), aisladas. NO son data tables de Excel: openpyxl no puede crearlos y la cuadrícula quedaba vacía | Tab dedicada; la celda central reproduce el valor del modelo (check en `Checks`) |
| — | `Summary` | Resumen ejecutivo completo (ver §Summary) | Una pantalla; solo fórmulas/links |

Modo `annual`: una sola hoja `Model` anual (estructura del §Operating con
años FY y la sección DCF incluida). La tab `Quarterly` YA NO EXISTE en ningún
modo (la captura trimestral vive en `Operating`).

## §Operating — el operating model (trimestral puro)

- Header: `1Q2016A … 4Q2031E` — SOLO trimestres, corte A/E en el último
  trimestre reportado; histórico trimestral = `quarterly_history_years`
  (perfil, default 10). Los años previos al corte NO aparecen aquí.
- Secciones (banda + outline colapsable, marcador `x` col A, labels col B):
  1. **Assumptions** — por trimestre en TODO el forecast (decisión deliberada:
     la fatiga se paga una vez, el contexto beneficia cada run). Única zona de
     input del libro (S6). Switch de escenarios, constantes nombradas.
  2. **IS / BS / CF** — histórico desde `canonical_quarterly.csv` (C9),
     forecast por drivers. Una serie = una fila (F11/F12).
  3. **Ratios** — `build_ratios`, UNA llamada sobre el horizonte completo:
     una fila por razón cruzando histórico y forecast. PROHIBIDO seccionar
     "Ratios histórico" y "Ratios forecast" (F13 unicidad).
  4. **Schedules** — bloques `Sch: <nombre>` + cores (PPE, Debt, WC). Nunca
     una hoja por schedule (S9).

## §Annual — agregados + valuación

- Header: `FY2016A … FY2031E` (formatos `0"A"`/`0"E"`). **Cero números
  tecleados**: toda celda es fórmula de agregación leyendo `Operating`, o
  link a captura anual observada.
  - Flujos (IS, CF): FY = Σ de los 4 trimestres del año fiscal.
  - Stocks (BS): FY = valor del 4Q.
  - Ratios: RECALCULADOS sobre agregados anuales (no promedio de trimestres).
  - Años pre-corte trimestral: observados anuales de `canonical_annual.csv`
    (misma fila, tramo distinto).
- Secciones: lectura informativa (crecimientos anuales implícitos de los
  supuestos trimestrales — calculada, NO editable) → IS → BS → CF → Ratios →
  Schedules agregados → **DCF/Valuación** → **DDM / FCFE** (si activos en el perfil).

### Sección DCF — línea por línea (obligatorio; cero fórmulas comprimidas)

```
EBIT (link a Annual §IS)                         [hist + forecast]
(-) Impuestos sobre EBIT (t efectiva x EBIT)     [hist + forecast]
NOPAT                                            [hist + forecast]
(+) D&A (link)                                   [hist + forecast]
(-) Capex (link)                                 [hist + forecast]
(-) Delta working capital (link)                 [hist + forecast]
FCFF                                             [hist + forecast]  <- sanity: FCFF realizado visible
Factor de descuento (mid-year opcional)          [solo forecast; año en curso = stub]
PV de FCFF                                       [solo forecast]
Suma PV explicitos
Tasa de reinversion terminal = g / RONIC
Flujo terminal normalizado = NOPAT_n x (1+g) x (1 - g/RONIC)
TV Gordon = Flujo terminal normalizado / (WACC - g)
TV exit = EBITDA_n x multiplo
PV del TV (Gordon) · PV del TV (exit)
Deuda neta actual (desglose citado)
EV (Gordon) · EV (exit) · Equity (Gordon) · Equity (exit)
Acciones diluidas · Valor por accion — Gordon · Valor por accion — exit
Cruce: multiplo implicito del TV Gordon (checks D4 y D11) · g implicita del exit (D4)
Reverse DCF: EV de mercado · TV implicita · g implicita de mercado (D4b)
Bloque Hamada (beta pure-play, mecanica visible)
```

**El terminal Gordon NO perpetúa `FCFF_n`.** El FCFF del último año explícito
trae la reinversión de una empresa que todavía crece (capex de crecimiento);
perpetuarlo a una g baja reinvierte como si se creciera rápido sin recibir ese
crecimiento y hunde el valor. El flujo terminal se normaliza con la
reinversión que la g terminal exige: `g / RONIC`. **RONIC** (retorno sobre el
capital nuevo invertido) es un input etiquetado de Assumptions; su default es
un link a WACC (el crecimiento terminal no crea valor) y el analista lo
sobreescribe con su racional. Lo escribe `gordon_terminal_block` del builder.

Las g implícitas (exit y reverse DCF) se despejan del mismo terminal
normalizado con `implied_growth_block` (forma cerrada con guarda — ver
`valuation-conventions.md`), nunca con un flujo proyectado con la g del
analista, y el año en curso con trimestres ya observados entra como **stub**
(solo trimestres E, fracción de año restante). Fórmulas y exponentes en
`valuation-conventions.md`.

**Layout de los bloques de valor único** (terminal, puente a equity, cruces,
reverse DCF, Hamada): los VALORES se anclan en la **columna C** (pegados a los
labels), jamás al final del horizonte de columnas — el lector no scrollea 16
periodos para ver el valor por acción.

### Secciones DDM / FCFE (condicionales, debajo del DCF)

```
Dividendos al accionista / FCFE (equity_flow_rows)   [hist + forecast]
Terminal de equity normalizado (equity_terminal_block, una vez)
DDM: VP explícitos · VP terminal · equity · por acción · % terminal · flag
FCFE: VP explícitos · VP terminal · equity · por acción · % terminal
```

Exponentes = fila de exponentes del DCF. El año en curso con trimestres
observados entra como stub por `stub_flow_ref` / `stub_exponent_ref` de
`equity_value_block` (suma de los trimestres E del año en `Operating`, con el
exponente del stub del DCF); el rango de flujos cubre entonces solo los años
completos posteriores. D5 en `Checks`; diagnóstico
(`method_consistency_block`) debajo de la última sección de equity.

## Periodicidad (`model_periodicity` del perfil)

- `annual`: columnas anuales FY puras.
- `quarterly` — **modelo trimestral-nativo en dos hojas**
  (`annual_plus_quarterly` DEPRECADO, se trata como `quarterly`): el modelo se
  CONSTRUYE sobre trimestres en `Operating`; lo anual es AGREGADO por fórmula
  en `Annual`, jamás serie paralela ni input (ver §Operating y §Annual).
  - Histórico trimestral: `quarterly_history_years` (default 10); el largo
    plazo pre-corte vive como FY observados en `Annual`.
  - Assumptions POR TRIMESTRE en todo el forecast (la fatiga se paga una vez;
    la entrevista de populate va trimestre por trimestre).
  - DCF sobre los FY agregados de `Annual` (el valor terminal domina; los
    trimestres aportan precisión del agregado, no el descuento). Excepción: el
    año en curso, que entra como stub con solo sus trimestres estimados.

## §Summary — resumen ejecutivo (una pantalla, valores en columna C)

Secciones, todas por fórmula/link — cero tecleo:

1. **Football field**: valor por acción de CADA método activo + precio de
   mercado con fecha + **upside/downside %** por método. Con la cuadrícula de
   Sensitivity: rango bull/bear por método (min-max de la cuadrícula), no solo
   el punto.
2. **Lecturas clave** (~8): CAGR de ventas del forecast · margen EBIT terminal
   · FCF del último año proyectado · spread ROIC−WACC · deuda neta/EBITDA ·
   g terminal del analista · g implícita del mercado (reverse DCF) · brecha de
   expectativas.
3. **Drivers vs guidance**: los 3 drivers más materiales del driver-map —
   supuesto del analista vs guidance citado (si existe) vs realizado UDM.
4. **Sensibilidad clave**: valor/acción a WACC ±1pp y g ±0.5pp (links a la
   cuadrícula de Sensitivity).
5. **Estado del modelo**: celda de error del Cover (link) · fecha del último
   trimestre capturado · fecha del industry-report vigente · fecha de
   macro-view (`updated_at`).
6. **Tesis**: fecha de la última entrada del thesis-journal + estado declarado
   (confirmada / erosionada / pendiente — texto que el analista mantiene en el
   journal; aquí solo se refleja).

## Respiro tipográfico (check F16)

**Una fila EN BLANCO antes de cada header y sub-header** (separa el bloque
anterior); entre headers CONSECUTIVOS no se inserta (un header que sigue
inmediatamente a otro no lleva respiro); tras un header el contenido empieza
en la fila siguiente, sin blanco. El colapso del outline se ve limpio y la
lectura expandida respira.

## Reglas transversales

- Periodos en columnas, líneas en filas; flujo top-to-bottom, left-to-right.
- **Menos tabs, mejor**: el modelo vive en `Operating` + `Annual` (o en `Model`
  en modo annual); solo lo que tiene estructura distinta (comps, cuadrículas
  de sensibilidad, summary) gana tab propia.
- **Formato por código**: el libro se construye vía `tools/xlsx_builder.py`
  (skill xlsx-building); los checks F auditan paleta, fuentes, formatos,
  gridlines, freeze, outline, series y completitud de Ratios.
- **FORECAST COMPLETO (F17, no negociable)**: toda fila con histórico lleva el
  tramo forecast entero con fórmula — una celda E vacía es falla automática.
- **Una serie = una fila** (checks F11 y F12): histórico calculado/observado y
  forecast en la misma fila, rol por columna; nunca filas "histórico" y
  "forecast" separadas; valores derivables hacia atrás (índices, ratios
  implícitos) se POBLAN por fórmula, no se dejan vacíos.
- Una fórmula por fila POR TRAMO (S5): un tramo histórico, un tramo trimestral
  estimado (si aplica), un tramo anual estimado — idéntica dentro de cada tramo.
  **Año de transición** (modo `annual`: el primer año del forecast ya trae
  trimestres observados, p. ej. 9M reales + 4T guiado): su columna es un tramo
  propio con fórmula anclada en lo observado. Se DECLARA con
  `set_transition_year(<año>)` y se escribe con
  `series_row(..., transition_first=...)`; sin la declaración, una fórmula
  distinta en la primera columna E sigue siendo falla de S5. En ese año el
  driver anclado es RESULTADO (fórmula negra), no input.
- Un archivo por modelo; **cero links entre libros**.
- Ninguna cifra nace de generación libre: fórmula o input etiquetado.
- **Fórmulas que Excel acepta (F21)**: `set_cell` rechaza al construir una
  función con aridad inválida (`ABS(a,b,c)`), un paréntesis sin cerrar, una
  función moderna sin prefijo `_xlfn.` o una operación de arreglo en fórmula
  normal. Una operación sobre rangos va dentro de `SUMPRODUCT` o se escribe
  con `set_array_formula` (p. ej. `MEDIAN(IF(rango=1, valores))`).
- Cada línea de forecast referencia un driver de la sección Assumptions —
  línea sin driver = flag de driver-inventory.
- Etiquetas de trazabilidad: observado (azul con cita en comentario), guidance
  (marcado), supuesto (azul + amarillo), output (fórmula negra), link (verde).
- Versionado `_YYYY-MM-DD_v#`; nunca borrar columnas a media serie.
