# Excel Practices — contrato de formato (base CFI, verificado empíricamente)

Convenciones obligatorias del modelo. Fuente: corpus CFI extraído 2026-08-30
(3-Statement Model Complete, caso AMZN avanzado, Valuation Model, template
library) — valores medidos, no doctrina de blog.

**Regla cero: el formato es CÓDIGO.** Todo workbook del plugin se construye vía
`tools/xlsx_builder.py` (`ModelStyler`). Prohibido aplicar estilos con openpyxl
crudo o "a juicio": paleta, fuentes, formatos numéricos, bordes, gridlines,
freeze y outline salen del módulo. `/model-check` audita esto con los checks F
(`python tools/xlsx_builder.py audit <modelo>`), whitelists en el código.

## Estructura

- **Dos hojas por granularidad y propósito** (modo `quarterly`, diseño
  2026-08-31): `Operating` = el operating model TRIMESTRAL PURO (secciones
  Assumptions → IS → BS → CF → Ratios → Schedules con outline; única zona de
  input); `Annual` = agregados por fórmula (flujos Σ4Q, stocks 4Q, ratios
  recalculados; cero tecleo) + sección DCF línea por línea. Modo `annual`:
  una hoja `Model`. Tabs aparte solo para estructura distinta: Rev_Reconcile,
  Val_Comps, Sensitivity, Summary, Macro.
- Menos tabs, mejor. El conteo de tabs es fijo (model-spec); crecen secciones.
- Histórico trimestral: `quarterly_history_years` (default 10); el largo plazo
  pre-corte vive como FY observados en `Annual`. Assumptions por trimestre en
  todo el forecast (la fatiga se paga una vez, el contexto beneficia cada
  run). S5 (una fórmula por fila) aplica POR TRAMO.
- Flujo top-to-bottom, left-to-right; periodos en columnas, líneas en filas.
- Un archivo por modelo; cero links entre libros.
- Cover con propósito, versión, autor, fecha, leyenda de colores y la celda
  única de error check.

## Layout de toda hoja de datos (lo aplica el builder)

- Fila 1-2: barra azul oscuro con copyright + título de hoja (16 bold blanco) +
  unidades en el header — "(USD millones salvo indicado)" — nunca en celdas.
- Fila 3: fila de CHECK por columna (ej. balance check), visible siempre.
- Fila 3/4: header de periodos con sufijo A/E (`0"A"` / `0"E"` → 2025A, 2026E).
- Freeze panes en A4/C4: título + años + check congelados al scrollear.
- Gridlines ocultas en TODAS las hojas (CFI: 100% del corpus).
- **Columna A = navegación** (patrón CFI del caso AMZN): angosta (~2.5), con
  SOLO una "x" en cada fila de header/sub-header — Ctrl+flecha sobre la
  columna A salta de sección en sección. Las etiquetas viven en columna B
  (~42), descriptivas, sin crípticos. El builder lo aplica solo
  (`_nav_mark` en section_header/subsection/check_row).
- Secciones dentro de la hoja: header con banda azul oscuro y texto blanco
  bold 16; sub-secciones con banda azul claro bold 14; contenido agrupado
  (outline nivel 1) para colapsar a vista resumen.

## Colores (paleta cerrada — check F3/F4)

| Uso | Valor |
|---|---|
| Barra de marca / cover | Azul oscuro `FF1F4E79`, texto blanco |
| Header de sección | Azul oscuro `FF1F4E79`, texto BLANCO bold 16 |
| Sub-sección | Azul claro `FFBDD7EE`, texto negro bold 14 |
| Acentos / tab color | Azul claro `FFBDD7EE` |
| Input del analista | Fuente azul `FF0000FF` + fill amarillo claro `FFFFF2CC` |
| Observado (histórico citado) | Fuente azul `FF0000FF`, sin fill, comentario con fuente |
| Fórmula | Fuente negra |
| Link a otra hoja | Fuente verde `FF00CC00` |
| Error/warn | Fuente roja `FFFF0000` |
| Área de escenario | Fill gris `FFF2F2F2` |

(La paleta CFI navy/naranja/teal queda como referencia histórica, FUERA de la
whitelist F4.)

Fuente única: **Aptos Narrow** — jerarquía por tamaño: 11 normal (y 11 bold
subtotales) → 14 bold sub-sección → 16 bold sección y título de hoja.

## Formatos numéricos (whitelist literal — check F5)

| Uso | Formato |
|---|---|
| Cifras (miles, negativo en paréntesis, cero como guion) | `_-* #,##0_-;\(#,##0\)_-;_-* "-"_-;_-@_-` |
| Variante negativos rojos | `#,##0_);[Red](#,##0);-` |
| Percents | `0.0%` (drivers) / `0.00%` (tasas WACC) |
| Múltiplos | `0.0\x` → "12.3x" |
| USD por acción | `"$"#,##0.00_);\("$"#,##0.00\)` |
| Años | `0"A"` / `0"E"` |
| Fechas | `mm-dd-yy` |
| Celda auxiliar invisible | `;;;` |

## Continuidad histórico → forecast (check F11)

- **Una serie = una fila.** Cada línea del modelo (revenue, driver, margen,
  crecimiento) cruza TODO el horizonte en la misma fila: en columnas históricas
  la celda es fórmula calculada u observado (negro / azul-observado), en
  columnas forecast es input del analista (azul + amarillo) o fórmula de driver.
  El rol cambia por columna en la frontera A/E — el color cuenta la historia.
- Prohibido partir una serie en fila "histórico" y fila "forecast", y prohibido
  dejar columnas históricas vacías en filas de driver: el histórico calculado
  (ej. crecimiento % realizado) es la referencia visual contra la que el
  analista teclea su forecast (patrón CFI: fila "Revenue Growth %" con
  histórico calculado + forecast azul en la misma fila).
- El builder lo implementa con `ModelStyler.series_row(hist_values,
  forecast_values)` — roles por tramo, misma fila.

## Fórmulas

- Una fórmula por fila, copiada idéntica en todos los periodos (check S5) —
  entendido por tramo: el tramo histórico comparte una fórmula, el forecast otra.
- Jamás números hard-coded en fórmulas (check S4). Constantes de conversión
  (365, 1000) viven como celdas etiquetadas en Assumptions con named range
  (`DAYS_YEAR`, `MM_TO_B`) creado por `ModelStyler.define_constant`.
- Named ranges SOLO para valores clave cross-tab (precio de mercado, acciones,
  WACC) y constantes — nunca en el cálculo fila a fila (convención AMZN).
- Fila de años por fórmula (`=+J2+1`); trimestres con `EOMONTH`/`DATE`.
- Schedules como roll-forward: Opening / Plus / Less / Closing; **interés sobre
  saldo de apertura** (mata la circularidad sin switch de iteración).
- IFERROR solo para errores esperados (múltiplo "na" en comps/precedents) y en
  la celda de check — nunca para tapar errores reales.
- **Ventanas sin periodo previo** (growth yoy, UDM en los primeros trimestres):
  la celda queda VACÍA — ni texto ni 0 falso; F15 permite hasta 4 huecos
  INICIALES por fila (después del primer dato, un hueco = serie rota).
- **Roll-forward de caja — el DESFASE es el contrato** (check C10). Cuatro
  fórmulas, sin excepción, idénticas en toda la serie:
  1. `Efectivo al inicio(t) = Efectivo al cierre(t−1)` — referencia LITERAL a
     la celda de cierre de la columna previa; el primer periodo de forecast
     abre con el último cierre OBSERVADO.
  2. `Efectivo al cierre(t) = inicio(t) + Δ neto de efectivo(t)`.
  3. `BS caja(t) = CF cierre(t)` — la fila de caja del balance LEE el cierre
     del CF de la misma columna (tie-out C2); jamás una serie propia.
  4. Todo rendimiento/interés que dependa de caja, inversiones o deuda usa el
     saldo de APERTURA (= cierre t−1) — cierre de la misma columna crea el
     ciclo caja→otros ingresos→utilidad→caja y Excel deja la cadena entera
     sin calcular (smoke #5: el balance "cuadraba" solo donde había
     observados, y daba Error justo donde arrancaba el forecast).
  5. **El roll debe CERRAR también en el histórico** (check F19): si el CF
     omite una sección del flujo, el balance y el tie-out siguen verdes
     (ambos leen el mismo efectivo observado) mientras `inicio + cambio ≠
     cierre` en silencio. El CF necesita TODAS sus secciones — en emisoras
     con tesorería grande, el movimiento de valores negociables (compras,
     ventas, vencimientos) es el mayor flujo después del operativo.

## Estructura obligatoria del estado de flujos

Cada subtotal se construye como SUMA DE SUS COMPONENTES y debe cuadrar contra
el subtotal OBSERVADO de la fuente:

```
CFO = utilidad neta + D&A + SBC + ΔWC (desglosado) + otros ajustes operativos
CFI = −capex − compras de valores + ventas/vencimientos + otros de inversión
CFF = −recompras − dividendos + movimiento de deuda + otros de financiamiento
cambio neto = CFO + CFI + CFF
```

**Línea de conciliación derivada** (el "otros …" de cada subtotal): cuando la
captura no desglosa todas las partidas, una fila explícita absorbe la
diferencia contra el observado — `otros = subtotal_observado − Σ componentes`.
Es un plug LEGÍTIMO si y solo si: (i) vive en su propia fila visible, (ii)
lleva comentario que dice qué absorbe, y (iii) cierra contra un dato
observado, no contra un número deseado. Un plug escondido dentro de otra línea
es falsificación; uno declarado es contabilidad honesta.

Ancla del histórico: el efectivo del BALANCE (saldo puntual, sin
desacumulación) es el dato más confiable — la conciliación se calibra contra
él. Ojo con el alcance: el CF suele reconciliar *cash + restricted cash*
mientras el balance reporta solo *cash and equivalents*; esa diferencia es
parte de lo que la línea derivada absorbe.
- **Sin dato en una serie: JAMÁS texto** (`n/d`, `n/a`) en columnas de periodo
  — envenena SUMPRODUCT y toda aritmética aguas abajo (#VALUE!). El hueco es
  decisión del ANALISTA con gate: 0 explícito con comentario de celda,
  carry-forward del último dato disponible (fórmula `=<celda previa>`, marcada
  supuesto), o exclusión documentada de la fila. Check F15(b) lo audita.
- Sin volátiles (OFFSET, INDIRECT, NOW, TODAY — check S3). INDEX/MATCH sobre
  VLOOKUP. Anclas `$` deliberadas (data tables).

## Integridad y controles

- Check por columna en fila 3 de `Operating` (o `Model` en modo annual) (estilo AMZN:
  `=+IF(ABS(ref1-ref2)>0.0001,"Error","OK")`).
- Tab `Checks` con la lista S/C/D en fórmulas; celda única de error en Cover
  agregando todo.
- Cálculo automático, nunca manual (el builder lo fija).
- Sin referencias circulares.

## Bordes y legibilidad

- `top: thin` sobre subtotales; `bottom: double` en totales finales.
- Subtotales bold; totales bold.
- Filas de soporte agrupadas (outline) — la hoja colapsa a vista ejecutiva.
- Escenarios: una celda switch (INDEX/CHOOSE) en Assumptions; casos bull/base/
  bear documentados; áreas de escenario con fill gris.

## Workflow

- Versionado `_YYYY-MM-DD_v#`; nunca borrar columnas a media serie.
- Todo supuesto no obvio con comentario de celda.
- Summary de una pantalla con football field y 3-5 outputs clave.
- Antes de entregar: `python tools/xlsx_builder.py audit <modelo>` en verde +
  checks S/C/D verdes. Un check rojo = FALLA, no "casi cuadra".
