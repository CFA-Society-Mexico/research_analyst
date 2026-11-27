---
name: xlsx-building
description: Construcción determinista del xlsx del plugin — TODO workbook se arma vía tools/xlsx_builder.py (ModelStyler), jamás con openpyxl crudo; formato (paleta brandeable desde brand/DESIGN.md, fuentes, formatos numéricos, gridlines, freeze, outline, bordes, series continuas hist→forecast) sale del código y se verifica con los checks F (F1-F21); incluye el modo REBUILD para reconstruir cualquier modelo existente al estándar con audit previo y paridad de números obligatoria. Usa esta skill siempre que haya que crear o editar un archivo Excel del plugin, aplicar o corregir formato de un modelo, reconstruir un modelo viejo o ajeno, correr el audit de formato, aplicar colores de marca, o cuando el usuario diga "construye el xlsx", "formatea el modelo", "el modelo salió feo", "reconstrúyelo al estándar", "corre el audit de formato" — model-standards la invoca en su paso de construcción y es obligatoria: si un xlsx se va a escribir y esta skill no está en uso, detente y cárgala.
---

# xlsx-building

Capa determinista de construcción. Regla única: **ningún workbook del plugin se
escribe sin `tools/xlsx_builder.py`** (fachada sobre `tools/xlsx_*.py`). El
módulo es la implementación; esta skill es el procedimiento; los checks F son
la verificación. Prosa no construye
formato — código sí (lección del smoke test AAPL 2026-08-30: el agente tenía
las best practices escritas y entregó Calibri con gridlines).

## Procedimiento

0. **Marca**: si existe `brand/DESIGN.md` en la cobertura, cárgalo con
   `load_brand(path)` y pásalo a `ModelStyler(brand=...)` — rebrandea SOLO los
   3 slots decorativos (primary/section/accent). Los colores semánticos (azul
   input, verde link, rojo warn, fills) jamás cambian. El audit se corre con el
   mismo archivo: `python tools/xlsx_builder.py audit <modelo> brand/DESIGN.md`.
1. **Scaffold**: instancia `ModelStyler` (fija calc auto + sello F10) y
   ESTAMPA la periodicidad del perfil con `set_periodicity(model_periodicity)`
   — obligatorio: sin sello, F14 no puede vigilar las columnas trimestrales.
   Crea cada tab con `new_sheet` (gridlines off + freeze), `brand_bar`,
   `label_col_width`. Modo `quarterly` (spec §Operating/§Annual): DOS hojas —
   `Operating` con header 100% trimestral (`quarter_header`, `1Q2016A …
   4Q2031E`, sin columnas FY) y `Annual` con header de años (`period_header`
   A/E) donde TODA celda es fórmula de agregación leyendo Operating (flujos
   Σ4Q, stocks 4Q, ratios recalculados) o link a canonical_annual — cero
   inputs (F14 lo audita). La sección DCF vive en `Annual`, línea por línea
   según el bloque del model-spec — cero fórmulas comprimidas; los bloques de
   valor único (terminal/puente/cruces/reverse/Hamada) anclan sus VALORES en
   la columna C, pegados a los labels — jamás al final del horizonte. Toda
   fila de serie llena TODOS sus periodos (F15): un tramo sin fórmula = FCFF
   chatarra en silencio. Modo `annual`: una hoja `Model` con `period_header`.
1b. **Agrupación TOTAL**: tras escribir cada sección (en Operating Y Annual),
   `group_rows` sobre TODO su contenido — sin excepciones. F7 audita por
   sección: una sola sección sin outline = falla.
1c. **Ratios una sola vez**: `build_ratios` se llama UNA vez por hoja sobre el
   horizonte completo — histórico y forecast en la misma fila. Secciones
   "Ratios histórico"/"Ratios forecast" separadas = F13 falla por duplicado.
2. **Secciones**: en modo `quarterly`, `Operating` lleva Assumptions → IS →
   BS → CF → Ratios → Schedules y `Annual` lleva IS → BS → CF → Ratios →
   Schedules agregados → DCF; en modo `annual`, la hoja `Model` lleva las
   mismas secciones con el DCF incluido. Cada sección con `section_header`
   (banda) y `group_rows` (colapsables). Schedules: un bloque por schedule con
   `schedule_block_header` ("Sch: X") — JAMÁS una hoja por schedule ni hojas
   IS/BS/CF sueltas. Las columnas FY intercaladas entre trimestres
   (`interleaved_header`) están deprecadas desde el diseño 2026-08-31.
2b. **Ratios por código**: la sección Ratios se genera con
   `ModelStyler.build_ratios(ws, start_row, first_col, n_cols, ref, wacc_ref)`
   — `ref` es el registro canon→referencia de fila que el build ya conoce
   (viene de poblar los estados desde los CSVs canónicos). PROHIBIDO armar
   Ratios a mano: el check F13 exige el set completo y `build_ratios` lo
   escribe por construcción. Si `build_ratios` reporta canons faltantes, eso
   es un hueco del mapeo de captura — repórtalo, no lo tapes. En `Operating`
   (header trimestral) pasa `window` SIEMPRE: `window=4` si `ref` apunta a
   flujos de un trimestre, `window=1` si ya apunta a filas UDM; sin él,
   `build_ratios` falla a propósito.
3. **Contenido**: toda celda vía `set_cell` con su `CellRole` (INPUT / OBSERVED /
   FORMULA / LINK / WARN / LABEL) y `NumFmt` de la whitelist. El role fija color
   y fill — no elijas colores. Series (drivers, líneas de estados): vía
   `series_row` — histórico calculado y forecast en LA MISMA fila (check F11),
   jamás filas hist/forecast separadas ni columnas históricas vacías.
   Toda fila con histórico OBSERVADO se vincula a su línea de la capa de
   captura: `styler.link_capture(ws, fila, "bs_ap", scale=1_000_000)`
   (`valor_csv = valor_xlsx × scale`; el histórico va sin redondear). Sin
   vínculo, C9 trata sus cifras como no respaldadas. Dos reglas del vínculo:
   (a) ventas totales (`is_ns_total`) y activos totales (`bs_ta`) se vinculan
   SIEMPRE, aunque sus filas sean fórmulas: declaran la unidad monetaria con
   la que C9 mide una cifra sin respaldo; (b) una línea respalda UNA fila por
   hoja — si la misma cifra aparece dos veces (caja del balance y cierre del
   flujo), una fila la teclea y la otra la lee por fórmula.
3b. **Año de transición** (modo `annual`, primer año del forecast con
   trimestres ya observados): `set_transition_year(<año>)` una vez y
   `series_row(..., transition_first="=<ancla>")` en las series ancladas
   (9M observado + 4T guiado). `transition_first` sustituye a
   `forecast_values[0]` y se escribe como fórmula negra: en ese año el driver
   es RESULTADO del ancla. Sin el sello el builder lo rechaza y S5 falla.
3c. **Fórmulas que Excel acepta (F21):** `set_cell` rechaza al construir una
   aridad inválida (`ABS(a,b,c)`), un paréntesis sin cerrar, una función
   moderna sin `_xlfn.` o una operación de arreglo en fórmula normal. Una
   operación sobre rangos va en `SUMPRODUCT` o con `set_array_formula`
   (`MEDIAN(IF(rango=1, valores))`). No rodees el rechazo escribiendo la celda
   con openpyxl crudo: Excel borraría la fórmula al abrir.
3d. **Valuación por código** (fórmulas y guardas en
   `skills/model-standards/references/valuation-conventions.md`):
   `gordon_terminal_block` (flujo terminal normalizado con RONIC),
   `implied_growth_block` (g implícita del exit y del reverse DCF, con la
   guarda de la raíz espuria), `sensitivity_grid` + `dcf_per_share_formula`
   (cuadrícula con fórmula cerrada por celda — openpyxl NO crea data tables),
   `center_check_formula`, `comps_market_cap_formula`,
   `multiple_range_flag_formula` y `terminal_multiple_alarm_formula`.
   DDM y FCFE: `equity_flow_rows` (dividendos y FCFE desde el modelo, pagos
   restados), `equity_terminal_block` (terminal de equity normalizado con ROE
   terminal), `equity_value_block` (valor del equity y por acción, con stub
   del año en curso por `stub_flow_ref` / `stub_exponent_ref`),
   `method_divergence_formula` (check D5) y `method_consistency_block`
   (diagnóstico de D5).
   PROHIBIDO re-derivar estas fórmulas en el script de la cobertura.
4. **Constantes**: 365, 1000 y similares como celdas etiquetadas en Assumptions
   con `define_constant` (`DAYS_YEAR`, `MM_TO_B`); las fórmulas referencian el
   named range, nunca el literal (check S4).
5. **Checks arriba**: `check_row` en fila 3 de `Operating`, o de `Model` en
   modo annual (balance / tie-out por
   columna, estilo `=+IF(ABS(a-b)>0.0001,"Error","OK")`).
6. **Bordes**: `subtotal_border` / `total_border` — no bordes manuales.

**Ejemplo mínimo del API**, modo `annual` (ancla la forma correcta de usarlo;
el libro v3 completo en modo `quarterly` está en `_demo` de `tools/xlsx_builder.py`):

```python
styler = ModelStyler(brand=load_brand("brand/DESIGN.md"))  # brand opcional
ws = styler.new_sheet("Model", freeze="C4")
styler.brand_bar(ws, "Modelo anual")
styler.period_header(ws, 3, 3, PeriodHeader(2019, 2031, 2025))  # 2025A|2026E
styler.series_row(ws, 7, "Crecimiento de ventas (%)", first_col=3,
                  hist_values=[None, "=D6/C6-1", ...],  # C vacía: sin año previo
                  forecast_values=[0.05, ...],      # input, azul+amarillo
                  numfmt=NumFmt.PCT1)
styler.total_border(ws, 26, 3, 13)
styler.save(path)  # luego: python tools/xlsx_builder.py audit <path>
```
7. **Verificación obligatoria antes de entregar**:

   ```
   python tools/xlsx_builder.py audit <modelo.xlsx>
   ```

   Exit 0 = formato verde. Exit 1 = FALLA: corrige y re-corre. Exit 3 =
   PENDIENTE (F19 o checks vivos SIN evaluar por falta de valores
   calculados). Nunca entregues con audit rojo. Luego corren S/C/D
   (integrity-checks.md; los S y el C10 con `python tools/model_check.py
   <modelo.xlsx>`, que re-corre los F).
8. **Recalcular y validar con un motor real (obligatorio antes de entregar):**

   ```
   python tools/model_check.py <modelo.xlsx> --recalc
   ```

   Recalcula una COPIA con Excel (COM) o LibreOffice y audita esa copia: R1
   (Excel acepta el libro — una fórmula que el motor descarta es falla), R2
   (sin celdas en error) y los checks vivos de `Checks`/`Cover` ya evaluados.
   Un F en verde NO prueba que Excel acepte el libro. Para dejar los valores
   calculados en el modelo: `python tools/recalc_validate.py <modelo.xlsx>
   --in-place` (solo escribe si R1 está verde). Sin motor disponible queda
   `[pendiente]: checks vivos SIN evaluar` — se reporta así, nunca verde.

## Modo rebuild (modelo existente → modelo al estándar)

Cuando el usuario entrega un xlsx (del plugin viejo o ajeno) y pide
reconstruirlo, el orden es fijo:

1. **Audit primero**: corre la lista completa S/C/D/F sobre el modelo viejo
   (`integrity-checks.md` + `python tools/xlsx_builder.py audit`). Presenta la
   tabla de violaciones — QUÉ viola y POR QUÉ se va a reconstruir.
2. **Gate**: mostrar al usuario qué se PRESERVA (valores, etiquetas de fila,
   lógica de fórmulas, supuestos) y qué se DESCARTA (todo el formato, tabs
   `Sch_*`, series partidas). Sin aprobación no hay rebuild.
3. **Rebuild**: el modelo viejo es FUENTE DE CONTENIDO, jamás de formato NI de
   forma — extrae valores/labels/fórmulas con openpyxl y reconstruye vía
   `ModelStyler` contra el model-spec vigente. Versión nueva `_YYYY-MM-DD_v#`;
   el original queda intacto (nada se borra).
3b. **Normalización de series (obligatoria)**: transcribir ≠ copiar. Series
   partidas del original (fila "histórico" + fila "forecast" de la misma
   métrica) se FUSIONAN en una fila (checks F11/F12); valores derivables hacia
   atrás (índices, ASP implícitos, ratios) se poblan por fórmula en el tramo
   histórico, jamás quedan vacíos. El patrón viejo no se hereda: la lección del
   re-smoke AAPL v2 fue transcribir demasiado literal.
4. **Re-audit + paridad (obligatorio)**: checks F en verde Y check de paridad —
   los outputs clave del viejo y el nuevo coinciden (balance check, revenue por
   periodo, resultado de valuación). Divergencia = FALLA con celdas afectadas:
   reconstruir jamás cambia números en silencio. Diferencias INTENCIONALES
   (ej. corregir un hardcode) se listan explícitas en el reporte.

## Mecánica de las herramientas (aprendido en producción, no negociable)

- **openpyxl BORRA la caché de valores calculados al guardar.** Todo `save()`
  de openpyxl deja el libro con fórmulas pero sin resultados: si auditas sin
  recalcular, S10/F19 reportan miles de falsos "sin calcular". **Después de
  CADA escritura con openpyxl, recalcular** antes de auditar o leer valores:
  `python tools/recalc_validate.py <modelo.xlsx> --in-place` (Excel COM o
  LibreOffice; no sobreescribe si el motor descartó fórmulas).
- **Cargar el builder desde un script de cobertura:** `sys.path.insert(0,
  "<plugin>/tools")` + `import xlsx_builder`, o por ruta con
  `importlib.util.spec_from_file_location` + `exec_module` — ambas funcionan
  sin registrar el módulo en `sys.modules` (antes la carga por ruta rompía las
  dataclasses). Importa siempre la fachada `xlsx_builder`, nunca un
  `xlsx_*` suelto; los `xlsx_*.py` deben vivir junto a ella en `tools/`.
- **Para INSERTAR o BORRAR filas usa Excel COM, jamás openpyxl.** COM ajusta
  todas las fórmulas del libro al desplazar; openpyxl no ajusta nada y deja
  cientos de referencias apuntando a la fila equivocada — corrupción
  silenciosa que ningún check de formato detecta.
- Tras insertar filas, **re-verifica un output clave** (valor por acción,
  balance check): un salto inexplicado significa referencias rotas.

## Qué NO hace esta skill

- No decide contenido ni supuestos (eso es driver-inventory / el analista).
- No inventa estilos: si necesitas un formato que no está en la whitelist, la
  respuesta es proponer extender el contrato (`tools/xlsx_contract.py`) + checks F
  (`tools/xlsx_audit.py`) con el usuario —
  no aplicarlo ad hoc.
- No toca modelos externos sin instrucción (audit sí — es solo lectura).

## Referencias

- `tools/xlsx_builder.py` — implementación única: la fachada que se importa y el
  CLI (`audit`, `demo`). El código vive en `tools/xlsx_contract.py` (paleta,
  NumFmt, roles), `tools/xlsx_formulas.py` (lint F21, fórmulas de valuación),
  `tools/xlsx_styler.py` (`ModelStyler`, bloques de valuación, Ratios),
  `tools/xlsx_scan.py` (lectores del libro) y `tools/xlsx_audit.py` (checks F).
- `tools/recalc_validate.py` — recálculo con Excel/LibreOffice y checks R.
- `skills/model-standards/references/excel-practices.md` — el contrato legible.
- `skills/model-standards/references/integrity-checks.md` — checks F en la lista.
