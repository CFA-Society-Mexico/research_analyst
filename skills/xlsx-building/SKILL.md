---
name: xlsx-building
description: Construcción determinista del xlsx del plugin — TODO workbook se arma vía tools/xlsx_builder.py (ModelStyler), jamás con openpyxl crudo; formato (paleta brandeable desde brand/DESIGN.md, fuentes, formatos numéricos, gridlines, freeze, outline, bordes, series continuas hist→forecast) sale del código y se verifica con los checks F (F1-F20); incluye el modo REBUILD para reconstruir cualquier modelo existente al estándar con audit previo y paridad de números obligatoria. Usa esta skill siempre que haya que crear o editar un archivo Excel del plugin, aplicar o corregir formato de un modelo, reconstruir un modelo viejo o ajeno, correr el audit de formato, aplicar colores de marca, o cuando el usuario diga "construye el xlsx", "formatea el modelo", "el modelo salió feo", "reconstrúyelo al estándar", "corre el audit de formato" — model-standards la invoca en su paso de construcción y es obligatoria: si un xlsx se va a escribir y esta skill no está en uso, detente y cárgala.
---

# xlsx-building

Capa determinista de construcción. Regla única: **ningún workbook del plugin se
escribe sin `tools/xlsx_builder.py`**. El módulo es la implementación; esta
skill es el procedimiento; los checks F son la verificación. Prosa no construye
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
2. **Secciones**: el modelo vive en la hoja única `Model` — secciones
   Assumptions → IS → BS → CF → DCF → Ratios → Schedules con `section_header`
   (banda) y `group_rows` (colapsables). Schedules: un bloque por schedule con
   `schedule_block_header` ("Sch: X") — JAMÁS una hoja por schedule ni hojas
   IS/BS/CF sueltas. Periodicidad mixta: `quarter_header` para el tramo
   trimestral estimado antes de `period_header` anual; el anual corriente =
   suma de sus trimestres por fórmula (C8 estructural).
2b. **Ratios por código**: la sección Ratios se genera con
   `ModelStyler.build_ratios(ws, start_row, first_col, n_cols, ref, wacc_ref)`
   — `ref` es el registro canon→referencia de fila que el build ya conoce
   (viene de poblar los estados desde los CSVs canónicos). PROHIBIDO armar
   Ratios a mano: el check F13 exige el set completo y `build_ratios` lo
   escribe por construcción. Si `build_ratios` reporta canons faltantes, eso
   es un hueco del mapeo de captura — repórtalo, no lo tapes.
3. **Contenido**: toda celda vía `set_cell` con su `CellRole` (INPUT / OBSERVED /
   FORMULA / LINK / WARN / LABEL) y `NumFmt` de la whitelist. El role fija color
   y fill — no elijas colores. Series (drivers, líneas de estados): vía
   `series_row` — histórico calculado y forecast en LA MISMA fila (check F11),
   jamás filas hist/forecast separadas ni columnas históricas vacías.
4. **Constantes**: 365, 1000 y similares como celdas etiquetadas en Assumptions
   con `define_constant` (`DAYS_YEAR`, `MM_TO_B`); las fórmulas referencian el
   named range, nunca el literal (check S4).
5. **Checks arriba**: `check_row` en fila 3 de IS/BS/CF (balance / tie-out por
   columna, estilo `=+IF(ABS(a-b)>0.0001,"Error","OK")`).
6. **Bordes**: `subtotal_border` / `total_border` — no bordes manuales.

**Ejemplo mínimo del API** (ancla la forma correcta de usarlo):

```python
styler = ModelStyler(brand=load_brand("brand/DESIGN.md"))  # brand opcional
ws = styler.new_sheet("IS", freeze="C4")
styler.brand_bar(ws, "Estado de resultados")
styler.period_header(ws, 3, 3, PeriodHeader(2019, 2031, 2025))  # 2025A|2026E
styler.series_row(ws, 7, "Crecimiento de ventas (%)", first_col=3,
                  hist_values=["=D6/C6-1", ...],   # calculado, negro
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
   PENDIENTE (p. ej. F19 sin valores calculados): recalcula con Excel COM y
   re-audita; sin Excel disponible se reporta pendiente, nunca verde. Nunca
   entregues con audit rojo. Luego corren S/C/D (integrity-checks.md).

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
  CADA escritura con openpyxl, recalcular con Excel COM**
  (`CalculateFullRebuild` + `Save`) antes de auditar o leer valores.
- **Para INSERTAR o BORRAR filas usa Excel COM, jamás openpyxl.** COM ajusta
  todas las fórmulas del libro al desplazar; openpyxl no ajusta nada y deja
  cientos de referencias apuntando a la fila equivocada — corrupción
  silenciosa que ningún check de formato detecta.
- Tras insertar filas, **re-verifica un output clave** (valor por acción,
  balance check): un salto inexplicado significa referencias rotas.

## Qué NO hace esta skill

- No decide contenido ni supuestos (eso es driver-inventory / el analista).
- No inventa estilos: si necesitas un formato que no está en la whitelist, la
  respuesta es proponer extender `xlsx_builder.py` + checks F con el usuario —
  no aplicarlo ad hoc.
- No toca modelos externos sin instrucción (audit sí — es solo lectura).

## Referencias

- `tools/xlsx_builder.py` — implementación única (paleta, NumFmt, roles, audit F).
- `skills/model-standards/references/excel-practices.md` — el contrato legible.
- `skills/model-standards/references/integrity-checks.md` — checks F en la lista.
