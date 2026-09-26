---
name: statement-mapper
description: Captura de estados financieros desde filings a la capa de captura del modelo (model/inputs/ — CSVs canónicos + extracts JSON con cita) — extrae cada cifra con documento y página, la mapea a la línea estándar del model-spec, propone reclasificaciones, extrae el guidance cuantitativo de transcripts de earnings calls (obligatorio si están en transcripts/), y arma comp-snapshots fechados de comparables. Usa esta skill siempre que haya que capturar históricos de un filing (10-K, 10-Q, reporte BMV, press release de 8-K), mapear un estado de resultados/balance/flujos al modelo, conciliar una línea contra la fuente, extraer guidance de un call, armar el snapshot de un comparable, o cuando el usuario diga "captura este trimestre", "mapea este PDF al modelo", "¿de dónde salió esta cifra?", "¿qué guidance dio management?", "arma el snapshot de este comp".
---

# statement-mapper

Dueña única de la lectura de contenido de filings — de la emisora Y de comparables.
Su producto: cifras `observado` con cita exacta (documento, página/nota), mapeadas a
las líneas del `model-spec`. **No escribe en el xlsx** (entrega el mapeo a
model-standards) y **no mueve archivos** (eso es coverage-folders).

**Captura masiva vía XBRL (emisoras SEC):** `python tools/xbrl_fetch.py
<TICKER> --dest model/inputs --ua "<nombre correo>"` baja TODA la historia
reportada (anual y trimestral, décadas) desde companyfacts de SEC como CSV
largo con periodo fiscal, form y filed — sin parsear HTML. Esta skill convierte
ese CSV en los `canonical_*.csv` (mapeo canon → línea con gate del analista).
El tool ya entrega el 4Q de IS y CF derivado (FY − acumulado a 3Q, tag
`derivado`) — esta skill no resta nada a mano; la UPA de 4Q no se deriva
porque no es aditiva. Filas con tag `acumulado YTD sin desacumular` NO son
trimestrales: se reportan al analista y jamás entran al canónico como
trimestre. Preferir esta ruta a extraer trimestre por trimestre de 10-Qs; el
filing HTML queda para lo que XBRL no trae (notas, segmentos, guidance).

**Salida formal — `model/inputs/`** (dueña también de esta carpeta):
- `extract_*.json`: papeles de trabajo por filing — cada cifra con valor,
  documento, página/nota, fecha.
- `canonical_annual.csv` / `canonical_quarterly.csv` / `consolidated_long.csv`:
  el contrato máquina que consume model-standards (columnas mínimas en
  `templates/coverage-tree.md` §model/inputs). El histórico del xlsx se puebla
  desde estos CSV — jamás re-tecleado (check C9 verifica coincidencia).

## Cuándo corre

- `/init-coverage` paso 3: captura de históricos (en paralelo con industry-analysis).
- `/update-quarter` paso 2: captura del trimestre nuevo.
- Modo ligero: comp snapshots, cuando industry-analysis entrega el universo.
- Standalone: capturar un solo estado, conciliar una línea, un snapshot suelto.

## Procedimiento — captura completa

1. Lee el filing desde `filings/` (la ruta la da coverage-folders).
2. Extrae estado por estado. **Cada cifra lleva: valor, documento, página o nota,
   fecha.** Cifra sin cita no existe. Extrae con herramienta determinista sobre el
   texto del filing (parseo, no memoria): una cifra "recordada" es una cifra
   inventada aunque acierte.

   **Ejemplo — una fila del CSV canónico** (columnas en coverage-tree §model/inputs):

   ```
   line_item,period,value,source_doc,source_ref,tag
   total_net_sales,FY2025,391035,AAPL_10-K_FY2025.htm,"Consolidated Statements of Operations, p.28",observado
   ```
3. Mapea a las líneas del model-spec. La presentación de la emisora rara vez coincide
   1:1 — propone el mapeo y las reclasificaciones como PROPUESTA, con el racional.
4. Gate: el usuario aprueba mapeo y reclasificaciones antes de entregar a
   model-standards. Las reclasificaciones aprobadas se registran en `journal/decisions.md`.
5. Etiquetado en la fuente: cifras de estados = `observado`; proyecciones, metas y
   rangos que da la administración (guidance del press release, del call, del MD&A)
   = `guidance` — NUNCA se capturan como observado. Doctrina: del management,
   priorizar lo cuantitativo verificable.
5b. **Transcripts obligatorios si existen:** si `transcripts/` tiene el
   transcript del periodo capturado, leerlo NO es opcional — extraer todo
   guidance cuantitativo (rangos, metas, capex/margen/tax guiado, color por
   segmento) etiquetado `guidance` con cita (transcript, sección/página). Es la
   fuente de guidance más rica del periodo; alimenta el contraste de
   driver-inventory y el triage.
5c. **Materiales de management (`filings/management/`):** investor day,
   presentaciones a inversionistas y documentos de guidance se leen con la
   misma regla que los transcripts — todo lo cuantitativo etiquetado
   `guidance` con cita (documento, página/diapositiva, fecha); nunca
   `observado`. Notas sell-side y consenso NO son fuente de esta skill (van a
   `research/sell-side/` y los abre driver-inventory después de poblar).
6. Anti-doble-captura: verifica contra lo ya capturado antes de escribir.

## Procedimiento — comp snapshot (modo ligero)

1. Universo viene del industry-report vigente en `research/industry/`
   (industry-analysis lo decide, no tú).
2. Por comparable: llena `templates/comp-snapshot.yaml` y guárdalo como
   `comps/<COMP>_YYYY-MM-DD.yaml` — snapshot NUEVO fechado, jamás sobreescribir
   el anterior (vigente = fecha más reciente; la historia reconstruye el
   football field de cualquier fecha). Balance e income items del último filing
   del comp con cita; market data (precio, acciones) de fuente pública con
   fuente y **timestamp** (`as_of` con hora: los precios envejecen).
3. Anota el marco contable del comp y las notas de comparabilidad (consulta la
   reference de framework-mapper si el marco difiere del de la emisora — pero la
   cita normativa la emite framework-mapper, no tú).
4. Sin fuente pública disponible para un campo ⇒ campo vacío y reportado; jamás
   estimar un dato de mercado.

## Reglas duras

- Cifra sin cita = no capturada. Sin excepciones, ni "es obvia".
- Documento ilegible (PDF escaneado sin texto, tabla rota): reporta el archivo y
  las líneas afectadas como NO CAPTURADAS y sigue con el resto — jamás rellenes
  el hueco estimando ni con datos de otra fuente sin decírselo al usuario.
- No corregir la aritmética de la emisora en silencio: si la suma del filing no
  cuadra, se reporta como hallazgo (posible input de impact-triage), no se "arregla".
- Períodos re-expresados o re-presentados por la emisora: capturar la versión más
  reciente y anotar la re-presentación en `journal/decisions.md`.

## Entrevista de cierre

Objetivo (adaptativa, per `templates/debate-protocol.md`): calidad de earnings.
¿Qué sorprendió de los históricos? ¿Qué línea merece desconfianza? Generar preguntas
desde lo capturado (ej. "cuentas por cobrar crecen 2× revenue tres trimestres
seguidos — ¿qué hipótesis tienes?"). Registrar en thesis-journal.
