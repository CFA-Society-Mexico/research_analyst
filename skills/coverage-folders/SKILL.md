---
name: coverage-folders
description: Estructura de carpetas y archivado para una cobertura de equity — instancia el árbol estándar v2 por emisora (contratos en raíz, research/ fechado, journal/ append-only, macro/ a nivel workspace), aplica la convención de nombres, detecta filings faltantes, archiva documentos nuevos (10-K, 10-Q, 8-K, BMV, eventos relevantes, transcripts de earnings calls), renombra lo descargado por sec_fetch a la convención FY/#Q, y migra coberturas del árbol v1 al v2. Usa esta skill siempre que haya que crear la estructura de una cobertura nueva, archivar o renombrar cualquier documento, organizar carpetas de una emisora, migrar una cobertura vieja, o cuando el usuario diga "archiva este 10-Q", "organiza estos filings", "crea la carpeta de la emisora", "¿qué documentos me faltan?", "migra mi cobertura al árbol nuevo".
---

# coverage-folders

Dueña única del árbol de carpetas, el naming y el archivado. **Nunca lee el contenido
de los filings** (eso es de statement-mapper) — clasifica por tipo de documento y
metadatos, mueve y renombra. Las demás skills solo leen rutas.

## Cuándo corre

- `/init-coverage` paso 1: instanciar el árbol.
- `/update-quarter` paso 1: archivar el filing nuevo.
- Standalone: "archiva este 10-Q de AMX", "¿qué falta en mi carpeta?".

## Procedimiento — setup

1. Instancia el árbol de `templates/coverage-tree.md` bajo `<raiz>/<TICKER>/`,
   donde `<raiz>` es la carpeta que designó el usuario — jamás una
   subcarpeta intermedia como `workspace/` (ver coverage-tree).
   Colisión con cobertura existente ⇒ preguntar; NUNCA sobrescribir.
2. Clasifica los archivos sueltos que el usuario entregó: tipo (10-K/10-Q/8-K/
   BMV-annual/BMV-quarterly/evento relevante/**transcript de earnings call**),
   periodo, y renombra según la convención
   (`<TICKER>_<tipo>_<periodo>[_fecha].<ext>`). Transcripts van a
   `transcripts/`; investor day, presentaciones y docs de guidance del
   management a `filings/management/`; consenso y notas sell-side a
   `research/sell-side/<fuente>_YYYY-MM-DD.md` (fechados, se leen después de
   poblar drivers). Ambigüedad ⇒ preguntar.
3. Checklist de faltantes contra la **HISTORIA COMPLETA disponible** — el
   default del plugin es perspectiva de largo plazo: TODOS los 10-K y 10-Q (y
   equivalentes BMV) que existan, porque los drivers, márgenes y calibración se
   analizan a través de ciclos completos, no de un año. El "último anual +
   trimestrales del año" es el PISO para arrancar el pipeline, jamás el
   objetivo ni la propuesta de descarga: al ofrecer `sec_fetch`/`xbrl_fetch`,
   la propuesta default es historia completa y el gate del analista existe
   para RECORTAR, no al revés. Transcripts: opcionales, pero si existen se
   reportan como ENCONTRADOS — el pipeline los considera (guidance). Reporta
   encontrado / faltante — no bloquea, informa.
3b. `brand/` se crea siempre (vacía está bien). Si el usuario entregó un
   DESIGN.md o menciona colores de marca, va ahí — xlsx-building lo carga al
   construir el modelo.
4. Verifica que `<raiz>/macro/macro-view.yaml` existe (si no: crear la
   estructura `macro/{macro-view.yaml, sources/, history/}` desde template,
   vacía, y avisar). **Regla dura: lo macro vive SOLO en `<raiz>/macro/`** —
   jamás dentro del folder del ticker; si aparece un macro-view suelto ahí
   (cobertura vieja o error), tratarlo con el procedimiento de migración.

## Procedimiento — migración árbol v1 → v2

Para coberturas creadas con el árbol viejo (profile/ assumptions/ notes/ log/
earnings-transcripts/): aplicar el mapa de `templates/coverage-tree.md`
(§Migración) — solo movimientos, nada se borra; `macro-view.yaml` suelto (raíz
del workspace o dentro de un ticker) va a `<raiz>/macro/` (si ya existe uno
ahí, preguntar cuál manda);
carpetas viejas vacías se eliminan al final; reportar cada movimiento.

## Procedimiento — archivado (mantenimiento)

1. Identifica tipo y periodo del documento nuevo. Sin periodo identificable en
   nombre ni contenido visible ⇒ PREGUNTA — adivinar el periodo corrompe el
   retrieval de todo el pipeline (statement-mapper cita por nombre de archivo).
2. Anti-duplicado: ¿ya existe ese periodo? Corrección de la emisora ⇒ conservar
   ambos, nuevo con sufijo `_amended`.
3. Renombra, coloca, reporta ruta final.

**Ejemplo** — entrada suelta → destino:

| El usuario entrega | Queda como |
|---|---|
| `q3 apple.pdf` (10-Q del 3T fiscal 2026) | `filings/sec/10-Q/AAPL_10-Q_3Q2026.pdf` |
| `earnings call jul.pdf` | `transcripts/AAPL_transcript_3Q2026.pdf` |
| `nota MS sobre cloud.pdf` | `research/industry/sources/` (nombre original — sources no se renombran, se citan por archivo) |
| deck macro de la casa | `<raiz>/macro/sources/` |

## Procedimiento — renombrado desde sec_fetch

`tools/sec_fetch.py` baja crudo (`<TICKER>_<form>_<reportDate>.htm`) + manifest
CSV. Esta skill convierte al naming de la convención usando el manifest y el
`fiscal_year_end` del perfil: 10-K ⇒ `FYyyyy`; 10-Q ⇒ `#Qyyyy` (trimestre
FISCAL derivado del reportDate vs cierre fiscal, no calendario); 8-K ⇒ fecha.
Sin perfil aún (init paso 1): dejar el nombre crudo y renombrar al confirmarse
el perfil en el paso 2. El manifest se conserva junto a los filings.

## Reglas duras

- **Nada se borra, jamás.** Retención mínima 7 años (Standard V(C)). "Limpiar la
  carpeta" = reorganizar, nunca eliminar.
- Un archivo = un documento. Sin copias "final_v2".
- Esta skill no abre PDFs para leer estados financieros; si el usuario pide extraer
  cifras, eso es statement-mapper.

## Entrevista de cierre (solo en setup de cobertura nueva)

Objetivo (adaptativa, per `templates/debate-protocol.md`): sesgo de entrada. ¿Por qué
esta emisora? ¿Qué esperas encontrar? ¿Qué opinión traes ANTES de ver los números?
Primera entrada del thesis-journal — el valor está en documentar el prior para poder
contrastarlo después.
