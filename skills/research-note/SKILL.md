---
name: research-note
description: Presentación de la tesis de una cobertura (etapa 7) — arma el deck de inicio o de actualización desde el modelo ya construido, con cada cifra atada a su celda o a un hecho citado; propone el guion por actos, registra los hechos externos con su cita, escribe el script research/notes/build_deck.py sobre tools/deck_builder.py y cierra con la verificación de cifras, espacio, afirmaciones y capturas. Usa esta skill siempre que haya que presentar una cobertura, armar o actualizar un deck, revisar si un deck sigue vigente contra el modelo, rastrear una cifra de una lámina, o cuando el usuario diga "arma el deck", "presenta la cobertura", "prepara la presentación de la tesis", "actualiza el deck con el trimestre", "¿de dónde salió esta cifra del deck?", "verifica el deck". Es la ÚNICA skill que escribe en research/notes/.
---

# research-note

Dueña única de `research/notes/` y de la etapa 7 del flujo: comunicar la tesis.
Dos formatos por diseño; hoy existe el **deck** (HTML autocontenido, láminas de
1280×720, imprimible a PDF). La nota .docx queda pendiente y compartirá el
registro de hechos y la verificación.

**En un deck un número no es texto.** Sale del modelo recalculado, de un hecho
citado o de aritmética en código entre ambos. Esta skill jamás teclea, redondea
ni estima una cifra: una sin origen impide generar el deck.

## Cuándo corre

- `/present-coverage` pasos 1 a 4: guion, hechos, script y verificación.
- `/update-quarter` post-run: solo `audit` del deck vigente contra el modelo
  nuevo. No regenera.
- Standalone: auditar un deck existente, rastrear una cifra hasta su celda,
  regenerar tras un cambio del modelo.

## Inputs (solo lectura — esta skill no escribe en ninguno)

| Input | De quién | Para qué |
|---|---|---|
| El modelo vigente de `model/` (fecha y versión más recientes), recalculado | `model-standards` | toda cifra del deck |
| `driver-map.md` (raíz) | `driver-inventory` | la identidad de ventas y los drivers del acto de crecimiento |
| `research/industry/industry-report_<fecha>.md` vigente | `industry-analysis` | contexto, posicionamiento, fuente de hechos |
| `journal/thesis-journal.md`, `journal/decisions.md` | gates previos | la tesis y sus contra-argumentos; supuestos con su porqué |
| `issuer-profile.yaml` (raíz) | `framework-mapper` | marco, moneda, métodos de valuación activos |
| `brand/DESIGN.md` | analista | slots `deck_*`; sin archivo, estilo neutro |

Falta el modelo o no tiene valores calculados → detente y repórtalo. El deck no
se arma sobre un modelo sin recalcular.

## Artefactos (`research/notes/`; vigente = fecha más reciente; nada se borra)

| Archivo | Qué es |
|---|---|
| `<ticker>-deck-outline_YYYY-MM-DD.md` | Guion aprobado en el gate 1 |
| `<ticker>-facts_YYYY-MM-DD.json` | Hechos externos citados |
| `build_deck.py` | Script vigente de la cobertura, escrito sobre el tool |
| `build_deck_YYYY-MM-DD.py` | Copia del script anterior, antes de reescribirlo |
| `<ticker>-deck_YYYY-MM-DD.html` | El deck |
| `<ticker>-deck-verify_YYYY-MM-DD.md` | Reporte de verificación del gate 2 |
| `shots_YYYY-MM-DD/` | Una captura por lámina |

## Procedimiento

Lee `references/deck-format.md` antes de empezar: trae el formato de cada
artefacto y la API real del tool. No supongas una firma; está ahí y en el
docstring de `tools/deck_builder.py`.

1. **Guion** (referencia §1). Lee el modelo, el driver-map, el industry report y
   el journal. Propón la lista de láminas: cada una con su hallazgo, su visual y
   las series, celdas y hechos que necesita. El orden es fijo (qué es, qué la
   mueve, qué tan rentable es, cuánto vale); los actos se adaptan a la emisora.
   Ninguna cifra se escribe en el guion: se nombra por su origen. Escribe el
   outline y detente: gate 1.
2. **Hechos** (§2). Cada dato del guion que no esté en el modelo entra al
   registro con documento, localizador y cita literal copiada del texto visible
   de la fuente. Sin fuente en la cobertura: el dato sale del guion o el
   analista aporta el documento. Nunca de memoria. Valida el registro con
   `deck_builder.py facts` antes de escribir láminas; una cita que no aparece
   se copia del volcado de `facts --text`.
3. **Script y build** (§3 a §5). Si ya existe `build_deck.py`, guarda la copia
   fechada primero. Vuelca las etiquetas del modelo con `deck_builder.py
   labels`: de ahí salen la etiqueta exacta, la sección y la columna de cada
   cifra. Escribe el script sobre el tool y córrelo. El bloqueo de
   cifras aplica aquí: un número suelto detiene el build y el mensaje dice la
   lámina. Se corrige el origen de la cifra, no el texto para esquivar la regla.
   `lit()` es para conteos editoriales y nombres propios, jamás para una cifra
   del negocio. Después: construir, medir, ajustar; dos o tres rondas es lo
   normal. La medición no decide el contenido: si para pasarla hace falta un
   visual o un hallazgo que el guion no tenía, se anota en el outline y va como
   excepción al gate 2; si cambia el argumento, vuelve al analista.
4. **Verificación** (§6 y §7). `audit --report` y `measure --shots --report`
   hasta verde, sobre el mismo reporte. Luego evidencia fila por fila de la
   sección "Afirmaciones", que ya lista prosa, tablas, portada y rótulos de
   gráfica (lo que no la tenga se elimina y se regenera; la evidencia de las
   frases que no cambian se conserva entre corridas): con subagentes
   disponibles la hace uno que no redactó el deck. Después, revisión de las capturas **una por una** y
   pasada de texto. La medición en verde no sustituye mirar las capturas.
   Entrega al gate 2 el reporte, las capturas y la lista de excepciones (usos
   de `lit()`, hechos con verificación manual, frases eliminadas, cambios
   posteriores al gate 1, lo que quedó sin documento primario).

## Reglas duras

- Sin rating, recomendación ni precio objetivo, salvo que el analista los dicte
  textualmente. El valor que da el modelo se presenta como salida del modelo.
- Atribución y aviso legal genéricos. Marca institucional solo si el analista
  la pide y aporta su `brand/DESIGN.md`; el repo no trae ninguna.
- Audiencia con conocimiento promedio de finanzas: se explica el negocio antes
  de valuarlo y cada término técnico la primera vez que aparece.
- Las cuatro etiquetas siguen vivas en el texto: lo proyectado se dice como
  supuesto del modelo, el guidance se atribuye a la dirección, lo observado
  lleva fuente.
- `[pendiente]` (sin navegador para medir) nunca se reporta como verde.
- Mismo `as_of`: se regenera cuantas veces haga falta hasta el gate 2. Aprobado
  el gate 2, un deck nuevo lleva `as_of` nuevo; el aprobado no se sobrescribe.
- No editar el HTML a mano: `audit` lo detecta (P1) y el deck deja de ser
  reproducible. Todo cambio pasa por el script.

## Entrevista de cierre

Dos gates, ambos con `templates/debate-protocol.md` y entrada en thesis-journal.

- **Gate 1, guion.** Objetivo: **¿el guion cuenta la tesis o recita el modelo?**
  Sondear qué lámina sobra, qué contra-argumento del journal no aparece, qué
  acto presupone que la audiencia ya conoce a la emisora, y si el orden lleva a
  la conclusión del analista o a la del asistente.
- **Gate 2, entrega.** Objetivo: **¿qué frase de este deck no firmaría el
  analista?** Se recorren las excepciones (cada `lit()`, cada hecho manual, cada
  frase eliminada) y las capturas. El analista aprueba la entrega; el asistente
  no decide que está listo.

## Referencias

- `references/deck-format.md` — guion, registro de hechos, API del tool, regla
  de bloqueo, checks P1 a P10 y verificación: leer antes de escribir el outline.
- `tools/deck_builder.py` — implementación única (`demo`, `labels`, `facts`,
  `audit`, `measure`, `claims`); su docstring y sus firmas mandan sobre
  cualquier resumen.
- `tools/recalc_validate.py` — recálculo del modelo cuando no trae valores.
