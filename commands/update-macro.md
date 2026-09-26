# /update-macro — actualizar el house view macro desde las fuentes del analista

Orquesta la actualización de `<raiz>/macro/macro-view.yaml` a partir del
research que el analista dejó en `<raiz>/macro/sources/` (propio, de su casa
o de terceros; pdf/html/md). Uso: `/update-macro [contexto del porqué]`.

Doctrina: el plugin PROPONE campo por campo con cita; el analista confirma cada
valor. El plugin jamás decide un supuesto macro.

## Preflight (cualquier fallo DETIENE el comando)

1. `macro/` existe con `macro-view.yaml` (si no: crear estructura desde
   template y avisar — no hay nada que actualizar todavía).
2. **Series frescas:** correr `python tools/fred_fetch.py --dest macro/series`
   (gate: avisar antes). Sin key: pedirla en el chat (el registro es gratuito,
   link en el mensaje del tool) y guardarla en `macro/fred.key` — jamás
   imprimirla. Sin red: el tool da las opciones; continuar con las series ya
   descargadas si existen.
3. Insumos: `macro/series/` (observados FRED) y/o `macro/sources/` (research
   del analista/terceros). Sin NINGUNO de los dos ni instrucción explícita,
   no hay base para proponer.
4. **Snapshot primero:** copiar el yaml vigente a
   `macro/history/macro-view_<YYYY-MM-DD>.yaml` ANTES de tocar nada.

## Pipeline

| Paso | Qué hace |
|---|---|
| 1 | Lee `macro/series/` (observados FRED: rf de la última obs de DGS10, FX de DEXMXUS, etc.) y `macro/sources/` completo. Regla de citado: serie ⇒ `observado` con serie y fecha de observación; dato duro de documento ⇒ `observado` con documento y página; opinión de tercero ⇒ atribuida, jamás supuesto sin gate |
| 2 | **Propuesta campo por campo**: tabla valor vigente → valor propuesto → fuente exacta → etiqueta. Solo campos con evidencia en las fuentes; el resto no se toca |
| 3 | **Gate POR CAMPO**: el analista confirma, corrige o rechaza cada propuesta. Sin confirmación, el campo queda como estaba |
| 4 | Escribe el yaml con los campos confirmados + `updated_at` de hoy |
| 5 | `impact-triage` **transversal**: un cambio macro pega a TODAS las coberturas del workspace — triage por emisora (rf/ERP ⇒ valuación; FX ⇒ emisoras expuestas; inflación ⇒ monitor B-10 de emisoras NIF) |

## Post-run

- Snapshot en `history/` + yaml vigente actualizado.
- Tabla de triage transversal: emisora → qué le pega → nivel.
- Los cambios confirmados y su porqué quedan en el `journal/decisions.md` de
  cada cobertura afectada (vía el triage), no en un log global.

## Reporte final

Una pantalla: campos cambiados (antes → después, con fuente), campos rechazados,
y el triage transversal por emisora.
