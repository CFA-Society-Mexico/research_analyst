# /present-coverage — deck de la tesis de una cobertura

Orquesta la etapa 7: guion → hechos citados → script y build → verificación, con
dos gates del analista. El deck es opcional y se regenera cuando el modelo
cambia; por eso es comando propio y no un paso de `/init-coverage`.
Uso: `/present-coverage <TICKER> [contexto: audiencia, inicio o actualización]`.

## Preflight (cualquier fallo DETIENE el comando)

1. Árbol de cobertura existe y hay un modelo vigente en `model/` (fecha y
   versión más recientes).
2. **Modelo verde con recálculo real:** `python tools/model_check.py <modelo>
   --driver-map <raiz>/<TICKER>/driver-map.md --recalc` termina sin `[x]` y sin
   checks vivos `[pendiente]`. Roto ⇒ `/model-check` primero; no se presenta un
   modelo que no cuadra. Vale también para un modelo que estaba en verde con
   una versión anterior del plugin y hoy falla un check nuevo: el deck hereda
   cada cifra del modelo. No hay excepción firmada; se arregla el modelo.
3. **El modelo EN DISCO trae valores calculados.** El recálculo del punto 2
   corre sobre una copia; el deck lee el archivo de la cobertura. Comprobación:
   la misma corrida SIN `--recalc`. Si S10, F19 o los checks vivos salen
   `[pendiente]` ("fórmulas sin valor calculado"), el archivo no trae valores
   (pasa tras cualquier escritura con openpyxl). Entonces, con el sí del
   analista: `python tools/recalc_validate.py <modelo> --in-place` (guarda los
   valores; solo si el motor conservó todas las fórmulas).
4. Industry report vigente en `research/industry/`. Vencido (D10): `[aviso]`,
   no bloquea.
5. Navegador para medir (Edge, Chrome o Chromium). Si falta: `[aviso]` ahora y
   el deck saldrá `[pendiente]`, nunca verde.

La valuación no es preflight: un guion sin acto de valuación es válido. Si el
guion la incluye y la celda no existe, el tool falla al leerla.

## Pipeline

| Paso | Skill | Qué hace |
|---|---|---|
| 1 | `research-note` | Guion: lista de láminas, cada una con su hallazgo, tipo de visual y las series, celdas y hechos que necesita; actos adaptados a la emisora. Si hay un outline anterior, sirve de base. Escribe `research/notes/<ticker>-deck-outline_<YYYY-MM-DD>.md` |
| Gate 1 | — | El analista aprueba, recorta o reordena el guion. Debate per `templates/debate-protocol.md`; entrada en thesis-journal |
| 2 | `research-note` | Registra cada dato externo del guion en `<ticker>-facts_<YYYY-MM-DD>.json`, con documento, localizador y cita. Sin fuente: sale del guion o el analista aporta el documento |
| 3 | `research-note` | Si existe `build_deck.py`, guarda la copia fechada. Escribe el script sobre `tools/deck_builder.py` y lo corre; el bloqueo de cifras aplica aquí. Ajusta con la medición hasta que no haya lámina vacía ni encimada. Un visual o un hallazgo que el guion no tenía se anota en el outline; si cambia el argumento, vuelve al analista |
| 4 | `research-note` | Verificación: `audit --report`, `measure --shots`, evidencia frase por frase, tablas y rótulos incluidos (lo que no la tenga se elimina y se regenera), y pasada de texto. Con subagentes en el runtime, la evidencia la revisa uno que no redactó el deck |
| Gate 2 | — | El analista revisa las capturas **una por una** y el reporte: excepciones (`lit()`, hechos con verificación manual, frases eliminadas, cambios posteriores al gate 1, lo que quedó sin documento primario) y estado de P1 a P10. Aprueba la entrega o devuelve al paso 3. Entrada en thesis-journal |

El orden 1→gate→2 es doctrina: primero el argumento, después las cifras. Un
guion aprobado con un dato que no consigue fuente vuelve al analista; no se
rellena.

## Post-run

- En `research/notes/`, fechados: outline, hechos, deck HTML, reporte de
  verificación y carpeta de capturas; `build_deck.py` vigente (y la copia del
  anterior, si lo había). Nada previo se borra.
- `audit` en verde. `measure` en verde, o `[pendiente]` reportado como tal. El
  reporte de verificación lleva la salida de ambos (la de `measure` se agrega a
  mano: el tool no la escribe). P9 y P10 solo aplican con `brand/DESIGN.md`;
  sin él se reportan `[n/a]`.
- Dos entradas en `journal/thesis-journal.md` (gate 1 y gate 2).
- Check `[x]` sin resolver = **FALLA**: el deck no se entrega.
- Este comando no escribe en el modelo, el driver-map, el industry report ni el
  perfil.

## Reporte final

Una pantalla: láminas por acto, cifras auditadas contra el modelo, hechos
citados (cuántos con verificación manual), usos de `lit()` con su porqué, frases
eliminadas por falta de evidencia, estado de P1 a P10 y ruta del deck. Sin
recomendación de inversión.
