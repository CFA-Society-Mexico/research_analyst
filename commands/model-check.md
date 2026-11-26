# /model-check — auditoría de integridad del modelo

Corre SOLO los checks de integridad sobre un modelo, standalone. Equivale al modo
auditoría de `model-standards`. Sirve para modelos creados por el plugin Y para
modelos propios del analista que quiera auditar contra el estándar.
Uso: `/model-check [ruta del xlsx]` (default: el modelo vigente de la cobertura activa).

## Preflight

1. El xlsx existe y es legible con herramienta determinista (openpyxl o equivalente).
2. Si es modelo del plugin: driver-map e issuer-profile presentes (habilita los
   checks D1, D2, D7). Si es modelo externo: los checks D que dependen de contratos
   del plugin se reportan como `n/a`, no como falla.

## Ejecución

Corre la lista completa de `skills/model-standards/references/integrity-checks.md`
en orden S (estructurales) → C (contables) → D (contenido) → F (formato). Los S,
el C10 y los F corren en UNA corrida con `python tools/model_check.py
<modelo.xlsx> --driver-map <raiz>/<TICKER>/driver-map.md [--brand
brand/DESIGN.md]` — implementación única, cero interpretación del agente:
exit 1 = falla, exit 3 = pendiente de recálculo (se reporta `[pendiente]`,
nunca como ok). JAMÁS reimplementar un check S con un script propio de la
cobertura. Los C de fórmula viva se leen de la tab `Checks`; C9 y los D los
evalúa el agente con la regla escrita. Solo lectura: este comando NUNCA
modifica el modelo.
En modelo externo se agrega `--external`: F10 (sello del builder) se reporta
`[aviso]`, no falla. Sin la bandera, la falta de sello es falla.

## Recalculo (detección de ciclos — S10)

Si hay Excel disponible (COM en Windows), recalcular el libro
(`CalculateFullRebuild` + `Save`) y volver a correr `tools/model_check.py`: **celda de FÓRMULA cuyo valor cacheado
queda vacío tras el recálculo = FALLA S10** (referencia circular o cadena de
dependencias rota — el síntoma exacto del forecast sin calcular del smoke #5).
Sin Excel disponible: reportar S10 como `[pendiente: recalc manual]`, nunca
como ok.

## Reporte

```
MODEL CHECK — <archivo>  [fecha]
Estructurales : [ok] x N   [x] ...celda/hoja...   (tools/model_check.py)
Contables     : [ok] x N   [x] ...
Contenido     : [ok] x N   [aviso] ...   [n/a] ...
Formato (F)   : [ok] x N   [x] ...   (misma corrida de model_check.py)
Resumen: N ok · M fallas · K avisos
Veredicto: VERDE | FALLA (con siguiente acción sugerida por cada falla)
```

Regla: cualquier `[x]` ⇒ veredicto FALLA. Los avisos (staleness de comps o de
macro-view, monitor B-10) no tumban el veredicto pero se listan siempre.
