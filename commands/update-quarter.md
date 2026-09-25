# /update-quarter — actualización trimestral de una cobertura

Orquesta la etapa 8 (mantenimiento): archivar → capturar → triage → re-checks →
actual-vs-driver + calibración. Uso: `/update-quarter <TICKER> [filing nuevo]`.

## Preflight (obligatorio — cualquier fallo DETIENE el comando)

1. Árbol de cobertura e `issuer-profile.yaml` existen y el perfil está confirmado.
2. **Baseline verde:** el modelo vigente pasa TODOS los checks ANTES de tocar nada.
   Roto ⇒ arreglar primero vía `/model-check`; nunca actualizar sobre un modelo roto.
3. Filing nuevo identificado, legible y no capturado previamente (anti-doble-captura).
   Sin filing entregado + emisora SEC ⇒ ofrecer bajarlo (gate — preguntar):
   `python tools/sec_fetch.py <TICKER> --since <inicio del trimestre> --dest
   <raiz>/<TICKER>/filings/sec --ua "<nombre correo>"`; coverage-folders
   renombra del manifest a la convención.
4. **Respaldo:** copia del modelo con versión nueva `_YYYY-MM-DD_v#` ANTES de editar.
5. Guard MNPI: el documento nuevo es público (un evento relevante filtrado antes de
   su publicación es exactamente el caso que este guard existe para atrapar).

## Pipeline

| Paso | Skill | Qué hace |
|---|---|---|
| 1 | `coverage-folders` | Archiva el filing con naming estándar |
| 2 | `statement-mapper` | Captura el trimestre, cada cifra con cita; si hay transcript del call en `transcripts/`, extrae su guidance (obligatorio); gate de mapeo |
| 3 | `impact-triage` | **Triage del diff ANTES de tocar el modelo**: crítico/relevante/cosmético con racional; gate del analista |
| 4 | `model-standards` | Integra lo aprobado; re-corre integrity-checks completa |
| 5 | `driver-inventory` | Actual vs driver por driver; append a forecast-accuracy.md; debate de calibración |

El orden 3→4 es doctrina: primero saber qué importa, luego actualizar. Un hallazgo
crítico puede cambiar QUÉ se actualiza y si amerita nota inmediata.

## Post-run

- Checks verdes en la versión nueva; versión anterior intacta como respaldo.
- Triage completo registrado (incluso cosméticos) en `journal/decisions.md`.
- forecast-accuracy.md con el trimestre appendeado.
- Entrada de thesis-journal del gate de impact-triage (¿confirma o erosiona?).
- Check rojo = **FALLA**, con la versión respaldada señalada como vigente.

## Reporte final

Tabla del triage (hallazgo → nivel → dónde pega), deltas de drivers (estimado vs
actual), estado de checks, y si el analista marcó la tesis como confirmada /
erosionada / pendiente.
