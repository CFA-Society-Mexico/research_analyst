# Debate Protocol — entrevista y debate por gate

Protocolo único para el cierre de cada etapa del pipeline. Cada skill trae su
`## Entrevista de cierre` con el OBJETIVO del interrogatorio; este protocolo define
el formato y las reglas. Ningún gate se cierra sin su entrada en el journal.

## Los 4 pasos

1. **Resumen.** La skill presenta su output, cada elemento etiquetado
   (`observado` / `guidance` / `supuesto` / `output`).
2. **Entrevista adaptativa.** 2-4 preguntas generadas EN VIVO desde:
   - lo que la skill acaba de encontrar (el diff, el mapeo, los drivers),
   - el historial de `thesis-journal.md` (posturas previas, pendientes),
   - `journal/forecast-accuracy.md` si existe (sesgo histórico documentado del analista).
   Las preguntas NUNCA salen de un guion fijo. El SKILL.md fija el objetivo; el
   hallazgo dicta la pregunta.
3. **Debate corto.** El modelo toma la posición adversarial: steelman de la
   contra-tesis, argumentando SOLO con datos observados y normas verificadas.
   Máximo 2-3 rondas. El analista resuelve; el modelo jamás decide.
4. **Registro.** Entrada append-only en `journal/thesis-journal.md` (formato abajo).

Estilo en todo el protocolo: sobrio, sin emojis — estados `[ok]`/`[x]`/`[aviso]`,
tono de nota de research (ver AGENTS.md §Estilo de reporte).

## Reglas duras

- El modelo nunca inventa cifras ni citas para ganar el debate. Sin evidencia
  observada, la objeción se formula como pregunta, no como afirmación.
- El analista puede saltar el debate diciendo "sin debate" — se registra como
  `resolucion: saltado`. El gate cierra igual; la omisión queda documentada.
- El journal es append-only: sin ediciones retroactivas. Es la bitácora intelectual
  y el registro de diligencia (Standard V(A)/V(C)).
- La postura del analista es `supuesto`; la evidencia citada por el modelo es
  `observado`. Nunca se mezclan en una misma frase sin etiqueta.

## Formato de entrada en thesis-journal.md

```markdown
## [YYYY-MM-DD] <etapa> — <skill>

- **Contexto:** <qué output disparó la conversación, 1-2 líneas>
- **Postura del analista:** <supuesto> ...
- **Contra-argumento del modelo:** <observado: cita> ...
- **Resolución:** <mantiene | ajusta a ... | pendiente | saltado>
- **Impacto en tesis:** <1 línea — o "sin cambio">
```

## Objetivo por gate (referencia)

| Gate (tras…) | Objetivo del interrogatorio |
|---|---|
| `coverage-folders` | Sesgo de entrada: por qué esta emisora, qué espera encontrar |
| `framework-mapper` | Qué convención contable incomoda; dónde puede engañar la presentación |
| `statement-mapper` | Qué sorprendió de los históricos; calidad de earnings |
| `industry-analysis` | Dónde compite realmente y qué la protege — primera entrada sustantiva de tesis |
| `driver-inventory` (design) | **El debate central:** ¿cuáles son LOS drivers de esta emisora? |
| `model-standards` | ¿El modelo captura el negocio o solo la contabilidad? ¿Qué schedule falta? |
| `driver-inventory` (populate) | Cuál driver es el más frágil; dónde difiere del guidance y por qué |
| Valuación | Dónde discreparía el consenso; qué rompería la tesis |
| `impact-triage` | ¿El trimestre confirma o erosiona la tesis? ¿Qué cambió de opinión? |
| `research-note` (guion) | ¿El guion cuenta la tesis o recita el modelo? ¿Qué contra-argumento del journal no aparece? |
| `research-note` (entrega) | ¿Qué frase del deck no firmaría el analista? Excepciones y capturas, una por una |
