---
name: impact-triage
description: Taxonomía de materialidad para hallazgos de equity research — clasifica cualquier hallazgo (diff de un filing nuevo, delta del diff de industria, evento relevante, noticia, cambio normativo, movimiento macro) como crítico, relevante o cosmético, anclado a su impacto en valuación y tesis vía las sensibilidades YA construidas del modelo; en hallazgos industriales recomienda /update-industry y en cambios macro corre triage TRANSVERSAL sobre todas las coberturas del workspace. Usa esta skill siempre que llegue un filing nuevo y haya que decidir qué importa, salga un evento relevante o noticia de la emisora, cambie una norma contable o un supuesto macro, se actualice el industry-report o el macro-view, o cuando el usuario pregunte "¿esto es material?", "¿qué cambia con este trimestre?", "¿me afecta este evento?", "triage de este 8-K".
---

# impact-triage

Dueña única de la clasificación crítico / relevante / cosmético. Consume outputs de
todas las demás; **jamás edita el modelo** — reporta, clasifica y recomienda dónde
mirar. La decisión de actuar es del analista.

## Cuándo corre

- `/update-quarter` paso 3: triage del diff ANTES de tocar el modelo.
- Standalone: evento relevante, noticia, 8-K, cambio normativo o macro suelto.

## La taxonomía (anclada a valuación, no a tamaño de titular)

Umbrales por emisora en `issuer-profile.yaml → materiality`:

| Nivel | Definición operativa | Acción que dispara |
|---|---|---|
| **Crítico** | Mueve el precio objetivo ≥ `critical_pct_target_price` O toca un pilar de la tesis (thesis-journal) O rompe un driver del driver-map | Actualización de modelo + debate obligatorio + probable nota. Si el hallazgo es de naturaleza INDUSTRIAL (competidor, regulación, participación, sustituto): recomendar `/update-industry` |
| **Relevante** | Entre `relevant_pct_target_price` y `critical_pct_target_price`, o cambia un supuesto secundario | Actualizar supuesto afectado; mención en próxima nota. Industrial ⇒ recomendar `/update-industry` (el analista decide) |
| **Cosmético** | Debajo de `relevant_pct_target_price` y sin contacto con tesis ni drivers | Registro en log; nada más |

El analista puede reclasificar cualquier hallazgo — la reclasificación se registra
con su porqué en `journal/decisions.md` (es señal de tesis, no corrección).

## Procedimiento — triage de diff trimestral

1. Toma el diff (lo nuevo capturado por statement-mapper vs el modelo vigente).
2. Por cada delta material: ¿qué driver lo explica? (driver-map) ¿toca un pilar de
   la tesis? (thesis-journal) ¿cuánto movería la valuación? — la estimación de
   impacto usa las sensibilidades YA construidas en el modelo (data tables), nunca
   cálculo propio de la skill.
3. Clasifica cada hallazgo con su racional en una tabla. Formato EXACTO (la
   columna "impacto estimado" siempre cita la sensibilidad del modelo o dice
   "no cuantificable aún" — nunca un número generado):

   | Hallazgo | Nivel | Por qué | Dónde pega | Impacto estimado |
   |---|---|---|---|---|
   | GM% Products −180bps vs estimado | CRÍTICO | rompe driver de margen del driver-map | Sch: COGS → IS → valuación | −7% precio objetivo (data table WACC×g, fila margen) |
   | Nuevo programa de recompras $110B | Relevante | cambia supuesto secundario de buyback | Sch: Buyback | pendiente de poblar |
   | Cambio de auditor | Cosmético | sin contacto con tesis ni drivers | — | — |
4. Entrega el triage ANTES de que nadie toque el modelo — el orden importa: primero
   saber qué importa, luego actualizar.

## Insumos de contexto (los usa, no los produce)

- **Posicionamiento** (industry-report §4): el mismo evento pega distinto al líder
  de costo que al premium — la materialidad depende de dónde compite.
- **Macro-view**: cambio macro ⇒ triage transversal a TODA la cobertura del
  workspace, no solo una emisora.
- **Monitor NIF B-10** (check D7, vía framework-mapper): emisora NIF con inflación
  trienal ≥ 22% acercándose al umbral de 26% ⇒ el escenario de reconexión
  retrospectiva se vuelve hallazgo RELEVANTE por defecto, CRÍTICO si ≥ 24%.
- **forecast-accuracy**: sorpresa recurrente en el mismo driver = hallazgo en sí
  mismo (el supuesto del analista está sesgado — citarlo).

## Reglas duras

- Sin acceso al modelo con sensibilidades construidas, el impacto en valuación se
  reporta como "no cuantificable aún" — nunca se estima por generación libre.
- Todo hallazgo clasificado queda en el log con fecha, aunque sea cosmético.

## Entrevista de cierre

Objetivo (adaptativa, per `templates/debate-protocol.md`): ¿el trimestre confirma o
erosiona la tesis? ¿Qué te hizo cambiar de opinión — y si nada, por qué nada?
Generar desde el triage concreto (ej. "clasificaste cosmético el cambio de working
capital, pero es el tercer trimestre consecutivo — ¿sigue siendo ruido?").
