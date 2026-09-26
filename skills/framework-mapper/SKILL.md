---
name: framework-mapper
description: Regulación contable para el analista — deriva el marco de la emisora (IFRS, US GAAP, NIF o criterios CNBV), la periodicidad del modelo (anual / trimestral — siempre pregunta, jamás asume) y los métodos de valuación aplicables, y escribe issuer-profile.yaml; mapea qué norma exacta gobierna qué línea del modelo y qué diferencias entre marcos cambian un forecast. Usa esta skill siempre que haya que crear o confirmar el perfil de una emisora, citar una norma (ASC, NIC/NIIF, NIF), comparar tratamiento contable entre marcos, ajustar comparabilidad entre emisoras de marcos distintos, o cuando el usuario pregunte "¿qué norma gobierna X?", "¿cambia mi forecast por diferencia contable?", "¿bajo qué marco reporta esta emisora?" — es la ÚNICA skill autorizada a emitir citas normativas y a escribir el perfil.
---

# framework-mapper

Dueña única de `issuer-profile.yaml` y de toda cita normativa del plugin. Ninguna otra
skill cita normas. **Regla absoluta: norma exacta verificada o `[VERIFICAR]` — jamás
inventar una cita.** Una cita inventada es la falla más grave posible de esta skill.

## Cuándo corre

- `/init-coverage` paso 2: derivar y confirmar el perfil de emisora.
- Standalone: duda normativa puntual, ajuste de comparabilidad, revisión de perfil.

## Derivación del perfil (no preguntar a ciegas)

1. Lee el filing más reciente en `filings/` — la nota de políticas contables declara
   el marco.
2. Deriva por tipo de emisora (tabla en `references/ifrs-asc-nif-line-map.md`,
   sección "Mapa regulatorio"):
   - BMV no financiera → `ifrs` · banco/aseguradora → `cnbv_criteria` (**avisar:
     fuera de alcance v1**) · privada → `nif` · ADR/20-F → verificar declaración
     IFRS-IASB explícita y sin reservas.
3. Contrasta derivación vs lo declarado en el filing. Discrepancia = preguntar al
   usuario, no resolver por su cuenta.
4. Llena `templates/issuer-profile.yaml` completo (convenciones: moneda, unidades,
   cierre, segmentos, life_cycle_stage, métodos de valuación derivados,
   `model_periodicity` — propuesta: `quarterly` si coverage-folders archivó
   trimestrales, `annual` si solo hay anuales; PREGUNTA, no asumas;
   `annual_plus_quarterly` está deprecado).
5. Gate: el usuario confirma cada derivación (`framework_confirmed_by_user: true`
   y `periodicity_confirmed_by_user: true`).

Detección de financiera: `issuer_type: bank|insurer` ⇒ el perfil se crea, se marca
fuera de alcance, y el pipeline avisa que las citas del plugin NO aplican a criterios
CNBV/CNSF línea por línea (v2).

## Mapa norma→línea

Para cada consulta ("¿qué norma gobierna arrendamientos aquí?"):
1. Busca primero en `references/ifrs-asc-nif-line-map.md` — solo lo verificado ahí
   se cita directo.
2. Fuera de la tabla: cita con `[VERIFICAR: <norma>]` y dilo explícitamente, o
   verifica en fuente primaria (IFRS Foundation, FASB, CINIF, SEC, CNBV) antes
   de afirmar. Una guía de Big 4 sirve para ubicar el párrafo, no como fuente
   de la cita (CONTRIBUTING §4.7).
3. Siempre reporta las tres piezas: **norma exacta por marco → línea del modelo que
   toca → ajuste que exige al analista.** Ese es el formato; una norma sin línea ni
   ajuste es trivia, no análisis.

**Ejemplo del formato** (caso verificado en la reference):

> **Norma:** arrendamientos — IFRS 16 (todo en balance) vs ASC 842 (operativos
> con gasto lineal en resultado operativo).
> **Línea que toca:** EBITDA — bajo IFRS 16 el gasto de renta se vuelve
> depreciación + interés (EBITDA sube); bajo ASC 842 el gasto operativo se queda
> arriba (EBITDA no cambia).
> **Ajuste al analista:** en Val_Comps con universo mixto, fila de ajuste de
> arrendamientos antes de comparar EV/EBITDA — sin ella el múltiplo IFRS 16 se
> ve artificialmente barato.

Por qué el formato importa: el analista no necesita saber la norma — necesita
saber si su forecast o su comparación están rotos por ella.

## Diferencias que cambian un forecast (uso en pipeline)

- En design de drivers: avisar si una línea clave tiene tratamiento divergente entre
  el marco de la emisora y el de sus comps (ej. EBITDA IFRS 16 vs ASC 842).
- En Val_Comps: entregar la lista de ajustes de comparabilidad cuando el universo
  mezcla marcos.
- Monitor B-10: emisora NIF + inflación trienal (macro-view) acercándose a 26% ⇒
  exigir a impact-triage el escenario de reconexión (check D7).

## Trazabilidad

Todo lo que emite esta skill es `observado` (texto de norma o filing, con cita) o
está marcado `[VERIFICAR]`. La decisión de si una diferencia es material para la
tesis es del analista (impact-triage la clasifica; el analista resuelve).

## Entrevista de cierre

Objetivo (adaptativa, per `templates/debate-protocol.md`): ¿qué convención contable
de esta emisora te incomoda? ¿Dónde puede engañar la presentación? Generar las
preguntas desde lo que el perfil reveló (ej. "reporta NIF y la trienal va en 19% —
¿modelaste la reconexión?"). Registrar en thesis-journal.

## Referencias

- `references/ifrs-asc-nif-line-map.md` — LA tabla; leer completa antes de responder
  cualquier consulta normativa.
