# /update-industry — refrescar el análisis de industria de una cobertura

Orquesta el refresh del industry-report cuando la industria se mueve (evento
sectorial, competidor nuevo, cambio regulatorio, o aviso de staleness D10).
Uso: `/update-industry <TICKER> [contexto del porqué]`.

## Preflight (cualquier fallo DETIENE el comando)

1. Árbol de cobertura existe y `research/industry/` tiene al menos un reporte
   previo — si no hay ninguno, esto es `/init-coverage` paso 4, no un update:
   abortar con esa instrucción.
2. `macro-view.yaml` dentro de staleness (aviso, no bloqueo — igual que D6).
3. Contratos sin fallas: `python tools/contract_check.py <raiz>/<TICKER>`.

## Pipeline

| Paso | Skill | Qué hace |
|---|---|---|
| 1 | `industry-analysis` | Refresh completo — lee filings, datos públicos Y `research/industry/sources/` (research del analista/terceros: si está, se considera; regla de citado en coverage-tree) → escribe `research/industry/industry-report_<YYYY-MM-DD>.md` NUEVO; el anterior queda intacto (vigente = fecha más reciente; nada se borra) |
| 2 | `industry-analysis` | **Diff contra la versión anterior**, sección por sección (§1–§7): tabla de cambios, cada uno con fuente |
| 3 | `impact-triage` | Clasifica cada delta (crítico/relevante/cosmético) — en especial los de §7 (supuestos de industria que alimentan el modelo) y §6 (cambios al comp universe) |
| 4 | `driver-inventory` | Para los deltas de §7 aprobados por el analista: contraste contra Assumptions vigentes; el analista decide qué supuesto cambia; registro en `journal/decisions.md` |
| 5 | — | Entrada de thesis-journal: ¿el movimiento de la industria confirma o erosiona la tesis? (debate per `templates/debate-protocol.md`) |

## Post-run

- Reporte nuevo fechado en `research/industry/`; el anterior intacto.
- Tabla del diff con clasificación de triage registrada.
- Si cambió el comp universe (§6): flag para regenerar snapshots de comps
  (statement-mapper) — no se regeneran solos.
- Si cambió `life_cycle_stage`: proponer actualización del issuer-profile
  (gate del usuario, framework-mapper la escribe).

## Reporte final

Una pantalla: qué cambió en la industria, clasificación de cada delta, qué
supuestos del modelo tocó el analista, y estado del thesis-journal.
