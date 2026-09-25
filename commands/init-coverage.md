# /init-coverage — iniciar cobertura de una emisora

Orquesta las etapas 0→6 del flujo: setup → ingesta ∥ industria → driver design →
modelo → populate. Comando = solo orquestación; toda la lógica vive en las skills.
Gates de usuario en cada transición + protocolo de debate
(`templates/debate-protocol.md`). Uso: `/init-coverage <TICKER> [folders con filings]`.

## Preflight (obligatorio — cualquier fallo DETIENE el comando; no corre parcial)

1. Ticker identificado y sin cobertura previa en el workspace (colisión ⇒ preguntar,
   nunca sobrescribir).
2. Folders del usuario existen y contienen filings legibles. Listar qué se encontró
   y qué falta (mínimo: último anual; ideal: + trimestrales del año). Transcripts
   de earnings calls: si están, se listan como encontrados y SE CONSIDERAN
   (statement-mapper los lee como fuente de guidance — no son opcionales una vez
   presentes). `brand/DESIGN.md` presente ⇒ xlsx-building lo usará para los
   colores de marca del modelo.
   **Filings faltantes + emisora SEC ⇒ ofrecer descarga** (gate — preguntar,
   nunca descargar solo). La propuesta default es la HISTORIA COMPLETA
   (todos los 10-K + 10-Q + 8-K disponibles; `sec_fetch` ya arranca en 1994 y
   `xbrl_fetch` trae todo) — perspectiva de largo plazo del plugin. JAMÁS
   proponer subconjuntos arbitrarios ("10 años", "solo el último año"): el
   gate del analista existe para recortar si ÉL quiere. ANTES de ofrecer, probar conectividad: en entornos
   con egress restringido (Claude Cowork, sandboxes) los tools reportan el
   bloqueo con opciones — la correcta suele ser permitir `www.sec.gov` y
   `data.sec.gov` en el allowlist del entorno; JAMÁS sustituir el filing
   íntegro por contenido procesado de un lector web (el pipeline cita por
   documento y página): `python tools/sec_fetch.py <TICKER> --dest
   <raiz>/<TICKER>/filings/sec --ua "<nombre correo>"` (o env `SEC_EDGAR_UA`;
   `--dry-run` para listar primero). El tool baja crudo por fecha de periodo +
   manifest CSV; coverage-folders renombra a la convención FY/#Q. Emisoras BMV:
   sin API pública — filings los trae el usuario.
3. Marco contable inferible del filing — si ambiguo, preguntar antes de crear perfil.
4. Herramienta determinista para xlsx disponible (openpyxl o equivalente) — sin
   ella no se construye modelo.
5. Usuario confirma alcance: periodos históricos, horizonte de forecast,
   **periodicidad del modelo** (anual / trimestral — pregunta OBLIGATORIA: si
   el paso 2 encontró trimestrales archivados, proponer `quarterly` como default
   y dejar decidir; JAMÁS asumir anual en silencio; `annual_plus_quarterly` está
   deprecado y se trata como `quarterly`) y métodos de valuación a estructurar
   (el perfil propondrá; aquí
   solo el marco general).
6. **Guard MNPI (Standard II(A)):** usuario confirma que TODO el contenido de los
   folders es información pública. Sin confirmación, no se procesa contenido.
7. `<raiz>/macro/macro-view.yaml`: existe y `updated_at` dentro de `staleness_warn_months` —
   stale o vacío AVISA (no bloquea): las etapas que lo consumen quedan marcadas.

## Pipeline (cada paso cierra con su gate + debate + entrada en thesis-journal)

| Paso | Skill | Gate del usuario |
|---|---|---|
| 1 | `coverage-folders` | — (entrevista de sesgo de entrada) |
| 2 | `framework-mapper` | Confirma marco y convenciones del perfil |
| 3 | `statement-mapper` | Aprueba mapeo y reclasificaciones |
| 4 | `industry-analysis` (∥ con 3) | Aprueba comp universe; FODA es suyo |
| 5 | `driver-inventory` · design | **Aprueba driver-map** — el debate central |
| 6 | `model-standards` | Checks verdes (integrity-checks.md completa) |
| 7 | `driver-inventory` · populate | Pone cada número; flags de líneas sin driver |

Los pasos 3 y 4 pueden correr en paralelo (la máquina captura mientras el analista
piensa industria); sus gates cierran por separado y AMBOS deben estar cerrados antes
del paso 5.

## Post-run (antes de reportar terminado)

- Balance check y tie-out de caja verdes; celda única de error check limpia.
- `Rev_Reconcile` dentro de umbral (o divergencia discutida y registrada).
- Artefactos completos: issuer-profile confirmado, industry-report, driver-map,
  modelo versionado, thesis-journal con las entradas de cada gate.
- Check rojo = reportar **FALLA** con detalle, nunca éxito parcial.

## Reporte final

Resumen de una pantalla: emisora, marco, drivers clave, estado de checks, precio
objetivo NO incluido (eso es de la nota, y es del analista), y pendientes abiertos.
