# research_analyst — guía para agentes (Cursor, Codex y cualquier agente sin soporte nativo de skills)

Este repo es un plugin de equity research. En Claude Code se instala como plugin
nativo (`.claude-plugin/plugin.json`). En cualquier otro agente, TÚ eres el runtime:
este archivo te enruta al contenido correcto. La fuente única de verdad son los
`skills/*/SKILL.md` — este archivo solo enruta, nunca duplica.

## Regla central (no negociable)

El modelo **redacta, estructura, verifica y señala — nunca calcula ni decide.**
Toda cifra sale de fórmula de Excel o código determinista. Todo supuesto lo pone el
analista. Toda cita normativa es exacta y verificada o se marca `[VERIFICAR]` —
jamás inventes una norma. Cuatro etiquetas siempre: `observado` (con cita) /
`guidance` (management) / `supuesto` (analista) / `output` (fórmula). Sin
recomendaciones de inversión.

## Enrutamiento: tarea → archivo a leer y seguir

| El usuario quiere… | Lee y sigue |
|---|---|
| Iniciar cobertura de una emisora ("inicia cobertura", "nueva emisora") | `commands/init-coverage.md` |
| Actualizar por trimestre nuevo / evento relevante | `commands/update-quarter.md` |
| Refrescar el análisis de industria ("la industria se movió", "actualiza el landscape con diff") | `commands/update-industry.md` |
| Actualizar el house view macro desde fuentes del analista ("actualiza el macro-view con mi research") | `commands/update-macro.md` |
| Auditar un modelo / correr checks | `commands/model-check.md` |
| Crear carpetas, archivar un filing | `skills/coverage-folders/SKILL.md` |
| Saber qué norma gobierna una línea, marco de la emisora, perfil contable | `skills/framework-mapper/SKILL.md` |
| Capturar cifras de un filing, mapear estados, snapshot de un comparable | `skills/statement-mapper/SKILL.md` |
| Analizar industria, competidores, FODA, comp universe | `skills/industry-analysis/SKILL.md` |
| Identificar drivers, revenue build-up, poblar forecast, calibración | `skills/driver-inventory/SKILL.md` |
| Construir el modelo xlsx, pestañas de valuación | `skills/model-standards/SKILL.md` |
| Escribir/formatear cualquier xlsx, audit de formato | `skills/xlsx-building/SKILL.md` |
| Clasificar materialidad de un hallazgo | `skills/impact-triage/SKILL.md` |

## Estilo de reporte (todos los runtimes)

Producto para analistas institucionales: los reportes al usuario son SOBRIOS.
Sin emojis ni iconos decorativos — estados como `[ok]` / `[x]` / `[aviso]` /
`[n/a]`, tablas y texto plano. Cifras con separador de miles y unidades en el
encabezado. El tono es el de una nota de research, no el de un chat: directo,
con fuentes, sin celebraciones. (El stdout de los tools ya es ASCII-only por
contrato; esta regla extiende lo mismo a la conversación.)

## Cómo ejecutar un comando

Los `commands/*.md` son orquestadores: preflight (si algo falla, DETENTE — no corras
parcial), secuencia de skills con gate de usuario entre pasos, y protocolo de debate
(`templates/debate-protocol.md`) al cierre de cada gate. Sigue el archivo del comando
paso a paso; cada paso te manda al SKILL.md correspondiente.

## Contratos (templates/)

`issuer-profile.yaml` (perfil por emisora) · `macro-view.yaml` (house view
compartido; se instancia en `<raiz>/macro/`, con sources/ e history/) · `comp-snapshot.yaml` (un comparable) ·
`driver-map` (formato en driver-inventory) · `model-spec.md` (estructura del xlsx) ·
`coverage-tree.md` (árbol y naming) · `debate-protocol.md` (formato del journal).
Cada artefacto tiene UNA skill dueña que lo escribe (está declarada en cada
SKILL.md); las demás solo leen.

## Restricciones de entorno

- Sin APIs de pago (v1); las keys gratuitas (FRED) las pone el usuario en su
  workspace, nunca en el repo ni en consola. Fuentes: filings locales del usuario,
  datos públicos, web search de tu plataforma si existe. Filings SEC faltantes:
  `python tools/sec_fetch.py <TICKER> --ua "<nombre correo>"` (EDGAR público,
  gratis; sin `--forms` detecta emisora doméstica — 10-K/10-Q/8-K — o
  extranjera — 20-F/40-F + 6-K; SIEMPRE con gate del usuario antes de descargar).
- xlsx solo vía `tools/xlsx_builder.py` (skill xlsx-building) — nunca openpyxl
  crudo para estructura/formato; audit de formato: `python tools/xlsx_builder.py audit <modelo>`;
  checks S + C10 + F en una corrida: `python tools/model_check.py <modelo>` (nunca
  reimplementes un check S a mano); con `--recalc` recalcula una copia con
  Excel o LibreOffice y evalúa además R1/R2 y los checks de fórmula viva
  (`python tools/recalc_validate.py <modelo>` hace solo ese paso). Sin
  recálculo los checks vivos quedan `[pendiente]`, nunca verdes.
- Historia financiera masiva de emisoras SEC: `python tools/xbrl_fetch.py <TICKER>`
  (companyfacts XBRL, US GAAP o IFRS en la moneda de reporte — décadas de
  anuales y trimestrales sin parsear HTML; reglas de vintage y de signo del CSV
  en statement-mapper).
- Nada se borra en la carpeta de cobertura; retención mínima 7 años.
- Documentación al usuario en español; nombres de archivos y skills en inglés.
