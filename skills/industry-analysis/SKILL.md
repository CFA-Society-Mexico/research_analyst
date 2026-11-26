---
name: industry-analysis
description: Análisis de industria y mercado estilo CFA para una cobertura de equity — landscape competitivo, five forces de Porter, ciclo de vida, FODA, posicionamiento y universo de comparables; produce reportes FECHADOS en research/industry/ (vigente = más reciente, historia completa) y en refresh genera el diff sección por sección contra la versión anterior; considera el research propio y de terceros que el analista deje en research/industry/sources/. Usa esta skill siempre que haya que analizar la industria de una emisora, mapear competidores, armar o refrescar un FODA, decidir el comp universe, evaluar posicionamiento o moats, comparar cómo cambió la industria desde el último reporte, o cuando el usuario diga "analiza la industria", "¿quiénes son los comparables?", "actualiza el landscape", "¿qué cambió en el sector?", "¿dónde compite esta emisora?".
---

# industry-analysis

Dueña única de `research/industry/` (reportes fechados: `industry-report_YYYY-MM-DD.md`,
vigente = fecha más reciente, nada se borra) y del universo de comparables. Es la etapa
top-down del proceso CFA: la tesis nace aquí, no en valuación. Los drivers de la
etapa siguiente se DERIVAN de la economía de industria que esta skill establece.

## Cuándo corre

- `/init-coverage` paso 4 (en paralelo con la captura de históricos) — primer reporte.
- `/update-industry`: refresh cuando la industria se mueve (evento sectorial,
  competidor nuevo, cambio regulatorio, aviso de staleness D10). El refresh
  escribe un reporte NUEVO fechado y produce el **diff sección por sección
  contra la versión anterior** — ese diff es el insumo del triage.
- Standalone: preguntas puntuales de industria sin refresh formal.

## Fuentes (sin APIs de pago — v1)

- Filings de la emisora Y de sus competidores (sección de negocio, MD&A, segmentos).
- Datos públicos (INEGI, reguladores sectoriales, cámaras) y web search de la
  plataforma si existe.
- `research/industry/sources/` — research del analista o de terceros (pdf/html/
  md): **si está, se considera**. Dato duro ⇒ `observado` con documento y
  página; opinión de tercero ⇒ atribuida ("<fuente> estima X"), jamás supuesto
  del analista sin su gate.
- `macro/macro-view.yaml` para el contraste industria vs ciclo macro (el
  bloque del país donde compite la emisora: primer nivel o
  `countries.<pais>`; si la industria es multi-país, se dice cuál se usó).
- **Etiquetado estricto:** hecho con fuente = `observado`; interpretación del
  analista (FODA, posicionamiento) = `supuesto`; lo que diga el management de
  cualquier emisora = `guidance`.

## Estructura de industry-report.md (marco CFA)

```markdown
# Industry Report — <industria> / <TICKER>   [fecha]

## 1. Estructura de la industria (five forces)
   Entrada · proveedores · compradores · sustitutos · rivalidad — cada fuerza con
   evidencia observada, no adjetivos.
## 2. Ciclo de vida y drivers de la industria
   embryonic|growth|mature|decline (alimenta life_cycle_stage del perfil) ·
   cíclica vs defensiva · sensibilidad a tasas/commodities/regulación/tecnología ·
   crecimiento del mercado vs PIB (insumo de la ruta top-down de revenue)
## 3. Landscape competitivo
   Tabla de competidores: participación (con fuente), marco contable, métricas
   comparables clave. De aquí sale el comp universe.
## 4. Posicionamiento de la emisora
   Moat (costo|marca|red|switching|IP) con evidencia · posición débil/media/fuerte
## 5. FODA
   Del ANALISTA (supuesto) — la skill estructura y desafía, no lo escribe por él
## 6. Universo de comparables
   Tickers seleccionados + selected_because por cada uno + quién quedó fuera y por qué
## 7. Supuestos de industria para el modelo
   Bloque que va a assumptions: crecimiento de mercado, participación, precios
   sectoriales — cada uno con fuente o marcado supuesto
```

## Formato del diff (refresh vía /update-industry)

Tabla sección por sección; solo filas con cambio real. La columna "¿toca el
modelo?" es la que consume impact-triage — sin ella el diff es lectura, no
pipeline:

| § | Antes | Ahora | Fuente del cambio | ¿Toca el modelo? |
|---|---|---|---|---|
| 2. Ciclo | crecimiento mercado 8% | 5% (desaceleración e-commerce) | reporte X p.4 (`sources/`) | SÍ — ruta top-down de revenue |
| 6. Comps | 5 tickers | +SHOP, −SSNLF | filing SHOP 10-K | SÍ — regenerar snapshot |
| 4. Posicionamiento | sin cambio | — | — | no |

## Contratos de salida

| Output | Consumidor |
|---|---|
| `industry-report.md` completo | analista; nota (etapa 7) |
| Universo de comparables (sección 6) | statement-mapper (snapshots) y Val_Comps |
| Bloque de supuestos de industria (sección 7) | driver-inventory → Assumptions |
| `life_cycle_stage` | issuer-profile (horizonte de forecast, peso de métodos) |
| Posicionamiento | impact-triage (materialidad depende de dónde compite) |

## Reglas duras

- Participaciones de mercado y tamaños SIEMPRE con fuente y fecha, o marcados
  `supuesto`. Un landscape sin fuentes es opinión con formato.
- El FODA es del analista: la skill propone estructura y evidencia, pregunta,
  desafía — el contenido final lo decide él.
- El comp universe es juicio del analista; la skill documenta el racional de cada
  inclusión Y de cada exclusión notable.

## Entrevista de cierre

Objetivo (adaptativa, per `templates/debate-protocol.md`): **¿dónde compite realmente
esta emisora y qué la protege?** El modelo desafía el FODA con lo observado en
filings de competidores (ej. "dices ventaja de costo; el competidor X reporta margen
bruto 400bps mayor — ¿cómo se sostiene tu F?"). Esta es la primera entrada
SUSTANTIVA del thesis-journal: la semilla de la tesis, con sus contra-argumentos.
