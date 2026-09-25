# research_analyst

Autor: [Alan Vazquez, CFA](https://www.linkedin.com/in/alan-vazquez-cfa-38515414a/)

> Copiloto de análisis bursátil: acompaña a un analista desde "me asignaron
> cubrir esta empresa" hasta "publiqué mi reporte".
> Open source · **CFA Society México · AI for Finance** · v0.6.0 (en dogfood).

Un analista de acciones pasa semanas leyendo reportes, armando un Excel y
escribiendo su nota. Este plugin **organiza, redacta, verifica y cuestiona** ese
trabajo. La regla que no se rompe:

> **El asistente nunca calcula ni decide.** Toda cifra sale de una fórmula de
> Excel o de código verificable. Todo supuesto lo pone el analista. Toda cita
> normativa es exacta o se marca `[VERIFICAR]`. No emite recomendaciones de
> inversión — esas llevan la firma del analista.

Cada dato queda etiquetado para poder rastrearlo siempre: `observado` (del
reporte oficial, con página) · `guidance` (lo dijo la dirección) · `supuesto`
(lo puso el analista) · `output` (resultado de fórmula).

## Instalación

**Claude Code / Claude Cowork:**

```
/plugin marketplace add CFA-Society-Mexico/research_analyst
/plugin install research-analyst@cfamexico
```

(En Cowork: menú **Customize → Add marketplace** → `CFA-Society-Mexico/research_analyst`.)

**Cursor / Codex:** clona el repo en tu workspace; [`AGENTS.md`](AGENTS.md)
enruta cada tarea a la skill correcta. Codex también lee el marketplace:
`codex plugin marketplace add CFA-Society-Mexico/research_analyst`.

> **Entornos con red restringida (Claude Cowork, sandboxes):** para que la
> descarga de filings, de historia XBRL y de series macro funcione, permite los
> dominios `www.sec.gov`, `data.sec.gov` y `api.stlouisfed.org` en el allowlist
> del entorno ANTES de abrir la cobertura. Si no se puede, corre
> `tools/sec_fetch.py`, `tools/xbrl_fetch.py` y `tools/fred_fetch.py` en tu
> máquina local y copia los archivos — los tools te lo indican solos al detectar
> el bloqueo.

## El flujo completo

Cinco comandos orquestan todo. Antes de correr verifican que nada falte
(incluida la confirmación de que no hay información privilegiada); si algo
falla, se detienen — nunca dejan trabajo a medias.

```
/init-coverage AAPL ./mis-filings/     # abrir una cobertura nueva
/update-quarter AAPL ./nuevo-10q.pdf   # actualizarla cada trimestre
/update-industry AAPL                  # refrescar la vista de industria, con diff
/update-macro                          # actualizar el view macro desde tus fuentes
/model-check                           # auditar el Excel en cualquier momento
```

`/init-coverage` recorre estas etapas, con una pausa (*gate*) entre cada una
donde el asistente debate contigo — con datos observados, nunca inventados — y
todo queda escrito en un diario de tesis:

1. **Carpetas y archivado** (`coverage-folders`) — instancia la estructura
   estándar y archiva tus documentos. Si faltan filings de una empresa de la
   SEC, ofrece descargarlos (`tools/sec_fetch.py`, EDGAR gratuito).
2. **Perfil de la empresa** (`framework-mapper`) — deduce el marco contable
   (IFRS / US GAAP / NIF), la periodicidad del modelo (anual o trimestral —
   siempre te pregunta) y los métodos de valuación
   que aplican. Tú confirmas cada derivación.
3. **Captura de históricos** (`statement-mapper`) — para emisoras SEC, décadas
   de anuales y trimestrales en un comando (`tools/xbrl_fetch.py`, datos XBRL
   oficiales); cada cifra queda con su cita en archivos canónicos, y tú
   apruebas el mapeo. Si hay transcripts de earnings calls, extrae el
   guidance — obligatorio.
4. **Industria y mercado** (`industry-analysis`, en paralelo con 3) — Porter,
   FODA, ciclo de vida y el universo de comparables.
5. **Diseño de drivers** (`driver-inventory`) — define QUÉ mueve los ingresos
   (precio × volumen, tiendas × venta por tienda…) ANTES de abrir Excel. Es el
   debate central del proceso.
6. **Construcción del modelo** (`model-standards` + `xlsx-building`) — genera
   el Excel con formato de código, no de criterio. En modo trimestral: hoja
   **Operating** (el modelo se construye sobre trimestres — supuestos,
   estados, razones y schedules en secciones colapsables, navegación estilo
   CFI) + hoja **Annual** (los años como agregado calculado de sus trimestres
   + el DCF desglosado línea por línea). Todo vigilado por ~50 checks.
7. **Poblar el forecast** (`driver-inventory`, segunda pasada) — el analista
   pone cada número viendo su serie histórica al lado; el asistente contrasta
   contra el guidance y registra las diferencias.

`/update-quarter` repite el ciclo corto cada trimestre: archiva el filing nuevo,
captura, **clasifica el impacto antes de tocar el modelo** (`impact-triage`:
crítico / relevante / cosmético), re-corre los checks y mide qué tan bien
estimaste cada driver (calibración).

`/update-industry` refresca la vista de industria cuando el sector se mueve (o
cuando el reporte lleva más de 6 meses sin tocarse — el plugin te avisa): escribe
un reporte NUEVO fechado, muestra el **diff contra la versión anterior**, y el
triage clasifica qué deltas tocan tus supuestos.

## Cómo se ve una cobertura en disco

Cada empresa vive en su carpeta, organizada en cuatro categorías. Regla general:
lo que cambia con el tiempo se guarda **fechado** — el vigente es el más
reciente y nada se borra jamás (retención de 7 años).

```
<tu-carpeta-raíz>/             # cualquier nombre; adentro SOLO emisoras + macro/
├── macro/                     # TODO lo macro, compartido por todas las coberturas:
│   ├── macro-view.yaml        #   los supuestos vigentes (tasa, FX, PIB, inflación)
│   ├── series/                #   históricos observados que baja `tools/fred_fetch.py`
│   │                          #     (un CSV por serie + manifest con fuente y fecha)
│   ├── fred.key               #   tu API key de FRED (gratuita); nunca sale del disco
│   ├── fred-series.txt        #   opcional: IDs de series extra que quieras bajar
│   ├── sources/               #   tu research macro y el de terceros (pdf, html, md)
│   └── history/               #   cada versión anterior del view, fechada
└── AAPL/
    ├── issuer-profile.yaml    # ficha de la empresa: marco contable, periodicidad,
    │                          #   métodos de valuación — tú confirmas cada campo
    ├── driver-map.md          # QUÉ mueve los ingresos y costos — el contrato del modelo
    ├── filings/               # ENTRADAS · reportes oficiales (SEC y BMV), organizados
    ├── transcripts/           # ENTRADAS · earnings calls — si están, se usan como guidance
    ├── comps/                 # ENTRADAS · fotos fechadas de cada comparable
    ├── brand/                 # CONFIG · tus colores de marca para el Excel (opcional)
    ├── model/                 # SALIDAS · el Excel versionado por fecha
    │   └── inputs/            #   cifras capturadas de los filings, con su cita
    ├── research/              # SALIDAS · análisis fechados
    │   ├── industry/          #   reportes de industria (historia completa, con diffs)
    │   │   └── sources/       #     tu research del sector y el de terceros
    │   └── notes/             #   notas de inversión publicadas
    └── journal/               # REGISTRO · solo se agrega, nunca se borra:
                               #   diario de tesis · decisiones · aciertos de forecast
```

## Las 8 skills — y cuándo usarlas sueltas

No hace falta correr el flujo completo ni memorizar nombres: cada skill también
responde a lenguaje natural, sobre coberturas del plugin o trabajo tuyo previo.

| Skill | Es dueña de | Úsala sola cuando digas… |
|---|---|---|
| `coverage-folders` | Carpetas, nombres, archivado (nada se borra; retención 7 años) | "archiva este 10-Q", "¿qué documentos me faltan?" |
| `framework-mapper` | Perfil de la emisora y TODAS las citas normativas | "¿qué norma gobierna arrendamientos?", "¿cambia mi proyección por diferencia contable?" |
| `statement-mapper` | Cifras capturadas con cita; snapshots de comparables | "mapea este PDF al modelo", "¿de dónde salió esta cifra?" |
| `industry-analysis` | Reporte de industria y universo de comparables | "actualiza el panorama de la industria", "¿quiénes son los comparables?" |
| `driver-inventory` | Mapa de drivers, supuestos, historial de aciertos | "¿qué líneas no tienen driver?", "¿qué tan bien estimé el trimestre?" |
| `model-standards` | El Excel: estructura, fórmulas, checks, valuación | "revisa mi modelo", "corre los checks" — funciona también sobre modelos que no creó el plugin |
| `xlsx-building` | Formato del Excel por código (colores, fuentes, checks F) | "el modelo salió feo", "corre el audit de formato" |
| `impact-triage` | Clasificar hallazgos: crítico / relevante / cosmético | "salió este evento, ¿es material?" |

## Qué empresas cubre

| Tipo | ¿Hoy? |
|---|---|
| SEC (10-K/10-Q, US GAAP) · BMV no financiera (IFRS) · privada (NIF) · FIBRA/REIT | ✅ |
| ADR (20-F, IFRS) | ⚠️ Perfil, modelo y valuación sí. `sec_fetch` baja 20-F y 6-K con `--forms 20-F,6-K`; `xbrl_fetch` solo lee US GAAP, así que la historia se captura desde el filing |
| Comparar empresas entre marcos contables | ⚠️ Las 7 diferencias más comunes verificadas; el resto se marca `[VERIFICAR]` |
| Bancos y aseguradoras | ❌ v2 — falta el mapeo CNBV/CNSF; el plugin lo detecta y te avisa |

## Extras que el analista controla

- **`brand/DESIGN.md`** — pon tus colores de marca y el Excel se genera con
  ellos (los colores de trazabilidad no se tocan).
- **`macro/`** — los supuestos macro de la casa (tasa, FX, PIB) viven una sola
  vez a nivel workspace y alimentan todas las coberturas. Deja tu research en
  `macro/sources/` y `/update-macro` te propone la actualización campo por
  campo, con cita — tú confirmas cada valor. Los históricos observados
  (UST 10Y, Fed Funds, CPI, PIB real, USDMXN) los baja `tools/fred_fetch.py` a
  `macro/series/` con API key gratuita de FRED, para que la propuesta salga de
  series con fuente y no de memoria.
- **`sources/`** (macro y por industria) — deja ahí research tuyo o de terceros
  en cualquier formato; el pipeline lo considera y lo cita. ¿Quieres leer un
  `.yaml` sin pelearte con el formato? Pídele al asistente: "explícame mi
  macro-view".
- **`transcripts/`** — sube los transcripts de earnings calls y el pipeline los
  considera automáticamente.

## Contribuir

El protocolo completo — invariantes que ningún PR puede romper, cómo agregar un
check, convenciones de commits, checklist de verificación — está en
[`CONTRIBUTING.md`](CONTRIBUTING.md). Léelo antes de abrir un PR.

Fronteras conocidas, cada una un PR bienvenido: bancos (Anexo 33 vs IFRS 9),
aseguradoras (CNSF vs IFRS 17), citas NIF pendientes de fuente primaria,
convención AFFO de FIBRAs, suite de evals, y el **conector de Banxico SIE**
(series MX, token gratuito) bajando a `macro/series/` como CSVs con manifest —
mismo patrón que `tools/fred_fetch.py`, que ya cubre FRED. Regla de la casa:
toda cita normativa con fuente primaria o `[VERIFICAR]`.

El diseño completo y el porqué de cada decisión: [`docs/architecture.md`](docs/architecture.md).

## Licencia

MIT. Proyecto educativo y de productividad. **No constituye asesoría de
inversión** — el analista es responsable de todo supuesto, cifra y
recomendación que publique.
