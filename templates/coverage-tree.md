# Coverage Tree — estructura por emisora y convención de nombres

## Workspace

Cuatro categorías: **inputs** (lo que entra) / **config** / **outputs** (lo que
el pipeline produce) / **registro** (memoria append-only). Archivos-contrato
únicos viven en la RAÍZ del ticker — un archivo no gana carpeta.

**La raíz es la carpeta que el usuario designó — su nombre es SUYO** (puede
ser `Valuation_Records`, `Valuation`, lo que sea; el plugin nunca la nombra ni
la renombra). El invariante es el CONTENIDO: la raíz contiene SOLO una carpeta
por emisora (`AAPL/`, `AMX/`, …) más `macro/` — nada más. JAMÁS crear una
subcarpeta intermedia (`workspace/`, `coberturas/`, etc.).

```
<carpeta-raiz>/                   # <- la del usuario, cualquier nombre
├── macro/                        # TODO lo macro — compartido; SOLO en la raiz
│   ├── macro-view.yaml           # VIGENTE, nombre fijo (lo consume código);
│   │                             #   dentro de un ticker es violación de contrato
│   ├── fred.key                  # API key de FRED (gratuita; la guarda el agente
│   │                             #   cuando el usuario la pega en el chat)
│   ├── fred-series.txt           # opcional: IDs de serie extra (uno por línea)
│   ├── series/                   # históricos macro observados (fred_fetch):
│   │   ├── DGS10.csv             #   date,value por serie + manifest.csv con
│   │   └── manifest.csv          #   título/unidades/última obs/fuente
│   ├── sources/                  # research macro del analista y de terceros
│   │                             #   (pdf/html/md — insumo de /update-macro)
│   └── history/
│       └── macro-view_YYYY-MM-DD.yaml   # snapshot ANTES de cada actualización
└── <TICKER>/
    ├── issuer-profile.yaml       # contrato raíz (framework-mapper)
    ├── driver-map.md             # contrato raíz (driver-inventory)
    ├── filings/                  # INPUTS
    │   ├── sec/
    │   │   ├── 10-K/             # emisora extranjera: 20-F/ (o 40-F/)
    │   │   ├── 10-Q/
    │   │   └── 8-K/              # emisora extranjera: 6-K/
    │   ├── bmv/
    │   │   ├── annual/
    │   │   ├── quarterly/
    │   │   └── eventos-relevantes/
    │   └── management/           # INPUTS: investor day, presentaciones y docs
    │                             #   de guidance del management (statement-mapper
    │                             #   los lee como fuente de guidance, igual que
    │                             #   transcripts/; jamás como observado)
    ├── transcripts/              # INPUTS: earnings calls — si están, SE CONSIDERAN
    │                             # (statement-mapper los lee como fuente de guidance)
    ├── comps/                    # INPUTS: snapshots fechados por comparable
    │   └── <COMP>_YYYY-MM-DD.yaml          # vigente = fecha más reciente; la
    │                             #   historia reconstruye el football field de
    │                             #   cualquier fecha (calibración)
    ├── brand/                    # CONFIG
    │   └── DESIGN.md             # opcional: colores de marca del analista (slots abajo)
    ├── model/                    # OUTPUTS
    │   ├── <ticker>-model_YYYY-MM-DD_v#.xlsx
    │   └── inputs/               # capa de captura (statement-mapper escribe,
    │       │                     #   model-standards consume; viaja con el modelo)
    │       ├── canonical_annual.csv      # CONTRATO: series anuales observadas
    │       ├── canonical_quarterly.csv   # CONTRATO: series trimestrales observadas
    │       ├── consolidated_long.csv     # CONTRATO: formato tidy largo (todas)
    │       └── extract_*.json            # papeles de trabajo por filing, con citas
    ├── research/                 # OUTPUTS: análisis (versionado por fecha)
    │   ├── industry/
    │   │   ├── sources/          # research de industria del analista y de
    │   │   │                     #   terceros (insumo de /update-industry)
    │   │   └── industry-report_YYYY-MM-DD.md   # generados; vigente = más reciente
    │   ├── notes/
    │   │   └── <ticker>-note_YYYY-MM-DD.md     # notas de inversión (etapa 7)
    │   └── sell-side/            # consenso y notas de terceros, FECHADOS
    │       └── <fuente>_YYYY-MM-DD.md   # GATE: se leen DESPUÉS de que el
    │                             #   analista pobló sus drivers, nunca antes
    │                             #   (sesgo de anclaje); el contraste consenso
    │                             #   vs propio se registra en journal/decisions.md
    └── journal/                  # REGISTRO append-only
        ├── thesis-journal.md     # entradas de cada gate/debate
        ├── decisions.md          # observado / guidance / supuesto / output, con fecha
        └── forecast-accuracy.md  # calibración por driver entre trimestres
```

### Convención de versionado por fecha

Artefactos de ESTADO se fechan: `<nombre>_YYYY-MM-DD.<ext>`, **vigente = fecha
más reciente**, nada se borra. Aplica a: `model/` (con `_v#` intra-día),
`research/industry/`, `research/notes/`, `comps/`. NO aplica a `journal/*`
(append-only: ya son historia) ni a los contratos de raíz (`driver-map.md`,
`issuer-profile.yaml`): son documentos vivos con gate cuya historia queda en
`journal/decisions.md` y thesis-journal — versionarlos duplicaría el journal.

Excepción para artefactos consumidos por CÓDIGO (`macro/macro-view.yaml`): el
vigente mantiene NOMBRE FIJO (los contratos y checks lo referencian por ruta) y
la historia vive en `history/` como snapshots fechados escritos ANTES de cada
actualización.

### Fuentes de terceros (`macro/sources/`, `research/industry/sources/`)

Research propio o de terceros en cualquier formato. **Si están, se consideran**
(como transcripts). Regla de citado: dato duro extraído ⇒ `observado` con
fuente = ese documento y página; opinión o estimación de tercero ⇒ se ATRIBUYE
("<fuente> estima X") y jamás se convierte en supuesto del analista sin su
decisión explícita en gate.

### model/inputs/ — la capa de captura

Los CSV canónicos son el CONTRATO máquina entre statement-mapper y
model-standards; los `extract_*.json` son papeles de trabajo (citas por cifra:
documento, página/nota, fecha). Columnas mínimas de los CSV: `line_item`,
`period`, `value`, `source_doc`, `source_ref`, `tag` (observado / derivado /
guidance). `period` es `3Q2016` en el trimestral y `FY2016` en el anual.
El histórico del xlsx se puebla DESDE los canónicos — check C9 verifica por
código que coincidan. Regenerables desde los JSON; los JSON, desde los
filings. Nada se teclea dos veces.

Reglas que C9 exige de esta capa:

- **`derivado`**: un subtotal o residual que el modelo necesita como cifra
  (el 4Q = FY − 3Q; "otros activos" = total − partidas) entra al CSV con
  `tag = derivado` y su fórmula en `source_ref`. No se calcula en el script
  de build: una cifra que no está en el CSV no tiene respaldo.
- **`guidance`** no respalda una cifra histórica.
- **Líneas reservadas**: `is_ns_total` (ventas) y `bs_ta` (activos totales).
  Son la referencia de materialidad de C9; sin ellas en el CSV y vinculadas
  en el modelo, toda cifra sin respaldo falla.
- **Unidades**: el CSV guarda la unidad de la fuente. El modelo declara la
  suya en cada vínculo (`link_capture(..., scale=1_000_000)`):
  `valor_csv = valor_xlsx × scale`. El histórico se escribe sin redondear;
  el formato de celda redondea la vista.

### Migración desde el árbol v1 (coberturas existentes)

Solo movimientos, nada se borra (Standard V(C)):
`profile/issuer-profile.yaml` → raíz · `assumptions/driver-map.md` → raíz ·
`notes/industry-report.md` → `research/` · `notes/thesis-journal.md` →
`journal/` · `log/*` → `journal/` · `earnings-transcripts/` → `transcripts/` ·
`macro-view.yaml` suelto (en la raíz o dentro de un ticker) →
`<raíz>/macro/macro-view.yaml` (si ya existe uno ahí, preguntar cuál manda —
nunca resolver en silencio). Carpetas viejas vacías se eliminan al final
(carpeta vacía no es dato).

## Convención de nombres de filings

Formato: `<TICKER>_<tipo>_<periodo>[_fecha].<ext>` — pensado para retrieval, no para humanos.

| Documento | Nombre |
|---|---|
| 10-K FY2025 | `AMX_10-K_FY2025.pdf` |
| 10-Q 2T2026 | `AMX_10-Q_2Q2026.pdf` |
| 8-K | `AMX_8-K_2026-08-15.pdf` |
| 8-K / 6-K bajado por `sec_fetch` | `TSM_6-K_2026-07-31_2026-08-15_000123.htm` — nombre crudo: periodo, fecha de presentación y `accn6` (últimos 6 dígitos de la accession); exhibits `…_000123_ex_<stem>.<ext>`. No se recorta: varios del mismo periodo quedan distintos |
| 20-F / 40-F FY2025 | `TSM_20-F_FY2025.htm` |
| Reporte anual BMV | `AMX_BMV-annual_FY2025.pdf` |
| Trimestral BMV | `AMX_BMV-quarterly_2Q2026.pdf` |
| Evento relevante | `AMX_evento-relevante_2026-08-15_<slug-corto>.pdf` |
| Transcript de earnings call | `AMX_transcript_2Q2026.pdf` |

Reglas:
- Ticker siempre en mayúsculas; periodos `FYyyyy` / `#Qyyyy`; fechas `YYYY-MM-DD`.
- Un archivo = un documento. Sin "final_v2_OK".
- **Nada se borra.** Filing reemplazado o corregido por la emisora: se archiva ambos, el
  nuevo con sufijo `_amended`. Retención mínima 7 años (Standard V(C)).
- Modelos: versionado `_YYYY-MM-DD_v#`; respaldo obligatorio antes de cada
  actualización trimestral (preflight de /update-quarter).

## brand/DESIGN.md — slots de marca (opcional)

El analista puede fijar los colores decorativos del xlsx. Formato determinista
(línea `slot: #RRGGBB`; lo parsea `tools/xlsx_builder.py::load_brand`):

```
brand_primary: #1F4E79    # barra de marca / cover (default: azul oscuro)
brand_section: #1F4E79    # bandas de sección, texto blanco (default: azul oscuro)
brand_accent:  #BDD7EE    # sub-secciones / acento (default: azul claro)
```

Solo esos 3 slots son rebrandeables. Los colores SEMÁNTICOS no se tocan jamás:
azul input `FF0000FF`, verde link `FF00CC00`, rojo warn, fill amarillo de input
— son el contrato de trazabilidad, no decoración. El audit F acepta la marca
pasándole el mismo DESIGN.md.
