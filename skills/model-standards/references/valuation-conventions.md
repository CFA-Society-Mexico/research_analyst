# Valuation Conventions — menú de métodos y pestañas deterministas

Doctrina: **el modelo arma la estructura y las fórmulas; el analista pone cada
supuesto; nada se calcula por generación libre.** Usuario en el loop en cada supuesto.
El football field agrega SOLO métodos activos en `issuer-profile.yaml`.

## Menú y activación

| Activación | Método | Supuestos del analista |
|---|---|---|
| **Core — siempre** | DCF FCFF multi-stage, terminal DUAL | WACC (componentes), g terminal, RONIC terminal, múltiplo de salida, horizonte |
| **Core — siempre** | Comps (EV/EBITDA, P/E, P/B, EV/Sales) | universo (viene de industry-analysis), múltiplo aplicable |
| **Core — siempre** | Sensibilidades (cuadrícula de 2 variables, fórmula cerrada por celda) + football field | pasos de sensibilidad |
| Dividendera estable | DDM sobre los dividendos del modelo + terminal de equity normalizado | payout por periodo (Assumptions), ROE terminal, ke |
| Apalancamiento estable | FCFE sobre CFO − capex + endeudamiento neto del modelo + el mismo terminal | trayectoria de deuda (Schedules), ROE terminal, ke |
| FIBRA / REIT | NAV + FFO/AFFO | cap rate, ajustes AFFO — convención local FIBRA [VERIFICAR] |
| Conglomerado | SOTP | método por segmento, descuento de holding |
| Banco (v2) | Residual income + justified P/B | r, ω persistencia — **especificado, INACTIVO v1** |

## Convenciones por pestaña

### Sección DCF de `Annual` (FCFF multi-stage) — antes tab Val_DCF

Vive como sección de la hoja `Annual` (model-spec §Annual), LÍNEA POR LÍNEA
según el bloque del spec: EBIT → impuestos → NOPAT → +D&A → −Capex → −ΔWC →
FCFF (histórico Y forecast — el FCFF realizado es sanity visible) → factor de
descuento → PV → terminales → EV → valor/acción → cruces → reverse DCF →
Hamada. Cero fórmulas comprimidas. Convenciones de cada pieza:
- FCFF derivado por fórmula desde el modelo (EBIT×(1−t) + D&A − CapEx − ΔWC), nunca
  re-tecleado.
- WACC: componentes visibles — rf de `Macro` (macro-view: primer nivel o
  `countries.<pais>.rates.risk_free_10y`), ERP de `Macro` (mismo bloque), beta
  (supuesto con fuente), costo de deuda, pesos a mercado. Cada uno etiquetado.
- **Diagnóstico del horizonte** (lectura del check D4): si NINGUNA g
  económicamente sostenible — por debajo del crecimiento nominal de largo
  plazo de la economía, ~4% — reconcilia el terminal Gordon con el múltiplo de
  salida anclado en comps, el problema NO son los supuestos: **el horizonte
  explícito es demasiado corto**. El mercado descuenta un periodo largo de
  crecimiento superior, no una perpetuidad a esa tasa; con el horizonte
  correcto ese crecimiento se agota DENTRO del tramo explícito y la g terminal
  vuelve a ser baja y defendible. Antes de forzar la g o el múltiplo para que
  D4 pase, evaluar si extender el horizonte es la respuesta honesta. Señal de
  alarma: el valor terminal domina el valor total (regla práctica: >75%).
- **Terminal Gordon con flujo NORMALIZADO** (`ModelStyler.gordon_terminal_block`):
  `flujo terminal = NOPAT_n × (1+g) × (1 − g/RONIC)` y
  `TV = flujo terminal / (WACC − g)`. `g/RONIC` es la tasa de reinversión que
  esa g exige en estado estable. PROHIBIDO perpetuar `FCFF_n × (1+g)`: el FCFF
  del último año explícito trae la reinversión de una empresa que todavía
  crece, y perpetuarlo reinvierte como si se creciera a esa tasa sin recibir el
  crecimiento (cobertura TSMC: capex de 45% de ventas en el último año
  explícito ⇒ 1,025 por acción con `FCFF_n` contra 2,065 con el flujo
  normalizado, y un múltiplo implícito de 2.7x EV/EBITDA contra 20.6x de
  mercado). **RONIC** es input etiquetado de Assumptions (`supuesto`); default
  = link a WACC: el crecimiento terminal no crea valor, el supuesto defendible
  sin evidencia de un moat perpetuo. Con `RONIC > WACC` el crecimiento sí crea
  valor y el analista lo justifica en el journal.
- **Terminal dual obligatorio:** Gordon (g) Y exit multiple, lado a lado.
  Cruce: Gordon ⇒ múltiplo implícito; exit multiple ⇒ g implícita. Ambos visibles.
  Divergencia grande = check D4: revisar supuesto con el analista, no promediar.
  Además, un múltiplo implícito del Gordon por debajo de `TV_MULT_FLOOR` × la
  mediana de comps (constante con nombre, default 0.5) es **alarma, no dato
  informativo** — check D11 (`terminal_multiple_alarm_formula`): casi siempre
  delata un terminal que perpetúa capex de crecimiento.
- **g implícita de CUALQUIER terminal** (exit multiple o reverse DCF) —
  `ModelStyler.implied_growth_block`, forma cerrada, sin Goal Seek. Con el
  terminal normalizado el despeje deja de ser lineal: es la **raíz menor** de
  `A·g² − B·g + C = 0`, con `A = N/R`, `B = TV + N·(1 − 1/R)` y
  `C = TV·W − N` (N = NOPAT_n, R = RONIC, W = WACC). El bloque lo escribe línea
  por línea y sin literales (S4): `m = B/(A+A)` y `g = m − √(m·m − C/A)`.
  - **Guarda obligatoria:** con `R = W` una raíz es exactamente `g = WACC`,
    solución espuria (el TV explota). Una raíz ≥ WACC se reporta como texto
    — `ninguna g < WACC reproduce el TV` —, jamás como número; discriminante
    negativo ⇒ `sin solucion real`.
  - Con `R = W` el bloque NO usa la cuadrática sino el despeje lineal exacto
    `g = TV·W/N − 1`: en punto flotante la raíz espuria sale `0.0899999…` y
    pasaría la guarda como si fuera válida (verificado con Excel).
  - Nunca se despeja con un flujo proyectado con la g del analista: N es el
    NOPAT del último año explícito.
- **Año en curso (stub).** En modo `quarterly`, si el último trimestre
  reportado no es 4Q, el primer año del DCF ya tiene trimestres observados
  cuya caja está dentro de la deuda neta del último balance. Descontar ese FY
  completo cuenta dos veces ese flujo. Regla: la fecha de valuación es el
  cierre del último trimestre reportado; el primer periodo del DCF es el
  **stub** = FCFF solo de los trimestres estimados que faltan del año (suma
  de sus columnas E en `Operating`), con fracción de año `f = trimestres
  restantes / 4`. Exponentes de descuento: el stub, `f` (o `f/2` con
  mid-year); el año completo k después del stub (k = 1…n), `f + k` (o
  `f + k − 0.5` con mid-year); el TV, al cierre del último año explícito,
  `f + n`. La deuda neta es la del mismo último trimestre reportado.
- Etapas: single-stage es el caso degenerado; default 2-3 etapas según
  `life_cycle_stage` (growth ⇒ horizonte largo).
- **Bloque beta pure-play (Hamada)** — mecánica visible, no "beta con fuente" a secas:
  - β_u por comp = β_l / (1 + (1−t)·D/E) — una fila por comp; β_l input `observado`
    con fuente en comentario; D y E del `comps/*.yaml`.
  - β_u grupo = mediana; β_relevered = β_u grupo × (1 + (1−t)·D/E objetivo).
  - D/E objetivo: `supuesto` del analista en Assumptions. El analista puede
    sobreescribir el beta final (documentado en journal/decisions.md); la mecánica
    queda visible como referencia.
- **Bloque reverse DCF (expectativas implícitas)** — forma cerrada, sin Goal Seek:
  - EV de mercado = mkt cap actual + deuda neta + minoritarios + preferentes.
  - TV implícita = (EV − PV de FCFF explícitos) ÷ factor de descuento que el
    DCF aplica al TV (`(1+WACC)^n` si el TV se descuenta a fin del año n; con
    stub, el mismo exponente `f + n`).
  - g implícita = `implied_growth_block` sobre la TV implícita (raíz menor de
    la cuadrática del terminal normalizado, con su guarda). Las formas
    anteriores quedan derogadas: `WACC − FCFF_{n+1} / TV` contaminaba la g "del
    mercado" con la g del analista, y `(TV × WACC − FCFF_n) / (TV + FCFF_n)`
    despejaba un terminal que perpetuaba `FCFF_n`.
  - Tercera columna junto al cruce de terminales: "el mercado descuenta g = X; tú
    supones g = Y" (check D4b — debate, nunca bloqueo).

### Val_Comps
- Múltiplos calculados POR FÓRMULA desde `comps/*.yaml`: EV = mkt cap + deuda − caja
  + minoritarios + preferentes; cada componente del snapshot, con fuente.
- **Market cap en la moneda de reporte** (`comps_market_cap_formula`):
  `precio ÷ acciones por ADS × acciones ordinarias × FX`, con
  `market_data.price_currency`, `shares_per_ads` y `fx_to_reporting` del
  snapshot. El precio de un ADR viene por ADS y en su moneda de cotización;
  multiplicarlo directo por las acciones ordinarias y sumarlo a deuda y EBITDA
  en otra moneda dio EV/EBITDA 1.0x en vez de 10.1x para un comparable (UMC en
  la cobertura de TSMC) y movió la mediana. Snapshot sin esos tres campos = no
  entra a Val_Comps hasta completarlo (statement-mapper).
- **Rango de cordura** (check D12, `multiple_range_flag_formula`): un EV/EBITDA
  fuera de `COMP_MULT_MIN`–`COMP_MULT_MAX` (constantes con nombre, default
  2x–60x) se marca `REVISAR` en su fila. Es aviso: el múltiplo no se excluye
  solo; el analista confirma el dato o corrige el snapshot.
- Promedio del grupo: **media armónica** (doctrina CFA para múltiplos), mediana como
  referencia; nunca media aritmética sola.
- Múltiplos no significativos (denominador ≤ 0: utilidad, EBITDA o valor en
  libros negativos o cero) se EXCLUYEN de la media armónica y de la mediana y
  se muestran como `NM` con su motivo, por fórmula (`IF` sobre el
  denominador). Con un negativo dentro, la media armónica cambia de signo o se
  dispara. Excluir el múltiplo no borra al comparable: su fila sigue visible.
- Staleness: `as_of` de cada snapshot flaggeado si viejo (check D3).
- Comparabilidad entre marcos: EBITDA IFRS 16 vs ASC 842 NO comparable directo —
  fila de ajuste de arrendamientos cuando el universo mezcla marcos (ver
  framework-mapper reference).
- **Bloque de múltiplos justificados + PVGO** (debajo del grupo de comps) — cero
  inputs nuevos, todo por referencia a b, g, r, ROE ya existentes en
  Assumptions/Val_DCF/Ratios:
  - P/E justificado leading = (1−b)/(r−g); trailing = (1−b)(1+g)/(r−g).
  - P/B justificado = (ROE−g)/(r−g); P/S justificado = (E₀/S₀)(1−b)(1+g)/(r−g).
  - PVGO = P₀ − E₁/r, y PVGO/P₀ (% del precio que es expectativa de crecimiento).
  - Lectura por fila: múltiplo de mercado vs justificado — de descriptivo ("a
    cuánto cotizan") a normativo ("a cuánto deberían").

### Secciones DDM y FCFE de `Annual` (condicionales)

Viven en `Annual` debajo del DCF (no en tabs propias): los flujos y la fila
de exponentes ya están ahí. Los exponentes (mid-year incluido) salen de la
fila del DCF; el año en curso con trimestres observados entra como **stub**
por `stub_flow_ref` / `stub_exponent_ref` de `equity_value_block`: el flujo
del stub es la suma de las columnas de trimestres E del año en `Operating`
(dividendos, o los componentes del FCFE) y su exponente, el del stub del DCF
(`f` o `f/2`). Con stub, el rango de flujos cubre solo los años completos
posteriores: descontar el FY completo contaría dividendos o FCFE ya pagados.
- **Flujos desde el modelo, nunca re-tecleados** (`ModelStyler.equity_flow_rows`):
  `Dividendos al accionista = dividendos pagados` y
  `FCFE = CFO − capex + emisión de deuda − amortización`, histórico y forecast
  en la misma fila. Signos del contrato C11: los pagos son magnitud positiva
  y la fórmula los resta.
- **Terminal de equity normalizado** (`ModelStyler.equity_terminal_block`):
  `flujo terminal = NI_n × (1+g) × (1 − g/ROE terminal)` y
  `TV = flujo / (ke − g)`. Es la misma matemática que el terminal de la firma
  (NI, ke y ROE en lugar de NOPAT, WACC y RONIC) y por la misma razón:
  perpetuar el payout del último año explícito arrastra la reinversión de una
  empresa que todavía crece. `ROE terminal` es `supuesto` en Assumptions;
  default = link a ke. Un solo terminal sirve a DDM y FCFE: en estado estable
  todo el FCFE se distribuye. Consecuencia: **DDM y FCFE solo difieren en el
  tramo explícito**; si divergen, la política de payout explícita se aparta de
  la capacidad de pago.
- **Valor** (`ModelStyler.equity_value_block`): VP de flujos explícitos, VP del
  terminal, valor del equity, valor por acción y % del valor en el terminal.
  Sin resta de deuda neta. El DDM marca `REVISAR` si sus dividendos explícitos
  suman ≤ 0 (método mal activado en el perfil). Esa fila vive en `Annual` y se
  enlaza a una fila de `Checks` (`=Annual!<celda>`) para que el reporte por
  código la liste como `[aviso]`: el escáner de checks vivos solo lee
  `Checks` y `Cover`.
- ke: link a la celda de costo de capital del bloque WACC.
- Two-stage = horizonte explícito + terminal. El H-model ya no forma parte del
  menú: con flujos del modelo, el crecimiento que se desvanece va dentro del
  horizonte explícito (lección D4: alargar el horizonte, no otra etapa cerrada).

### Val_NAV_AFFO (condicional — FIBRAs)
- FFO = NI + D&A inmobiliaria ± partidas no recurrentes; AFFO = FFO − capex de
  mantenimiento − comisiones lineales. Convención local de ajustes FIBRA:
  [VERIFICAR — contribución bienvenida].
- NAV: cap rate del analista sobre NOI forward; sensibilidad cap rate obligatoria.

### Val_SOTP (condicional)
- Un bloque por segmento (segmentos del issuer-profile); método por segmento
  elegido por el analista; descuento de holding como supuesto explícito, no
  escondido en el múltiplo.

### Val_RI (v2 — inactivo)
- Especificación: V₀ = B₀ + Σ RI descontado; RI = (ROE − r) × B; justified
  P/B = (ROE − g)/(r − g); continuing RI con ω.
- Exige clean surplus — verificar antes de activar. Se activa cuando el mapeo
  Anexo 33 (bancos) entre al alcance.

### Sensitivity + Summary
- **Cuadrículas con fórmula cerrada por celda, sin data table de Excel**
  (`ModelStyler.sensitivity_grid`): openpyxl no puede escribir un data table y
  la tab se entregaba con encabezados y celdas vacías. Cada celda calcula
  `(VP de los FCFF explícitos + VP del terminal normalizado − deuda neta) /
  acciones` como función del encabezado de su fila y de su columna
  (`dcf_per_share_formula`; los exponentes salen de la fila del DCF, así que
  respeta stub y mid-year). Encabezados: el central enlaza al valor base del
  modelo y los demás suman o restan el paso (`SENS_WACC_STEP`, `SENS_G_STEP`:
  constantes con nombre en Assumptions). Variables típicas: WACC × g, cap rate
  × NOI, múltiplo × EBITDA.
- **Verificación obligatoria:** la celda central reproduce el valor por acción
  del modelo — check de fórmula viva en `Checks` (`center_check_formula`).
- Football field: rango por método activo, precio actual como línea, precio objetivo
  del analista marcado — el precio objetivo es SUPUESTO del analista informado por
  los métodos, jamás un promedio automático.

## Principio de convergencia (check D5)

DDM, FCFE y RI convergen bajo supuestos idénticos. Divergencia extrema entre métodos
activos = casi siempre un supuesto inconsistente entre métodos (g, payout, ROE,
r); con apalancamiento cambiante en el horizonte, FCFE y FCFF divergen
legítimamente. El modelo la reporta y el debate la discute; nunca se resuelve
promediando.

D5 es fórmula viva en `Checks` (`method_divergence_formula`): dispersión
`(MAX − MIN) / MEDIANA` de los valores por acción numéricos de los métodos
activos contra `METHOD_DIVERGENCE_MAX` (constante con nombre, default 0.2).
Entradas: el DCF entra UNA vez, con su valor por acción del terminal Gordon
(no también el del múltiplo de salida, que ya cruza D4); DDM y FCFE con el
suyo. Todos los métodos dividen entre la MISMA celda de acciones diluidas: un
denominador distinto fabrica divergencia. Avisa con `REVISAR`, nunca bloquea.
El diagnóstico, `method_consistency_block`, va en `Annual` debajo de la última
sección de equity (no en `Checks`) y muestra las fuentes típicas: g terminal vs
ROE × retención, ke de equity vs ke del WACC, payout terminal implícito vs
último explícito y deuda neta/EBITDA al inicio y al final del horizonte.
