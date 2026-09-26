# Valuation Conventions — menú de métodos y pestañas deterministas

Doctrina: **el modelo arma la estructura y las fórmulas; el analista pone cada
supuesto; nada se calcula por generación libre.** Usuario en el loop en cada supuesto.
El football field agrega SOLO métodos activos en `issuer-profile.yaml`.

## Menú y activación

| Activación | Método | Supuestos del analista |
|---|---|---|
| **Core — siempre** | DCF FCFF multi-stage, terminal DUAL | WACC (componentes), g terminal, múltiplo de salida, horizonte |
| **Core — siempre** | Comps (EV/EBITDA, P/E, P/B, EV/Sales) | universo (viene de industry-analysis), múltiplo aplicable |
| **Core — siempre** | Sensibilidades (data tables 2 vars) + football field | rangos de sensibilidad |
| Dividendera estable | DDM: Gordon / two-stage / H-model | g por etapa, payout |
| Apalancamiento estable | FCFE directo | r equity, trayectoria de deuda neta |
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
- WACC: componentes visibles — rf de `Macro` (macro-view), ERP de `Macro`, beta
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
- **Terminal dual obligatorio:** Gordon (g) Y exit multiple, lado a lado.
  Cruce: Gordon ⇒ múltiplo implícito; exit multiple ⇒ g implícita. Ambos visibles.
  Divergencia grande = check D4: revisar supuesto con el analista, no promediar.
  La g implícita de CUALQUIER terminal (exit multiple o reverse DCF) se despeja
  con el FCFF del último año explícito, nunca con uno proyectado con la g del
  analista:
  `g implícita = (TV × WACC − FCFF_n) / (TV + FCFF_n)`, que es el despeje de
  `TV = FCFF_n × (1+g) / (WACC − g)`.
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
  - g implícita = (TV implícita × WACC − FCFF_n) / (TV implícita + FCFF_n).
    La forma anterior, `WACC − FCFF_{n+1} / TV`, usaba `FCFF_{n+1} = FCFF_n ×
    (1 + g del analista)`: la g "del mercado" salía contaminada por el mismo
    supuesto que el bloque existe para contrastar.
  - Tercera columna junto al cruce de terminales: "el mercado descuenta g = X; tú
    supones g = Y" (check D4b — debate, nunca bloqueo).

### Val_Comps
- Múltiplos calculados POR FÓRMULA desde `comps/*.yaml`: EV = mkt cap + deuda − caja
  + minoritarios + preferentes; cada componente del snapshot, con fuente.
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

### Val_DDM (condicional)
- Gordon: V₀ = D₁/(r−g). Two-stage y H-model con etapas explícitas.
- Consistencia: g = ROE × (1−payout) visible como check suave.

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
- Data tables aisladas; variables típicas: WACC × g, cap rate × NOI, múltiplo × EBITDA.
- Football field: rango por método activo, precio actual como línea, precio objetivo
  del analista marcado — el precio objetivo es SUPUESTO del analista informado por
  los métodos, jamás un promedio automático.

## Principio de convergencia (check D5)

DDM, FCFE y RI convergen bajo supuestos idénticos. Divergencia extrema entre métodos
activos = supuesto inconsistente entre pestañas (g, payout, ROE, r). El modelo la
reporta y el debate la discute; nunca se resuelve promediando.
