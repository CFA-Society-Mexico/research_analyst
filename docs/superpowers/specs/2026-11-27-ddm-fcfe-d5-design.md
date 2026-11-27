# Diseño — DDM, FCFE y check D5 (métodos de equity en el builder)

Fecha: 2026-11-27 · Aprobado por Alan en brainstorming (P1 = flujos desde el
modelo · P2 = Gordon normalizado del lado equity, sin H-model · P3 = D5 avisa,
no bloquea · P4 = secciones de `Annual`, no tabs).
Subproyecto 1 de 3 de "métodos condicionales con código" (2 = SOTP,
3 = NAV/AFFO tras verificar la convención AFFO de FIBRAs).

## El problema

`valuation-conventions.md` ofrece DDM y FCFE, `issuer-profile.yaml` los activa
(`valuation_methods.ddm`, `.fcfe`) y `model-spec.md` les reserva lugar, pero el
builder no tiene un solo bloque que los construya ni un test que los cubra
(verificado en `origin/main`, 2026-11-27: `tools/xlsx_styler.py` y
`tools/xlsx_formulas.py` no mencionan DDM ni FCFE). Activar uno hoy obliga al
agente a escribir las fórmulas a mano: justo lo que la regla central prohíbe.
D5 (convergencia de métodos) es doctrina sin código: no hay contra qué medir.

## Decisiones de brainstorming

| # | Decisión | Elección | Descartada y por qué |
|---|---|---|---|
| P1 | De dónde salen los flujos | **Del modelo**: dividendos y FCFE como filas de `Annual`, histórico y forecast, igual que el FCFF | Bloques autónomos de libro de texto (D₀, g por etapa como inputs): duplican supuestos y D5 marcaría incoherencias de inputs como divergencias |
| P2 | Terminal | **Gordon normalizado del lado equity**: `NI_n × (1+g) × (1 − g/ROE_terminal) / (ke − g)`. El two-stage es el horizonte explícito + terminal. **H-model sale del menú** | H-model como terminal alterno: dos supuestos más (H, g_S) para lo que ya hace el horizonte explícito; la lección D4 dice extender el horizonte, no meter otra etapa cerrada |
| P3 | D5 | **Aviso** (`REVISAR`), con bloque de diagnóstico | Bloqueo como D11: con apalancamiento cambiante FCFE y FCFF divergen legítimamente; bloquear empuja a forzar supuestos |
| P4 | Dónde viven | **Secciones de `Annual` debajo del DCF** | Tabs `Val_DDM` / `Val_FCFE`: cada fórmula cruzaría a `Annual!` y la fila de exponentes podría desalinearse en silencio |

## D1 — Componentes

Layout de `main` (fachada `xlsx_builder.py`; código por responsabilidad en
módulos hermanos). Bloques nuevos en `ModelStyler` (`tools/xlsx_styler.py`),
fórmulas de una celda en `tools/xlsx_formulas.py`; ambos visibles por la
fachada como todo lo demás.

### `_normalized_terminal(...)` — privada, extraída

La matemática de `gordon_terminal_block` ya es genérica: `NOPAT_n` y
`NI_n` son "flujo base", WACC y ke son "tasa", RONIC y ROE son "retorno del
crecimiento". Se extrae a una privada que recibe las refs y los tres labels.
`gordon_terminal_block` la envuelve con sus labels actuales: **salida idéntica
celda por celda**, y sus tests existentes no se tocan (red del refactor).

### `equity_terminal_block(ws, row, *, ni_ref, g_ref, ke_ref, roe_ref, col=3)`

Tres filas vía `_normalized_terminal`:

| Fila | Label | Fórmula |
|---|---|---|
| 1 | Tasa de reinversión terminal del equity (g / ROE terminal) | `=IF(roe<=0,"n/s",g/roe)` |
| 2 | Flujo terminal al accionista (NI_n x (1+g) x (1 - g/ROE)) | `=IF(roe<=0,"n/s",ni*(1+g)*(1-g/roe))` |
| 3 | TV del equity (flujo terminal / (ke - g)) | `=IF(OR(ke<=g,NOT(ISNUMBER(flujo))),"n/s",flujo/(ke-g))` |

Devuelve `reinvestment`, `flow`, `tv`, `next_row`. Un solo terminal sirve a
DDM y FCFE: en estado estable todo el FCFE se distribuye, así que el
dividendo terminal y el FCFE terminal coinciden. Consecuencia que el reporte
debe decir: **DDM y FCFE solo difieren en el tramo explícito**; una
divergencia entre ellos significa que la política de payout explícita se
aparta de la capacidad de pago.

`ROE_terminal` es input `supuesto` de Assumptions; default = link a ke (el
crecimiento terminal no crea valor), mismo criterio que RONIC = WACC.

### `equity_value_block(ws, row, *, title, flow_range, exponent_range, ke_ref, tv_ref, tv_exponent_ref, shares_ref, require_positive_flows=False, stub_flow_ref=None, stub_exponent_ref=None, col=3)`

| Fila | Label | Fórmula |
|---|---|---|
| 1 | {title}: VP de los flujos explícitos | `=SUMPRODUCT(flows/(1+ke)^exps)`; con stub, `=stub/(1+ke)^stub_exp+SUMPRODUCT(flows/(1+ke)^exps)` |
| 2 | {title}: VP del terminal | `=IF(ISNUMBER(tv),tv/(1+ke)^tv_exp,"n/s")` |
| 3 | {title}: valor del equity | `=IF(ISNUMBER(pv_tv),pv_exp+pv_tv,"n/s")` |
| 4 | {title}: valor por acción | `=IF(OR(NOT(ISNUMBER(eq)),shares<=0),"n/s",eq/shares)` |
| 5 | {title}: % del valor en el terminal | `=IF(OR(NOT(ISNUMBER(eq)),eq<=0),"n/s",pv_tv/eq)` |
| 6 (solo si `require_positive_flows`) | {title}: flujos explícitos positivos | `=IF(SUM(flows)<=0,"REVISAR","OK")`; con stub, `=IF(stub+SUM(flows)<=0,...)` |

- Sin resta de deuda neta: los flujos ya son del accionista.
- `exponent_range` y `tv_exponent_ref` son las celdas de la fila de
  exponentes del DCF: mid-year se hereda, no se recalcula.
- Stub (modo `quarterly`, último trimestre reportado ≠ 4Q): `stub_flow_ref`
  es la suma de las columnas de trimestres E del año en curso en `Operating`
  (dividendos o componentes del FCFE), escrita por el llamador como fórmula;
  `stub_exponent_ref` es el exponente del stub del DCF. Van los dos o
  ninguno (`ValueError`). Con stub, `flow_range` / `exponent_range` cubren
  solo los años completos posteriores: el FY completo contaría flujos ya
  pagados. Sin stub, la salida es idéntica a la de antes.
- DDM llama con `require_positive_flows=True`: una emisora sin dividendos
  explícitos tiene el método mal activado en el perfil.
- Devuelve `pv_explicit`, `pv_tv`, `equity`, `per_share`, `tv_share`,
  `next_row` (y `flag` si aplica).

### `equity_flow_rows(ws, row, first_col, n_cols, *, dividends_row, cfo_row, capex_row, issuance_rows, repayment_rows)`

Dos filas en TODO el horizonte de `Annual` (una serie = una fila, F11),
histórico y forecast — el flujo realizado es sanity visible, como el FCFF:

| Fila | Label | Fórmula por columna `c` |
|---|---|---|
| 1 | Dividendos al accionista | `={c}{dividends_row}` |
| 2 | FCFE (CFO - capex + emisión de deuda - amortización) | `={c}{cfo} - {c}{capex} + Σ {c}{issuance} - Σ {c}{repayment}` |

**Signos — contrato del repo, no elección de este diseño:** los canónicos
guardan los pagos (capex, dividendos, amortización de deuda) como MAGNITUD
POSITIVA y la fórmula del modelo los resta (`xbrl-csv-rules.md` §2, check
C11). Por eso los dividendos entran sin voltear y capex y amortización se
restan. `issuance_rows` / `repayment_rows` son listas: deuda y bonos pueden
venir en líneas separadas. Devuelve los números de fila de ambas filas (`dividends_row`, `fcfe_row`) y `next_row`: el llamador arma el rango del tramo que descuenta.

### `method_divergence_formula(values, max_ratio)` — `xlsx_formulas.py`

Una celda: dispersión `(MAX − MIN) / MEDIANA` sobre los valores por acción
de los métodos activos (DCF FCFF siempre; DDM y FCFE si existen).

- Solo cuentan los numéricos (un método en `n/s` se excluye, no rompe D5).
- Menos de 2 numéricos ⇒ `n/a`; mediana ≤ 0 ⇒ `n/a`.
- Dispersión > `max_ratio` ⇒ `REVISAR`; si no, `OK`.
- `max_ratio` = constante con nombre `METHOD_DIVERGENCE_MAX` en Assumptions
  (`supuesto`, default del plugin 0.2). Sin literales (S4) y sin operaciones
  de arreglo fuera de `SUMPRODUCT` (F21); el conteo de numéricos usa `COUNT`.

### `method_consistency_block(ws, row, *, ...)` — diagnóstico de D5

Debajo de la fila D5, cero inputs nuevos, una fila por fuente típica de
divergencia:

| Fila | Qué muestra |
|---|---|
| g terminal vs crecimiento sostenible | g del terminal contra ROE × retención del último año explícito |
| ke de los métodos de equity vs ke del WACC | diferencia (debe ser 0 si ambos leen la misma celda) |
| Payout terminal implícito vs último payout explícito | `1 − g/ROE_terminal` contra dividendos ÷ NI del año n |
| Trayectoria de apalancamiento | deuda neta / EBITDA del primer y del último año explícito |

Es material de entrevista de cierre, no veredicto.

## D2 — Flujo de datos y activación

```
Operating CF (dividendos pagados, CFO, capex, emisión/amortización de deuda)
  -> Annual: agregados FY por fórmula (sin cambio)
  -> equity_flow_rows (histórico + forecast)
  -> equity_terminal_block (una vez) -> equity_value_block (DDM y/o FCFE)
  -> D5 en Checks (method_divergence_formula) + method_consistency_block
```

- `valuation_methods.ddm` / `.fcfe` del perfil deciden si la sección se
  construye. Con ambos en `false` no se escribe nada y la fila D5 dice
  `n/a (un solo metodo activo)`.
- ke: parámetro `ke_ref`. El agente enlaza la celda de costo de capital del
  bloque WACC que ya escribe; este diseño no agrega helper de WACC.

## D3 — D5 en `Checks` y en el reporte

- Fila D5 en `Checks` con `method_divergence_formula`: `OK` / `REVISAR` /
  `n/a`.
- Hoy el escáner de checks vivos (`live_checks_finding` en
  `tools/xlsx_audit.py`, sobre `Checks` y `Cover`) solo falla con texto que empieza con
  `ERROR` o error de Excel: `REVISAR` pasa en silencio. Cambio: una celda
  viva que vale `REVISAR` se junta aparte y, si no hay fallas, el `Finding`
  sale con `warning=True` (`[aviso]`, lista las celdas, no cambia el exit
  code). `Error` sigue fallando igual; un aviso nunca tapa una falla ni un
  pendiente.
- La celda de error del `Cover` cuenta `Error`; `REVISAR` no entra en esa
  cuenta (regla documentada en `integrity-checks.md`; la escribe el agente,
  no hay helper que cambiar).

## D4 — Errores (todos por fórmula; nunca un número falso)

| Caso | Resultado |
|---|---|
| ke ≤ g | `n/s` en TV, valor del equity y por acción |
| ROE terminal ≤ 0 | `n/s` |
| acciones ≤ 0 | `n/s` en valor por acción |
| equity ≤ 0 | `n/s` en % del valor en el terminal |
| DDM con Σ dividendos explícitos ≤ 0 | `REVISAR` en el bloque DDM |
| D5 con < 2 valores numéricos | `n/a` |
| D5 con mediana ≤ 0 | `n/a` |
| método en `n/s` | excluido de D5 |

## D5 — Pruebas (TDD)

| Clase | Qué prueba | Excel |
|---|---|---|
| `ValuationBlockStructureTests` (existente, se amplía) | filas y labels de cada bloque; sin literales (S4) y `lint_formula` limpio (F21) en todo lo nuevo; `equity_flow_rows` resta capex y amortización y no voltea dividendos | no |
| Refactor `_normalized_terminal` | los tests actuales de `gordon_terminal_block` pasan sin tocarse | no |
| Escáner de checks vivos (`test_recalc_validate.py` / `test_model_check.py`) | `REVISAR` ⇒ `[aviso]` no bloqueante; `Error` sigue fallando; `Error` + `REVISAR` ⇒ falla | no |
| `ExcelEvaluationTests` (existente, se salta sin Excel) | valor del equity = cálculo a mano en caso de 3 años + TV; ke ≤ g ⇒ `n/s`; D5 con DDM y FCFE de flujos idénticos ⇒ dispersión 0 y `OK`; D5 sobre el umbral ⇒ `REVISAR`; un solo numérico ⇒ `n/a` | sí |

El demo del builder NO se extiende: es el fixture de los checks S/F y no
trae dividendos, capex ni deuda. Los bloques se prueban aislados, como los
de valuación existentes.

## D6 — Documentos que cambian

| Archivo | Cambio |
|---|---|
| `skills/model-standards/references/valuation-conventions.md` | menú: DDM "Gordon normalizado sobre flujos del modelo" (sin two-stage/H-model como métodos aparte); `Val_DDM` ⇒ "Sección DDM de Annual"; sección FCFE nueva; terminal de equity normalizado; D5 con bloque de diagnóstico |
| `skills/model-standards/references/integrity-checks.md` | fila D5: por fórmula, aviso, `METHOD_DIVERGENCE_MAX`, qué excluye |
| `templates/model-spec.md` | `Val_<método>` solo NAV_AFFO / SOTP; §Annual suma secciones DDM y FCFE |
| `skills/model-standards/SKILL.md` | paso 5: cuándo y con qué bloques se construyen DDM y FCFE |
| `.claude-plugin/plugin.json` | bump de versión (minor) sobre la de `main` al cerrar |

## Fuera de alcance

- SOTP (subproyecto 2) y NAV/AFFO (subproyecto 3).
- Helper de WACC.
- Reverse DDM (g implícita del lado equity) y sensibilidad ke × g del DDM.
- Extender el demo del builder.
- Residual income (sigue especificado e inactivo hasta el mapeo Anexo 33).
