# Reglas del CSV de `xbrl_fetch` al mapear a los canónicos

`tools/xbrl_fetch.py` escribe `xbrl_facts_<TICKER>.csv` con los valores **tal
como los reporta XBRL**. Antes de pasar una fila a `canonical_*.csv`, aplica
estas tres reglas. Las columnas `accession` y `sign` van al final del CSV.

## 1. Un vintage por estado y periodo

Las líneas de un estado (IS, BS, CF) en un periodo salen del mismo
`accession`: el presentado más recientemente que cubre ese estado y periodo
(al menos 80% de las líneas del filing mejor cubierto). Antes cada línea
tomaba el `filed` más reciente por su cuenta y un balance podía mezclar la
presentación original con la re-expresada: el FY2017 de TSMC quedó descuadrado
por 1,246 NT$ mn sin que ninguna cifra estuviera mal.

- Tag `vintage distinto: <accession>` = línea que el filing primario no trae y
  se completó de OTRO filing. Se reporta al analista y se concilia contra el
  filing antes de entrar al canónico; nunca se mezcla en silencio.
- El 4Q derivado (FY − acumulado a 3Q) resta dos filings distintos por
  construcción; eso no se marca.

## 2. Signo

El tool NO invierte signos. `sign = salida(+)` marca un pago que XBRL reporta
en positivo (capex, dividendos, recompras, impuestos pagados, amortización de
deuda, compras de valores). **Convención de los canónicos** (contrato en
`templates/coverage-tree.md`): el pago se guarda igual, como MAGNITUD POSITIVA,
y quien resta es la fórmula del modelo. Al mapear:

- **no voltees el signo**: una fila `salida(+)` entra al canónico con el valor
  de XBRL. Si capturas del filing y el estado muestra el pago entre paréntesis,
  guarda la magnitud. Check C11 falla un pago reservado en negativo;
- los subtotales (`cf_cfi`, `cf_cff`, cambio neto) y las líneas netas entran
  con su signo natural;
- **puentea contra el filing**: en el modelo, la suma de componentes con el
  signo de la estructura (`CFI = −capex − compras de valores + …`, `CFF =
  −recompras − dividendos + …` — `skills/model-standards/references/excel-practices.md`)
  = subtotal observado. Sin ese puente el FCF sale inflado.

## 3. `_alt` e IFRS

- `_alt` es otro concepto, no un duplicado: `is_ni` y `bs_equity` son de la
  controladora; `is_ni_alt` y `bs_equity_alt`, el total con participación no
  controladora (el balance cuadra con el total).
- `is_other_income_finance` (`FinanceIncome`) va a otros ingresos, NO a
  intereses; el ingreso por intereses es `is_interest_income`
  (`RevenueFromInterest`).
- `cf_fx_effect` entra al roll de caja: CFO + CFI + CFF + efecto cambiario =
  cambio neto. Sin él el roll no cierra en una emisora con operación en otra
  moneda.
- Los 6-K de una emisora extranjera normalmente no traen XBRL: la historia
  IFRS que sale de companyfacts es anual; los trimestres se capturan del filing.

## 4. Captura incompleta: conceptos sin mapear

Cada emisora elige sus conceptos, y uno que no está en el mapa no emite filas.
El tool lo reporta en dos líneas de consola — no las ignores:

- `[i] N conceptos con datos sin mapear` — siempre. Con `--unmapped` escribe
  `xbrl_unmapped_<TICKER>.csv` (concepto, etiqueta, unidades, número de hechos,
  primer y último cierre), lo vigente primero.
- `[aviso] el payload trae cierres anuales posteriores al último capturado` —
  el CSV termina antes que los filings de la emisora: falta un año entero.

- `[i] último cierre anual en companyfacts` — siempre. companyfacts puede ir
  atrasado respecto a los filings: un año ya presentado que no llega al
  payload se captura desde el filing, con cita.

Prueba en vivo con TSM (2026-11-25): la captura salió sin D&A, sin impuestos
pagados y sin deuda porque TSMC los reporta bajo otros conceptos (ya en el
mapa), y sin FY2025 porque companyfacts terminaba en 2024-12-31 con el 20-F
de FY2025 ya presentado. Antes de capturar a mano una línea que falta, corre
con `--unmapped`: si la emisora la reporta bajo otro concepto, se agrega al
mapa de su taxonomía en `tools/xbrl_fetch.py` (con prueba), no al canónico de
una sola cobertura.

## 5. Líneas que se suman

Dos conceptos que se SUMAN llevan canon propio y esta skill los agrega, con
gate, al mapear: `cf_depreciation` + `cf_amortisation` = D&A; `is_sm` +
`is_ga` (+ `is_rd`) = gastos de operación (contrastar con `is_opex_total`);
`bs_bonds_lt` + `bs_borrowings_lt` = deuda de largo plazo; `cf_bonds_issued` +
`cf_borrowings_lt_issued` = deuda emitida. Solo comparten canon los conceptos
que son alternativas de una misma línea (un emisor usa uno u otro).

**Que dos líneas tengan canon propio no prueba que se sumen: se puentea
contra el balance antes de agregar.** Caso real (20-F de TSMC, FY2024):
`bs_borrowings_lt_current` es la línea del balance "Long-term liabilities -
current portion" (59,858 NT$ mn) y YA incluye la porción circulante de los
bonos; `bs_bonds_st` (57,148, de la nota de bonos) es un componente de ella,
no un sumando: 59,858 = 57,148 bonos + 2,710 préstamos bancarios. Sumarlas
duplica 57,148 de deuda. Regla: una cifra que solo aparece en una nota se
trata como desglose de la línea del estado hasta demostrar lo contrario.

Puente FY2024 de las líneas IFRS contra el 20-F de TSMC (2026-11-25): las 34
líneas revisadas coinciden con la cifra y la etiqueta de fila del estado
— p. ej. `cf_depreciation` = "Depreciation expense" del flujo de efectivo,
`bs_ap` = "Accounts payable", `bs_ar` = "Notes and accounts receivable, net",
`is_sm` = "Marketing", `is_other_income_finance` = "Other income".
