# Ratios & Analytics — spec de la tab `Ratios` y del bloque `Sch: WC`

Sección core de `Operating` y `Annual` (de `Model` en modo annual), siempre
activa. Solo fórmulas sobre los 3
estados — **cero inputs del analista**. Histórico + forecast; una fórmula por
fila y tramo (check S5). Denominadores de balance: promedio de periodo
((inicio+fin)/2), consistente en TODAS las filas.

**Generación: SOLO vía `ModelStyler.build_ratios`** (registro canon→fila del
build; completitud y unicidad auditadas por check F13 contra
`REQUIRED_RATIO_LABELS`). Este documento es el contrato legible de lo que ese
código implementa. Lo marcado *pendiente* abajo está especificado pero
`build_ratios` todavía no lo escribe, y F13 no lo exige.

**Ventanas móviles (modo trimestral)**: UDM/LTM = EXACTAMENTE 4 trimestres
(t−3 … t) — ni 5 (columna −4 a la actual: el bug del smoke #4, infla ~25%) ni
"trimestre × 4". Toda razón anualizada sobre base trimestral (DSO/DIO/DPO,
deuda/EBITDA, cobertura) usa flujos UDM y stocks promedio de los 5 cierres
t−4..t: apertura y cierre del periodo de 12 meses, igual que (inicio+fin)/2
en modo anual. La fórmula de la ventana es idéntica en toda la fila (S5).
En una hoja con header trimestral, `build_ratios` EXIGE `window` y falla si no
se pasa, porque no puede saber qué trae `ref`: con `window=4` el caller pasa
flujos de un trimestre y el builder los suma (t−3..t), promedia los saldos de
t−4..t, usa `DAYS_YEAR` y marca cada fila con `[UDM]` (las 4 primeras columnas
quedan vacías); con `window=1` el caller ya pasa filas UDM y el builder no suma
nada. En hojas anuales el default es `window=1`: saldos (inicio+fin)/2.

## Bloque A — DuPont

3 factores:

    ROE = (NI/Ventas) × (Ventas/Activos) × (Activos/Capital)
          margen neto   rotación          apalancamiento

5 factores:

    ROE = (NI/EBT) × (EBT/EBIT) × (EBIT/Ventas) × (Ventas/Activos) × (Activos/Capital)
          carga fiscal  carga int.  margen op.     rotación           apalancamiento

La identidad ROE directo = ROE DuPont 5 se cumple por álgebra (los factores
se cancelan), así que restarlas no detecta nada: C6 quedó derogado. El valor
del bloque está en leer cada factor por separado.

## Bloque B — ROIC y economic profit

    NOPAT = EBIT × (1 − t efectiva)         ← t efectiva por fórmula desde IS
    Capital invertido = deuda total + capital contable − caja y equivalentes
    ROIC = NOPAT / capital invertido PROMEDIO      (promedio, como todo
    Economic profit = (ROIC − WACC) × capital     denominador de balance)
                      invertido promedio

**Guarda obligatoria**: capital invertido promedio ≤ 0 (caja > deuda+capital —
caso real en emisoras con caja neta enorme) ⇒ ROIC y economic profit reportan
`n/s` (no significativo). Jamás un porcentaje inflado: el smoke #5 mostró
ROIC de 369.6% con capital invertido de 34,838 — matemáticamente correcto,
financieramente ruido.

WACC referenciado de la sección DCF de `Annual` (link verde entre hojas). Lectura de tesis: spread
ROIC−WACC positivo sostenido = evidencia numérica del moat que industry-analysis
afirma cualitativamente; la entrevista de cierre lo confronta.

## Bloque C — Ratios estándar

| Grupo | Filas | Estado en `build_ratios` |
|---|---|---|
| Actividad | rotación de inventarios, CxC, CxP, activos totales | rotación de activos; inventarios, CxC y CxP se leen como DIO/DSO/DPO (Bloque D) |
| Liquidez | corriente, quick, cash ratio | corriente y quick; cash ratio *pendiente* |
| Solvencia | deuda/capital, deuda/EBITDA, deuda neta/EBITDA, cobertura de intereses (EBIT/gasto fin.), cobertura de cargos fijos | deuda/EBITDA y cobertura de intereses; deuda/capital y deuda neta/EBITDA *pendientes*; cargos fijos *pendiente* (requiere un canon de arrendamientos que la captura no trae) |
| Rentabilidad | margen bruto, operativo, EBITDA, neto; ROA; ROE (link a Bloque A) | bruto, operativo, neto y ROE; margen EBITDA y ROA *pendientes* |

## Bloque D — Ciclo de conversión de efectivo

    DIO = inventario prom. / COGS × <días>
    DSO = CxC prom. / ventas × <días>
    DPO = CxP prom. / COGS × <días>
    CCC = DIO + DSO − DPO

**Regla de ventana (crítica en modelo trimestral)**: numerador = STOCK
promedio, denominador = FLUJO; ambos deben cubrir la MISMA ventana. En hojas
trimestrales: o el flujo es UDM (12 meses) con `DAYS_YEAR`, o el flujo del
trimestre con `DAYS_QUARTER`. Mezclar flujo trimestral con `DAYS_YEAR` infla
los días ~4×. `build_ratios` recibe `days_ref` explícito — no lo hardcodea.

Histórico: calculado. Forecast: referencia los días del bloque `Sch: WC` — el CCC
forecast es output del schedule, no fila independiente (check C7).

## Bloque E — Apalancamiento operativo/financiero

    DOL = %ΔEBIT / %ΔVentas       (sobre periodos del modelo)
    DFL = EBIT / (EBIT − gasto financiero neto)
    DTL = DOL × DFL

Filas de lectura, no drivers. Valor principal: emisoras cíclicas apalancadas.
*Pendiente en `build_ratios`*: hoy solo se escribe DFL. DOL y DTL comparan
contra el año previo; en hoja trimestral eso exige 8 trimestres de ventana y
más huecos iniciales de los que F15 admite.

## Bloque F — Crédito y screening (solo `issuer_type: non_financial`)

### Altman Z'' (verificado — Altman, Hartzell & Peck 1995)

    Z'' = 6.56·(WC/TA) + 3.26·(RE/TA) + 6.72·(EBIT/TA) + 1.05·(BVE/TL)

- BVE = capital contable a **libros** (no mercado — esa es la diferencia del Z''
  vs el Z original; aplica a privadas, no-manufactureras y EM).
- Zonas sobre Z'' sin constante: seguro > 2.60; gris 1.10–2.60; distress < 1.10.
- Fila adicional EM score = Z'' + 3.25 — estandariza para que 0 ≈ bono D;
  informativa (equivalente de rating), las zonas se leen sobre Z'' sin constante.

### Piotroski F-score (verificado — Piotroski 2000)

Nueve señales binarias (1 si cumple, 0 si no); cada una fila auxiliar agrupada
(outline), suma visible. ROA y CFO escalados por activos totales **iniciales**.

| # | Grupo | Señal = 1 si |
|---|---|---|
| 1 | Rentabilidad | ROA > 0 (NI antes de extraordinarios / TA inicial) |
| 2 | Rentabilidad | CFO > 0 (CFO / TA inicial) |
| 3 | Rentabilidad | ΔROA > 0 vs año anterior |
| 4 | Rentabilidad | CFO > ROA (accrual: caja respalda la utilidad) |
| 5 | Apalancamiento/liquidez | razón deuda LP / TA promedio bajó |
| 6 | Apalancamiento/liquidez | razón corriente subió |
| 7 | Apalancamiento/liquidez | sin emisión de capital común en el año (no dilución) |
| 8 | Eficiencia | margen bruto subió |
| 9 | Eficiencia | rotación de activos (ventas / TA inicial) subió |

Bancos/aseguradoras: bloque omitido (v2, junto con Anexo 33 y residual income).
Son **screening, no veredicto**: filas informativas, jamás gate.
*Pendiente en `build_ratios`*: Altman Z'' y Piotroski están especificados y
verificados, pero el código todavía no los escribe.

## Bloque G — Calidad de utilidades

    CFO/NI                                   (por periodo)
    NOA = (activos − caja) − (pasivos − deuda)
    Accruals ratio (BS) = ΔNOA / NOA promedio

*Implementación actual*: `build_ratios` escribe un proxy de accruals por
flujo, (NI − CFO) / activos promedio, no el ratio de balance ΔNOA / NOA
promedio. D9 se evalúa con el proxy hasta que se implemente el de NOA.

Alimenta check D9 (aviso, no bloqueo). Convención de minoritarios: son
FINANCIAMIENTO — dentro de capital, fuera de pasivos operativos de NOA
(consistente con capital invertido del Bloque B, que usa capital contable total).

---

## Bloque `Sch: WC` (sección Schedules de `Operating` — core, siempre existe)

Working capital por días — arregla que ΔWC no tenía driver explícito (espíritu D1):

- DIO / DSO / DPO forecast = **inputs del analista** (azul, en Assumptions, como
  todo driver); histórico calculado al lado como referencia.
- Inventario = DIO × COGS / `DAYS_YEAR`, con COGS de la misma ventana anual
  (UDM en `Operating`); o DIO × COGS del trimestre / `DAYS_QUARTER`. Igual
  CxC con ventas y CxP con COGS. Es la regla de ventana del Bloque D: con COGS
  trimestral y 365 el inventario sale ~4 veces menor. Nunca el literal 365
  (S4).
- ΔWC del CF y del FCFF sale de este bloque, nunca de una fila suelta.
- driver-inventory (pasada design): el driver-map DEBE incluir la fila de WC en
  "Drivers de costo y capex" — o justificarla en "Líneas sin driver".
