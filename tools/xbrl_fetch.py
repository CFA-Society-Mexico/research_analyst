"""SEC XBRL companyfacts -> serie historica completa (anual + trimestral).

Resuelve la captura masiva de trimestres sin parsear HTML: la API publica
``data.sec.gov/api/xbrl/companyfacts`` trae TODOS los valores reportados de la
emisora con su accession: 10-K y 10-Q de emisoras domesticas, 20-F y 40-F de
emisoras extranjeras. Este tool baja los conceptos mapeados a lineas canon y
escribe un CSV largo que statement-mapper convierte (con gate) en los
canonical_*.csv de model/inputs/.

Marcos soportados: taxonomia ``us-gaap`` e ``ifrs-full`` (emisoras IFRS como
TSM, UMC o GFS), en la moneda de reporte de la emisora (TSM reporta en TWD,
GFS en USD). Por default ambas se detectan del payload; ``--taxonomy`` y
``--currency`` las fijan. Las emisoras 20-F suelen traer solo anuales: sus
6-K trimestrales normalmente no llevan XBRL.

Reglas del CSV que el consumidor debe conocer:

- Un solo vintage por estado y periodo. Todas las lineas de un estado (is_,
  bs_, cf_) en un periodo salen del MISMO accession: el mas reciente que cubre
  el periodo. Mezclar original y re-presentado linea por linea descuadra el
  balance sin que ninguna cifra este "mal". Una linea que ese accession no
  reporta se completa del siguiente mas reciente y queda marcada en ``tag``
  (``vintage distinto: <accession>``); columna ``accession`` en cada fila.
- Valores como se reportan. XBRL trae varias SALIDAS de efectivo en positivo
  (capex, dividendos, impuestos pagados, amortizacion de deuda). El tool no
  cambia signos: marca esas filas con ``sign = salida(+)`` y statement-mapper
  aplica la convencion del canonico con puente contra el filing.
- Captura incompleta visible. Un concepto fuera del mapa no emite filas: la
  consola dice cuantos conceptos con datos quedaron sin mapear y avisa si el
  payload trae un anio posterior al ultimo capturado. ``--unmapped`` escribe
  la lista en ``xbrl_unmapped_<TICKER>.csv``.

Uso:
    python tools/xbrl_fetch.py AAPL --dest <raiz>/AAPL/model/inputs \
        --ua "Nombre correo@dominio.com"
    python tools/xbrl_fetch.py TSM --dest <raiz>/TSM/model/inputs \
        --taxonomy ifrs-full --currency TWD --ua "Nombre correo@dominio.com"

User-Agent obligatorio (--ua o env SEC_EDGAR_UA), igual que sec_fetch.
Consola ASCII-only.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable, Optional, Sequence

FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"

# Concepto US GAAP -> linea canon. Lista ampliable; un concepto ausente en la
# emisora simplemente no emite filas (se reporta al final).
CONCEPT_MAP = {
    "RevenueFromContractWithCustomerExcludingAssessedTax": "is_ns_total",
    "Revenues": "is_ns_total_alt",
    "CostOfGoodsAndServicesSold": "is_cogs_total",
    "GrossProfit": "is_gross",
    "ResearchAndDevelopmentExpense": "is_rd",
    "SellingGeneralAndAdministrativeExpense": "is_sga",
    "OperatingIncomeLoss": "is_ebit",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest": "is_ebt",
    "IncomeTaxExpenseBenefit": "is_tax",
    "NetIncomeLoss": "is_ni",
    "EarningsPerShareDiluted": "is_eps_diluted",
    "Assets": "bs_ta",
    "AssetsCurrent": "bs_ca",
    "LiabilitiesCurrent": "bs_cl",
    "Liabilities": "bs_tl",
    "StockholdersEquity": "bs_equity",
    "CashAndCashEquivalentsAtCarryingValue": "bs_cash",
    "InventoryNet": "bs_inv",
    "AccountsReceivableNetCurrent": "bs_ar",
    "AccountsPayableCurrent": "bs_ap",
    "RetainedEarningsAccumulatedDeficit": "bs_re",
    "LongTermDebtNoncurrent": "bs_debt_lt",
    "LongTermDebtCurrent": "bs_debt_st",
    "CommercialPaper": "bs_commercial_paper",
    # --- Flujo de efectivo: SUBTOTALES primero. Sin CFI/CFF/cambio neto el
    # roll de caja no cierra (inicio + cambio != cierre) y el modelo hereda
    # un flujo de inversion incompleto — el bug del smoke #5.
    "NetCashProvidedByUsedInOperatingActivities": "cf_cfo",
    "NetCashProvidedByUsedInInvestingActivities": "cf_cfi",
    "NetCashProvidedByUsedInFinancingActivities": "cf_cff",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect": "cf_net_change",
    "CashAndCashEquivalentsPeriodIncreaseDecrease": "cf_net_change_alt",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents": "cf_cash_eop_incl_restricted",
    # Efecto cambiario sobre el efectivo: sin el, CFO + CFI + CFF != cambio
    # neto y el roll de caja no cierra en emisoras con operacion en otra moneda.
    "EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents": "cf_fx_effect",
    "DepreciationDepletionAndAmortization": "cf_da",
    "ShareBasedCompensation": "cf_sbc",
    "PaymentsToAcquirePropertyPlantAndEquipment": "cf_capex",
    "PaymentsForRepurchaseOfCommonStock": "cf_buybacks",
    "PaymentsOfDividends": "cf_dividends",
    "PaymentsOfDividendsCommonStock": "cf_dividends_common",
    # Movimiento de valores negociables: para emisoras con tesoreria grande
    # es el mayor flujo despues del operativo; omitirlo descuadra el roll.
    "PaymentsToAcquireAvailableForSaleSecuritiesDebt": "cf_buy_securities",
    "ProceedsFromSaleOfAvailableForSaleSecuritiesDebt": "cf_sell_securities",
    # Vencimientos: el concepto vigente NO lleva sufijo "Debt" (Apple lo usa
    # de FY2007 a FY2025; la variante "...Debt" no existe en sus datos).
    "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities": "cf_mature_securities",
    "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecuritiesDebt": "cf_mature_securities",
    "PaymentsToAcquireBusinessesNetOfCashAcquired": "cf_acquisitions",
    "PaymentsForProceedsFromOtherInvestingActivities": "cf_other_investing",
    # Deuda
    "ProceedsFromIssuanceOfLongTermDebt": "cf_debt_issued",
    "RepaymentsOfLongTermDebt": "cf_debt_repaid",
    "ProceedsFromRepaymentsOfCommercialPaper": "cf_commercial_paper_net",
    "ProceedsFromIssuanceOfCommonStock": "cf_stock_issued",
    "PaymentsRelatedToTaxWithholdingForShareBasedCompensation": "cf_tax_withholding",
}

# Conceptos de la taxonomia anterior a 2018. Sin ellos la historia se corta:
# Apple reporto ventas como SalesRevenueNet de FY2007 a FY2017 y compras/ventas
# de valores sin el sufijo "Debt" hasta FY2018. Van DESPUES del mapa principal:
# dentro de un mismo accession gana el concepto vigente.
LEGACY_CONCEPT_MAP = {
    "SalesRevenueNet": "is_ns_total",
    "PaymentsToAcquireAvailableForSaleSecurities": "cf_buy_securities",
    "ProceedsFromSaleOfAvailableForSaleSecurities": "cf_sell_securities",
    # Efecto cambiario antes de que el flujo reconciliara efectivo restringido.
    "EffectOfExchangeRateOnCashAndCashEquivalents": "cf_fx_effect",
}

# Concepto ifrs-full -> la MISMA linea canon del mapa US GAAP, para que el
# resto del pipeline no distinga marcos. Solo entran conceptos de la taxonomia
# IFRS con equivalencia directa: uno ausente no emite filas, uno mal mapeado
# contamina el modelo. Donde dos conceptos alimentarian la misma linea, el
# segundo va a un canon ``_alt`` (igual que Revenues en US GAAP): nunca se pisan.
IFRS_CONCEPT_MAP = {
    "Revenue": "is_ns_total",
    "CostOfSales": "is_cogs_total",
    "GrossProfit": "is_gross",
    "ResearchAndDevelopmentExpense": "is_rd",
    "SellingGeneralAndAdministrativeExpense": "is_sga",
    "ProfitLossFromOperatingActivities": "is_ebit",
    # Ingreso por intereses = RevenueFromInterest. FinanceIncome NO lo es: el
    # puente FY2024 contra el 20-F de TSMC (issue 35) lo ubico en OTROS
    # ingresos; mapearlo a intereses es el error ingenuo. Canon propio.
    "RevenueFromInterest": "is_interest_income",
    "FinanceIncome": "is_other_income_finance",
    "ProfitLossBeforeTax": "is_ebt",
    "IncomeTaxExpenseContinuingOperations": "is_tax",
    # is_ni = atribuible a la controladora (misma semantica que NetIncomeLoss
    # en US GAAP); el consolidado con participacion no controladora va a _alt.
    "ProfitLossAttributableToOwnersOfParent": "is_ni",
    "ProfitLoss": "is_ni_alt",
    "DilutedEarningsLossPerShare": "is_eps_diluted",
    "Assets": "bs_ta",
    "CurrentAssets": "bs_ca",
    "CurrentLiabilities": "bs_cl",
    "Liabilities": "bs_tl",
    # bs_equity = atribuible a la controladora (como StockholdersEquity); el
    # capital TOTAL, que es el que cierra Activo = Pasivo + Capital, va a _alt.
    "EquityAttributableToOwnersOfParent": "bs_equity",
    "Equity": "bs_equity_alt",
    # AdditionalPaidinCapital queda FUERA del mapa a proposito: el mismo puente
    # FY2024 (issue 35) encontro incorrecto su mapeo ingenuo a una linea de
    # capital. El desglose del capital se captura del filing, no de aqui.
    "CashAndCashEquivalents": "bs_cash",
    "Inventories": "bs_inv",
    "RetainedEarnings": "bs_re",
    "CashFlowsFromUsedInOperatingActivities": "cf_cfo",
    "CashFlowsFromUsedInInvestingActivities": "cf_cfi",
    "CashFlowsFromUsedInFinancingActivities": "cf_cff",
    # Obligatorio: sin el efecto cambiario el roll de caja no cierra (CFO +
    # CFI + CFF != cambio neto) en una emisora que reporta en TWD (issue 35).
    "EffectOfExchangeRateChangesOnCashAndCashEquivalents": "cf_fx_effect",
    "IncreaseDecreaseInCashAndCashEquivalents": "cf_net_change",
    "AdjustmentsForDepreciationAndAmortisationExpense": "cf_da",
    "AdjustmentsForSharebasedPayments": "cf_sbc",
    "IncomeTaxesPaidRefundClassifiedAsOperatingActivities": "cf_taxes_paid",
    "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities": "cf_capex",
    "DividendsPaidClassifiedAsFinancingActivities": "cf_dividends",
    "PaymentsToAcquireOrRedeemEntitysShares": "cf_buybacks",
    "ProceedsFromIssuingShares": "cf_stock_issued",
    "ProceedsFromBorrowingsClassifiedAsFinancingActivities": "cf_debt_issued",
    "RepaymentsOfBorrowingsClassifiedAsFinancingActivities": "cf_debt_repaid",
    # --- Conceptos que TSMC SI reporta (lista --unmapped de la corrida en vivo
    # del 2026-11-25; el mapa inicial, armado sin payload real, dejaba la
    # captura sin D&A, sin impuestos pagados y sin deuda). Regla: dos conceptos
    # que se SUMAN (depreciacion y amortizacion; bonos y prestamos) llevan
    # canon propio y statement-mapper los agrega con gate; solo comparten canon
    # los que son alternativas de una misma linea (impuestos pagados).
    "SalesAndMarketingExpense": "is_sm",
    "GeneralAndAdministrativeExpense": "is_ga",
    "OperatingExpenseExcludingCostOfSales": "is_opex_total",
    "FinanceCosts": "is_finance_costs",
    "ShareOfProfitLossOfAssociatesAccountedForUsingEquityMethod": "is_equity_method_income",
    "ProfitLossAttributableToNoncontrollingInterests": "is_ni_nci",
    "BasicEarningsLossPerShare": "is_eps_basic",
    "PropertyPlantAndEquipment": "bs_ppe",
    "IntangibleAssetsAndGoodwill": "bs_intangibles_goodwill",
    "RightofuseAssets": "bs_rou_assets",
    "CurrentTradeReceivables": "bs_ar",
    "TradeAndOtherCurrentPayablesToTradeSuppliers": "bs_ap",
    "NoncontrollingInterests": "bs_nci",
    "NoncurrentPortionOfNoncurrentBondsIssued": "bs_bonds_lt",
    "CurrentBondsIssuedAndCurrentPortionOfNoncurrentBondsIssued": "bs_bonds_st",
    "LongtermBorrowings": "bs_borrowings_lt",
    "CurrentPortionOfLongtermBorrowings": "bs_borrowings_lt_current",
    "NoncurrentLeaseLiabilities": "bs_lease_lt",
    "DepreciationExpense": "cf_depreciation",
    "AmortisationExpense": "cf_amortisation",
    "IncomeTaxesPaidClassifiedAsOperatingActivities": "cf_taxes_paid",
    "PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities": "cf_capex_intangibles",
    "ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities": "cf_ppe_disposals",
    "ProceedsFromIssueOfBondsNotesAndDebentures": "cf_bonds_issued",
    "RepaymentsOfBondsNotesAndDebentures": "cf_bonds_repaid",
    "ProceedsFromNoncurrentBorrowings": "cf_borrowings_lt_issued",
    "InterestPaidClassifiedAsFinancingActivities": "cf_interest_paid",
    "PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities": "cf_lease_payments",
}

# Mapas por taxonomia, en orden de preferencia: dentro de un mismo accession
# el concepto que aparece antes gana (vigente sobre legado).
TAXONOMY_MAPS: dict[str, tuple[dict[str, str], ...]] = {
    "us-gaap": (CONCEPT_MAP, LEGACY_CONCEPT_MAP),
    "ifrs-full": (IFRS_CONCEPT_MAP,),
}
DEFAULT_TAXONOMY = "us-gaap"
DEFAULT_CURRENCY = "USD"
AUTO = "auto"

# Salidas de efectivo que XBRL reporta en POSITIVO (la definicion del concepto
# es un pago). Conjunto explicito, no heuristica por nombre: cada alta se
# decide leyendo la definicion del concepto. El CSV NO cambia el signo; marca
# la fila con ``sign = salida(+)`` y statement-mapper aplica la convencion.
SIGN_OUTFLOW_POSITIVE = "salida(+)"
OUTFLOW_POSITIVE_CONCEPTS = frozenset({
    # us-gaap
    "PaymentsToAcquirePropertyPlantAndEquipment",
    "PaymentsForRepurchaseOfCommonStock",
    "PaymentsOfDividends",
    "PaymentsOfDividendsCommonStock",
    "PaymentsToAcquireAvailableForSaleSecuritiesDebt",
    "PaymentsToAcquireAvailableForSaleSecurities",
    "PaymentsToAcquireBusinessesNetOfCashAcquired",
    "PaymentsForProceedsFromOtherInvestingActivities",   # neto: positivo = pago
    "RepaymentsOfLongTermDebt",
    "PaymentsRelatedToTaxWithholdingForShareBasedCompensation",
    # ifrs-full
    "IncomeTaxesPaidRefundClassifiedAsOperatingActivities",  # neto: positivo = pago
    "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
    "DividendsPaidClassifiedAsFinancingActivities",
    "PaymentsToAcquireOrRedeemEntitysShares",
    "RepaymentsOfBorrowingsClassifiedAsFinancingActivities",
    "IncomeTaxesPaidClassifiedAsOperatingActivities",
    "PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities",
    "RepaymentsOfBondsNotesAndDebentures",
    "InterestPaidClassifiedAsFinancingActivities",
    "PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities",
})

# Seleccion de vintage: el accession primario de un (estado, periodo) es el mas
# reciente que reporta al menos este porcentaje de las lineas del accession
# mejor cubierto. Sin el umbral, un filing posterior que solo repite una o dos
# cifras de un periodo viejo (p. ej. el capital en el estado de cambios)
# ganaria y romperia la consistencia entre lineas.
VINTAGE_COVERAGE_MIN_PCT = 80
VINTAGE_MIX_NOTE = "vintage distinto"
_VINTAGE_MIX_KEY = "_vintage_mix"

# Flujos que se desacumulan y cuyo 4Q se deriva FY - acumulado a 3Q. La UPA no
# es aditiva entre trimestres (cambia el conteo de acciones): nunca se deriva.
FLOW_PREFIXES = ("is_", "cf_")
NON_ADDITIVE = frozenset({"is_eps_diluted", "is_eps_basic"})
YTD_UNRESOLVED_TAG = "acumulado YTD sin desacumular"

FIELDNAMES = ("canon", "period", "value", "months", "start", "end",
              "form", "filed", "concept", "unit", "source", "tag",
              # Columnas nuevas SIEMPRE al final: los lectores existentes
              # indexan las doce anteriores.
              "accession", "sign")
# Diagnostico --unmapped: conceptos con datos que el mapa no cubre.
UNMAPPED_FIELDNAMES = ("taxonomy", "concept", "label", "units", "facts",
                       "first_end", "last_end", "last_filed", "last_form")
UNMAPPED_TOP = 10


BLOCKED_MSG = """[x] Sin acceso de red a SEC ({err}).
    El entorno bloquea data.sec.gov (proxy con allowlist - tipico en Claude
    Cowork). Opciones: 1) permitir www.sec.gov y data.sec.gov en el allowlist
    del entorno y reintentar; 2) correr este comando en una maquina con red y
    copiar el CSV a model/inputs/; 3) captura manual via statement-mapper.
    Nota: SEC tambien responde 403 si el User-Agent no trae contacto valido
    o si se excede la tasa de solicitudes."""

NO_FACTS_MSG = """[x] SEC no tiene companyfacts XBRL para este CIK (HTTP 404).
    Tipico en CIKs antiguos o en emisoras que nunca presentaron estados
    etiquetados en XBRL. No depende del marco contable: las emisoras IFRS
    (20-F/40-F) si tienen companyfacts, en la taxonomia ifrs-full. No es un
    problema de red: captura manual via statement-mapper desde los filings."""

NO_MAPPED_MSG = """[aviso] sin conceptos mapeados; taxonomias en el payload: {present}.
    Este tool lee {supported}. El CSV sale sin filas: la emisora no
    etiqueta sus estados en esas taxonomias; captura manual via
    statement-mapper desde los filings."""

RETRY_CODES = (429, 500, 502, 503, 504)
RETRIES = 3


def _get_json(url: str, user_agent: str) -> dict:
    """GET JSON con clasificacion de errores.

    HTTPError se atrapa ANTES que URLError porque es su subclase: sin ese
    orden, un 404 o un 429 se reportaban como "red bloqueada".
    """
    req = urllib.request.Request(url, headers={"User-Agent": user_agent})
    for attempt in range(RETRIES + 1):
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code in (403, 407):
                raise SystemExit(BLOCKED_MSG.format(err=f"HTTP {exc.code}"))
            if exc.code == 404 and "companyfacts" in url:
                raise SystemExit(NO_FACTS_MSG)
            if exc.code in RETRY_CODES and attempt < RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise SystemExit(f"[x] SEC HTTP {exc.code} en {url}")
        except urllib.error.URLError as exc:
            raise SystemExit(BLOCKED_MSG.format(err=f"URLError {exc.reason}"))
        except (TimeoutError, ConnectionError) as exc:
            if attempt < RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise SystemExit(f"[x] SEC sin respuesta ({type(exc).__name__}) en {url}")
    raise SystemExit(f"[x] SEC: reintentos agotados en {url}")  # pragma: no cover


def resolve_cik(ticker: str, user_agent: str) -> str:
    data = _get_json(TICKERS_URL, user_agent)
    wanted = ticker.upper()
    for entry in data.values():
        if str(entry.get("ticker", "")).upper() == wanted:
            return f"{int(entry['cik_str']):010d}"
    raise SystemExit(f"[x] ticker {wanted} no encontrado")


def _parse_date(value: Optional[str]) -> Optional[date]:
    try:
        return date.fromisoformat(value) if value else None
    except (TypeError, ValueError):
        return None


def duration_months(item: dict) -> Optional[int]:
    """Meses cubiertos por el hecho (None si es un saldo puntual)."""
    start = _parse_date(item.get("start"))
    end = _parse_date(item.get("end"))
    if not start or not end:
        return None
    return round((end - start).days / 30.44)


def nominal_month(end: date) -> tuple[int, int]:
    """(anio, mes) NOMINAL de un cierre, tolerante a anios de 52/53 semanas.

    Las emisoras con anio de 52/53 semanas cierran en el sabado o domingo mas
    cercano a fin de mes, a veces unos dias DENTRO del mes siguiente: Costco
    cerro FY2023 el 2023-09-03 y Starbucks el 2023-10-01. Ese cierre pertenece
    al mes anterior. Regla: dia <= 15 -> mes previo; dia > 15 -> ese mes.
    """
    if end.day <= 15:
        return (end.year - 1, 12) if end.month == 1 else (end.year, end.month - 1)
    return end.year, end.month


def period_label(item: dict, fye_month: int) -> Optional[str]:
    """Periodo FISCAL derivado de la FECHA DE CIERRE, no de fp/fy.

    Critico: en companyfacts, ``fy``/``fp`` describen el filing donde aparece
    el hecho, NO el periodo que mide — un dato con end 2024-12-28 aparece
    etiquetado fy=2026 y produce series corridas un anio o mas. El unico
    ancla confiable es ``end`` contra el cierre fiscal de la emisora.

    Con cierre fiscal en septiembre: end 2024-12-28 -> 1Q2025 (FY2025 corre
    de oct-2024 a sep-2025); end 2025-09-27 -> 4Q2025. El mes se toma NOMINAL
    (ver ``nominal_month``): con cierre en septiembre, end 2023-10-01 ->
    FY2023 y end 2023-01-01 -> 1Q2023. FY = anio calendario del cierre.
    """
    end = _parse_date(item.get("end"))
    if not end:
        return None
    year, month = nominal_month(end)
    fy = year + (1 if month > fye_month else 0)
    offset = (month - fye_month) % 12           # 0 = cierre de anio fiscal
    months = duration_months(item)
    if months is not None and months >= 11:
        return f"FY{fy}"
    quarter = 4 if offset == 0 else (offset + 2) // 3
    if quarter not in (1, 2, 3, 4):
        return None
    return f"{quarter}Q{fy}"


def detect_fye_month(facts: dict) -> int:
    """Mes de cierre fiscal, deducido de los hechos anuales (duracion ~12m)."""
    counts: dict[int, int] = {}
    for node in facts.values():
        for items in node.get("units", {}).values():
            for item in items:
                months = duration_months(item)
                end = _parse_date(item.get("end"))
                if months and months >= 11 and end:
                    month = nominal_month(end)[1]
                    counts[month] = counts.get(month, 0) + 1
    if not counts:
        return 12
    return max(counts.items(), key=lambda kv: kv[1])[0]


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def is_flow(canon: str) -> bool:
    """Linea de flujo aditiva (IS o CF), excluida la UPA."""
    return canon.startswith(FLOW_PREFIXES) and canon not in NON_ADDITIVE


def deaccumulate(rows: dict[tuple[str, str], dict]) -> list[str]:
    """Convierte flujos YTD a TRIMESTRALES y deriva el 4Q (in place).

    En los 10-Q el flujo de efectivo viene acumulado del anio fiscal: 2Q cubre
    6 meses, 3Q nueve, FY doce. El trimestre real es el acumulado menos el
    acumulado previo del MISMO anio fiscal. El 4Q no se reporta aislado en el
    10-K: se deriva FY - acumulado a 3Q, para IS y CF por igual (salvo UPA).

    Un acumulado sin el acumulado previo NO se puede desacumular: la fila se
    conserva con tag ``YTD_UNRESOLVED_TAG`` y una nota ``[aviso]`` para que
    nunca entre al modelo como si fuera un trimestre.
    """
    notes: list[str] = []
    by_canon_fy: dict[tuple[str, str], dict[str, dict]] = {}
    for (canon, period), row in rows.items():
        if not is_flow(canon) or row.get("months") is None:
            continue
        if period.startswith("FY"):
            fy, q = period[2:], "FY"
        else:
            q, fy = period[0], period[2:]
        by_canon_fy.setdefault((canon, fy), {})[q] = row
    for (canon, fy), qs in by_canon_fy.items():
        # cum[q]: acumulado del anio fiscal hasta el trimestre q (None = no se sabe)
        cum: dict[int, Optional[float]] = {0: 0.0}
        for q in (1, 2, 3):
            row = qs.get(str(q))
            if row is None or not _is_number(row.get("value")):
                cum[q] = None
                continue
            months = row.get("months") or 0
            if months <= 4:                           # trimestre discreto
                prev = cum[q - 1]
                cum[q] = None if prev is None else prev + row["value"]
            elif abs(months - 3 * q) <= 1:            # acumulado YTD de q trimestres
                ytd = row["value"]
                prev = cum[q - 1]
                if prev is None:
                    row["tag"] = f"{YTD_UNRESOLVED_TAG} ({months}m, falta {q - 1}Q)"
                    notes.append(f"[aviso] {canon} {q}Q{fy}: acumulado {months}m "
                                 f"sin {q - 1}Q -> NO es trimestral")
                else:
                    row["value"] = ytd - prev
                    row["months"] = 3
                    row["tag"] = "observado (desacumulado YTD)"
                    notes.append(f"{canon} {q}Q{fy}")
                cum[q] = ytd
            else:
                row["tag"] = f"{YTD_UNRESOLVED_TAG} (duracion {months}m inesperada)"
                notes.append(f"[aviso] {canon} {q}Q{fy}: duracion {months}m inesperada")
                cum[q] = None
        # 4Q = FY - acumulado a 3Q (solo si no hay un 4Q discreto reportado)
        fy_row = qs.get("FY")
        q4 = qs.get("4")
        cum3 = cum.get(3)
        if (fy_row and _is_number(fy_row.get("value")) and cum3 is not None
                and (q4 is None or (q4.get("months") or 0) > 4)):
            q3_end = _parse_date(qs["3"].get("end"))
            start = (q3_end + timedelta(days=1)).isoformat() if q3_end else ""
            rows[(canon, f"4Q{fy}")] = {
                **fy_row,
                "canon": canon, "period": f"4Q{fy}", "start": start,
                "value": fy_row["value"] - cum3, "months": 3,
                "tag": "derivado (FY - acumulado 3Q)",
            }
            notes.append(f"{canon} 4Q{fy} (derivado)")
    return notes


@dataclass(frozen=True)
class Fact:
    """Un hecho de companyfacts ya mapeado a (canon, periodo), aun sin elegir."""

    canon: str
    period: str
    value: object
    months: Optional[int]
    start: str
    end: str
    form: str
    filed: str
    accession: str
    concept: str
    unit: str
    rank: int  # posicion del concepto en el mapa: menor = preferido

    @property
    def vintage(self) -> str:
        """Identidad del filing que reporta el hecho.

        Es el accession; un payload sin ``accn`` (sintetico o muy viejo) cae a
        la fecha de presentacion, que es lo que antes desempataba.
        """
        return self.accession or f"filed:{self.filed}"


def taxonomy_concepts(taxonomy: str) -> list[tuple[str, str]]:
    """Pares (concepto, canon) de una taxonomia, en orden de preferencia."""
    return [pair for mapping in TAXONOMY_MAPS[taxonomy] for pair in mapping.items()]


def _has_data(node: object) -> bool:
    return isinstance(node, dict) and any(node.get("units", {}).values())


def choose_taxonomy(all_facts: dict, requested: str = AUTO) -> tuple[str, dict[str, int]]:
    """Taxonomia a leer y conceptos mapeados con datos por taxonomia soportada.

    ``auto`` toma la que trae mas conceptos mapeados con datos. En empate (o
    payload vacio) gana us-gaap: es el comportamiento historico del tool.
    """
    counts = {name: sum(1 for concept, _ in taxonomy_concepts(name)
                        if _has_data(all_facts.get(name, {}).get(concept)))
              for name in TAXONOMY_MAPS}
    if requested != AUTO:
        if requested not in TAXONOMY_MAPS:
            raise SystemExit(f"[x] taxonomia no soportada: {requested!a} "
                             f"(opciones: {AUTO}, {', '.join(TAXONOMY_MAPS)})")
        return requested, counts
    # max() devuelve el primero de los empatados: us-gaap va primero en el mapa.
    return max(TAXONOMY_MAPS, key=lambda name: counts[name]), counts


_CURRENCY_RE = re.compile(r"[A-Z]{3}")


def normalize_currency(requested: Optional[str]) -> str:
    """``auto`` o un codigo ISO 4217 en mayusculas; cualquier otra cosa aborta."""
    if not requested or requested.lower() == AUTO:
        return AUTO
    code = requested.upper()
    if not _CURRENCY_RE.fullmatch(code):
        raise SystemExit("[x] --currency espera un codigo ISO 4217 de 3 letras "
                         f"(USD, TWD, EUR); recibio {requested!a}")
    return code


def accepted_units(currency: str) -> tuple[str, str]:
    """Unidades XBRL que se aceptan para una moneda: importes y por accion."""
    return (currency, f"{currency}/shares")


def currency_counts(facts: dict, concepts: Sequence[tuple[str, str]]) -> dict[str, int]:
    """Hechos mapeados por moneda, contando solo unidades monetarias puras."""
    counts: dict[str, int] = {}
    for concept, _ in concepts:
        for unit, items in (facts.get(concept) or {}).get("units", {}).items():
            if items and _CURRENCY_RE.fullmatch(unit):
                counts[unit] = counts.get(unit, 0) + len(items)
    return counts


def choose_currency(counts: dict[str, int], requested: Optional[str] = AUTO) -> str:
    """Moneda de reporte: la pedida o, en ``auto``, la de mas hechos mapeados.

    Empate -> orden alfabetico (determinista); sin hechos -> USD.
    """
    code = normalize_currency(requested)
    if code != AUTO:
        return code
    if not counts:
        return DEFAULT_CURRENCY
    return min(counts, key=lambda cur: (-counts[cur], cur))


def collect_candidates(facts: dict, concepts: Sequence[tuple[str, str]],
                       currency: str, fye_month: int) -> list[Fact]:
    """TODOS los hechos mapeados de la moneda pedida, sin deduplicar.

    La eleccion entre hechos del mismo (canon, periodo) es de
    ``select_by_vintage``: aqui no se descarta nada.
    """
    units = accepted_units(currency)
    out: list[Fact] = []
    for rank, (concept, canon) in enumerate(concepts):
        for unit, items in (facts.get(concept) or {}).get("units", {}).items():
            if unit not in units:
                continue
            for item in items:
                label = period_label(item, fye_month)
                if not label:
                    continue
                out.append(Fact(
                    canon=canon, period=label, value=item.get("val"),
                    months=duration_months(item),
                    start=item.get("start", ""), end=item.get("end", ""),
                    form=item.get("form", ""), filed=item.get("filed", ""),
                    accession=item.get("accn", ""), concept=concept, unit=unit,
                    rank=rank))
    return out


def statement_of(canon: str) -> str:
    """Estado al que pertenece una linea canon: ``is``, ``bs`` o ``cf``."""
    return canon.split("_", 1)[0]


def _preference(fact: Fact) -> tuple[int, int]:
    """Orden DENTRO de un accession (menor gana).

    Primero la duracion mas corta: el hecho del trimestre, no el acumulado
    que cierra el mismo dia. A igual duracion, el concepto vigente sobre el
    legado (posicion en el mapa).
    """
    return (fact.months if fact.months is not None else 0, fact.rank)


def select_by_vintage(candidates: Iterable[Fact]) -> tuple[
        dict[tuple[str, str], Fact], dict[tuple[str, str], str]]:
    """Un accession primario por (estado, periodo); de el salen sus lineas.

    El dedup por concepto ("gana el filed mas reciente") dejaba que cada linea
    de un balance saliera de un filing distinto (original vs re-presentado) y
    el balance no cuadraba sin que ninguna cifra estuviera mal. Aqui:

    - Se agrupa por estado (prefijo ``is``/``bs``/``cf``) y periodo, no por
      periodo completo: un 10-K posterior repite tres anios de resultados y
      flujos pero solo dos balances, y un 10-Q posterior repite los flujos del
      trimestre del anio previo pero no su balance.
    - Primario = el accession presentado mas recientemente (empate: accession
      descendente) que reporta al menos ``VINTAGE_COVERAGE_MIN_PCT`` % de las
      lineas del accession mejor cubierto del grupo.
    - Una linea que el primario no reporta se toma del accession mas reciente
      que si la tenga y se devuelve en ``fill_ins`` para que quede marcada.

    Devuelve ``(chosen, fill_ins)``: el hecho elegido por (canon, periodo) y,
    solo para las lineas completadas de otro accession, el vintage primario.
    """
    groups: dict[tuple[str, str], dict[str, dict[str, Fact]]] = {}
    for fact in candidates:
        lines = groups.setdefault((statement_of(fact.canon), fact.period), {}) \
                      .setdefault(fact.vintage, {})
        prev = lines.get(fact.canon)
        if prev is None or _preference(fact) < _preference(prev):
            lines[fact.canon] = fact
    chosen: dict[tuple[str, str], Fact] = {}
    fill_ins: dict[tuple[str, str], str] = {}
    for (_, period), vintages in groups.items():
        newest_first = sorted(
            vintages, reverse=True,
            key=lambda v: (max(f.filed for f in vintages[v].values()), v))
        best = max(len(lines) for lines in vintages.values())
        primary = next(v for v in newest_first
                       if len(vintages[v]) * 100 >= VINTAGE_COVERAGE_MIN_PCT * best)
        for canon in sorted({c for lines in vintages.values() for c in lines}):
            source = primary if canon in vintages[primary] else next(
                v for v in newest_first if canon in vintages[v])
            chosen[(canon, period)] = vintages[source][canon]
            if source != primary:
                fill_ins[(canon, period)] = primary
    return chosen, fill_ins


def fact_to_row(fact: Fact, cik: str, taxonomy: str = DEFAULT_TAXONOMY) -> dict:
    """Fila del CSV para un hecho elegido (valor tal como lo reporta XBRL)."""
    concept_ref = fact.concept if taxonomy == DEFAULT_TAXONOMY else f"{taxonomy}:{fact.concept}"
    return {
        "canon": fact.canon, "period": fact.period,
        "value": fact.value, "months": fact.months,
        "start": fact.start, "end": fact.end,
        "form": fact.form, "filed": fact.filed,
        "concept": fact.concept, "unit": fact.unit,
        "source": f"XBRL companyfacts CIK{cik} {concept_ref}",
        "tag": "observado",
        "accession": fact.accession,
        "sign": SIGN_OUTFLOW_POSITIVE if fact.concept in OUTFLOW_POSITIVE_CONCEPTS else "",
    }


def build_rows(facts: dict, concepts: Sequence[tuple[str, str]], currency: str,
               fye_month: int, cik: str,
               taxonomy: str = DEFAULT_TAXONOMY) -> dict[tuple[str, str], dict]:
    """Hechos de una taxonomia -> filas elegidas por vintage (sin I/O).

    Las filas completadas de un accession distinto al primario llevan la
    clave interna ``_VINTAGE_MIX_KEY`` = (accession de la fila, primario);
    ``mark_vintage_mix`` la vuelve tag DESPUES de desacumular, para que el 4Q
    derivado de una fila completada herede la marca.
    """
    chosen, fill_ins = select_by_vintage(
        collect_candidates(facts, concepts, currency, fye_month))
    rows = {key: fact_to_row(fact, cik, taxonomy) for key, fact in chosen.items()}
    for key, primary in fill_ins.items():
        rows[key][_VINTAGE_MIX_KEY] = (chosen[key].vintage, primary)
    return rows


def _annotate(tag: str, note: str) -> str:
    """Agrega una nota al tag conservando lo que ya decia."""
    return f"{tag[:-1]}; {note})" if tag.endswith(")") else f"{tag} ({note})"


def mark_vintage_mix(rows: dict[tuple[str, str], dict]) -> list[str]:
    """Marca en ``tag`` las filas de otro vintage (in place) y las describe."""
    notes: list[str] = []
    for (canon, period), row in sorted(rows.items()):
        mix = row.get(_VINTAGE_MIX_KEY)
        if not mix:
            continue
        source, primary = mix
        row["tag"] = _annotate(str(row.get("tag", "")), f"{VINTAGE_MIX_NOTE}: {source}")
        notes.append(f"{canon} {period} <- {source} (primario {primary})")
    return notes


def _row_order(row: dict) -> tuple[str, str, int, str]:
    """Orden del CSV: canon, cierre y, en el mismo cierre, FY antes que 4Q."""
    return (row["canon"], row["end"], -(row.get("months") or 0), row["period"])


def _counts_text(counts: dict[str, int]) -> str:
    return ", ".join(f"{name} {n}" for name, n in counts.items()) or "-"


@dataclass(frozen=True)
class UnmappedConcept:
    """Concepto que la emisora reporta y el mapa de su taxonomia no cubre."""

    concept: str
    label: str
    units: str
    facts: int
    first_end: str
    last_end: str
    last_filed: str
    last_form: str


def unmapped_concepts(facts: dict,
                      concepts: Sequence[tuple[str, str]]) -> list[UnmappedConcept]:
    """Conceptos CON datos que no estan en el mapa, lo vigente primero.

    Es el diagnostico de una captura incompleta: cada emisora elige sus
    conceptos (TSMC no etiqueta D&A con el concepto que el mapa esperaba) y un
    concepto ausente del mapa no emite filas ni avisa. Orden: ultimo cierre
    descendente, luego numero de hechos; una linea descontinuada queda al final.
    """
    mapped = {concept for concept, _ in concepts}
    found: list[UnmappedConcept] = []
    for concept in sorted(facts):
        node = facts[concept]
        if concept in mapped or not _has_data(node):
            continue
        items = [(unit, item) for unit, unit_items in node["units"].items()
                 for item in unit_items if item.get("end")]
        if not items:
            continue
        last = max((item for _, item in items),
                   key=lambda item: (item["end"], item.get("filed", "")))
        found.append(UnmappedConcept(
            concept=concept, label=str(node.get("label") or ""),
            units=",".join(sorted({unit for unit, _ in items})), facts=len(items),
            first_end=min(item["end"] for _, item in items), last_end=last["end"],
            last_filed=last.get("filed", ""), last_form=last.get("form", "")))
    # sorted() es estable: el orden alfabetico previo desempata.
    return sorted(found, key=lambda u: (u.last_end, u.facts), reverse=True)


def latest_annual_end(facts: dict) -> tuple[str, list[str]]:
    """Ultimo cierre de duracion anual de una taxonomia y quienes lo reportan."""
    latest, reporters = "", []
    for concept in sorted(facts):
        node = facts[concept]
        if not isinstance(node, dict):
            continue
        for items in node.get("units", {}).values():
            for item in items:
                end = item.get("end", "")
                if (duration_months(item) or 0) < 11 or end < latest:
                    continue
                if end > latest:
                    latest, reporters = end, []
                if concept not in reporters:
                    reporters.append(concept)
    return latest, reporters


def period_gaps(all_facts: dict, taxonomy_used: str,
                rows: dict[tuple[str, str], dict]) -> list[str]:
    """Avisos si el payload trae anios posteriores al ultimo anio capturado.

    Un CSV que termina un anio antes que los filings de la emisora es una
    captura incompleta que parece completa (TSM: FY2025 ausente). Se revisa la
    taxonomia leida y las demas del payload (``dei`` no trae estados).
    """
    captured = max((str(r.get("end", "")) for r in rows.values()
                    if (r.get("months") or 0) >= 11), default="")
    mapped = {concept for concept, _ in taxonomy_concepts(taxonomy_used)}
    notes: list[str] = []
    for name in sorted(all_facts):
        if name == "dei" or not isinstance(all_facts[name], dict):
            continue
        latest, reporters = latest_annual_end(all_facts[name])
        if not latest or latest <= captured:
            continue
        if name != taxonomy_used:
            notes.append(f"{name}: hasta {latest} (otra taxonomia; revisar --taxonomy)")
            continue
        outside = [c for c in reporters if c not in mapped]
        if outside:
            notes.append(f"{name}: hasta {latest} en {len(outside)} conceptos sin mapear "
                         f"({', '.join(outside[:3])}{' ...' if len(outside) > 3 else ''})")
        else:
            notes.append(f"{name}: hasta {latest} en conceptos mapeados pero en otra "
                         "unidad o moneda (revisar --currency)")
    if not notes:
        return []
    return [f"[aviso] el payload trae cierres anuales posteriores al ultimo capturado "
            f"({captured or 'ninguno'}):",
            *(f"    {note}" for note in notes),
            "    El CSV NO cubre ese periodo: corre con --unmapped para ver que conceptos",
            "    faltan en el mapa, o capturalo desde el filing."]


def write_unmapped(path: Path, taxonomy: str, items: Sequence[UnmappedConcept]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(UNMAPPED_FIELDNAMES)
        for u in items:
            writer.writerow((taxonomy, u.concept, u.label, u.units, u.facts,
                             u.first_end, u.last_end, u.last_filed, u.last_form))


def fetch(ticker: str, dest: Path, user_agent: str,
          taxonomy: str = AUTO, currency: str = AUTO, unmapped: bool = False) -> Path:
    cik = resolve_cik(ticker, user_agent)
    data = _get_json(FACTS_URL.format(cik=cik), user_agent)
    all_facts = data.get("facts", {})
    taxonomy_used, tax_counts = choose_taxonomy(all_facts, taxonomy)
    how = "auto" if taxonomy == AUTO else "fijada con --taxonomy"
    print(f"[ok] taxonomia: {taxonomy_used} ({how}; conceptos mapeados con datos: "
          f"{_counts_text(tax_counts)})")
    if not any(tax_counts.values()):
        print(NO_MAPPED_MSG.format(supported=", ".join(TAXONOMY_MAPS),
                                   present=", ".join(sorted(all_facts)) or "ninguna"))
    elif tax_counts[taxonomy_used] < max(tax_counts.values()):
        print(f"[aviso] {taxonomy_used} trae menos conceptos mapeados que otra "
              f"taxonomia del payload ({_counts_text(tax_counts)}); revisar --taxonomy")
    facts = all_facts.get(taxonomy_used, {})
    concepts = taxonomy_concepts(taxonomy_used)
    cur_counts = currency_counts(facts, concepts)
    currency_used = choose_currency(cur_counts, currency)
    how = "auto" if normalize_currency(currency) == AUTO else "fijada con --currency"
    print(f"[ok] moneda: {currency_used} ({how}); unidades aceptadas: "
          + ", ".join(accepted_units(currency_used)))
    by_volume = dict(sorted(cur_counts.items(), key=lambda kv: (-kv[1], kv[0])))
    if len(cur_counts) > 1:
        print(f"[aviso] la emisora reporta en mas de una moneda (hechos mapeados: "
              f"{_counts_text(by_volume)}); se toma {currency_used} "
              "(--currency <ISO> para otra)")
    if cur_counts and currency_used not in cur_counts:
        print(f"[aviso] sin hechos mapeados en {currency_used}; con datos: "
              f"{_counts_text(by_volume)}")
    fye_month = detect_fye_month(facts)
    print(f"[ok] cierre fiscal detectado: mes {fye_month}")
    rows = build_rows(facts, concepts, currency_used, fye_month, cik, taxonomy_used)
    deacc = deaccumulate(rows)
    warnings = [n for n in deacc if n.startswith("[aviso]")]
    print(f"[ok] flujos desacumulados/derivados: {len(deacc) - len(warnings)}")
    if warnings:
        print(f"[aviso] {len(warnings)} filas acumuladas YTD que NO son trimestrales "
              f"(tag '{YTD_UNRESOLVED_TAG}'):")
        for note in warnings[:5]:
            print(f"    {note[len('[aviso] '):]}")
    mixed = mark_vintage_mix(rows)
    print("[ok] vintage: un accession primario por estado y periodo "
          "(columna accession)")
    if mixed:
        plural = "s" if len(mixed) != 1 else ""
        print(f"[aviso] {len(mixed)} fila{plural} de un accession distinto al primario "
              f"de su estado y periodo (tag '{VINTAGE_MIX_NOTE}: <accession>'):")
        for note in mixed[:5]:
            print(f"    {note}")
    dest.mkdir(parents=True, exist_ok=True)
    out = dest / f"xbrl_facts_{ticker.upper()}.csv"
    ordered = sorted(rows.values(), key=_row_order)
    with out.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(ordered)
    quarters = sorted({r["period"] for r in ordered if "Q" in r["period"]})
    years = sorted({r["period"] for r in ordered if r["period"].startswith("FY")})
    print(f"[ok] {ticker.upper()}: {len(ordered)} observaciones -> {out.name}")
    print(f"[ok] trimestres cubiertos: {len(quarters)} ({quarters[0] if quarters else '-'}"
          f" .. {quarters[-1] if quarters else '-'})")
    print(f"[ok] anios cubiertos: {len(years)}")
    canons_with_data = {r["canon"] for r in ordered}
    missing = sorted(c for c, canon in TAXONOMY_MAPS[taxonomy_used][0].items()
                     if not _has_data(facts.get(c)) and canon not in canons_with_data)
    if missing:
        print(f"[!] conceptos sin datos en esta emisora ({len(missing)}): "
              + ", ".join(missing[:6]) + (" ..." if len(missing) > 6 else ""))
    # companyfacts puede ir atrasado respecto a los filings (TSM, 2026-11-25:
    # ultimo cierre 2024-12-31 con el 20-F de FY2025 ya presentado): el
    # analista debe ver hasta donde llega el payload, no deducirlo del CSV.
    print(f"[i] ultimo cierre anual en companyfacts ({taxonomy_used}): "
          f"{latest_annual_end(facts)[0] or 'ninguno'} - un anio posterior ya "
          "presentado se captura desde el filing")
    for line in period_gaps(all_facts, taxonomy_used, rows):
        print(line)
    outside = unmapped_concepts(facts, concepts)
    print(f"[i] {len(outside)} conceptos con datos sin mapear en {taxonomy_used}"
          + ("" if unmapped else " (--unmapped escribe la lista)"))
    if unmapped:
        unmapped_path = dest / f"xbrl_unmapped_{ticker.upper()}.csv"
        write_unmapped(unmapped_path, taxonomy_used, outside)
        print(f"[ok] lista completa -> {unmapped_path.name}; los mas recientes:")
        for u in outside[:UNMAPPED_TOP]:
            print(f"    {u.concept} ({u.facts} hechos, {u.first_end} .. {u.last_end}, {u.units})")
        print("    Un concepto que deba entrar al modelo se agrega al mapa de su")
        print("    taxonomia en tools/xbrl_fetch.py con su linea canon; no se captura a mano.")
    outflows = sum(1 for r in ordered if r.get("sign") == SIGN_OUTFLOW_POSITIVE)
    print(f"[i] signos: {outflows} filas con sign = {SIGN_OUTFLOW_POSITIVE} - pagos que XBRL reporta")
    print("    en POSITIVO (capex, dividendos, recompras, impuestos pagados,")
    print("    amortizacion de deuda). El valor NO se invierte aqui: statement-mapper")
    print("    aplica el signo del canonico y lo puentea contra el filing.")
    print("[i] siguiente paso: statement-mapper convierte este CSV largo en los")
    print("    canonical_*.csv (mapeo canon con gate del analista). El 4Q de")
    print("    IS y CF ya viene derivado (FY - acumulado 3Q); la UPA de 4Q no.")
    return out


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="SEC XBRL companyfacts -> CSV largo")
    parser.add_argument("ticker")
    parser.add_argument("--dest", default=".", help="carpeta destino (model/inputs)")
    parser.add_argument("--ua", default=None, help="User-Agent 'Nombre correo' (o env SEC_EDGAR_UA)")
    parser.add_argument("--taxonomy", default=AUTO, choices=(AUTO, *TAXONOMY_MAPS),
                        help="taxonomia XBRL a leer (default auto: la que trae mas "
                             "conceptos mapeados con datos)")
    parser.add_argument("--currency", default=AUTO, metavar="ISO",
                        help="moneda de reporte, codigo ISO 4217 como USD o TWD "
                             "(default auto: la de mas hechos mapeados)")
    parser.add_argument("--unmapped", action="store_true",
                        help="escribir xbrl_unmapped_<TICKER>.csv: conceptos con datos que "
                             "el mapa no cubre (diagnostico de una captura incompleta)")
    args = parser.parse_args(argv[1:])
    currency = normalize_currency(args.currency)      # falla antes de tocar la red
    ua = args.ua or os.environ.get("SEC_EDGAR_UA", "")
    if not ua or "@" not in ua:
        raise SystemExit("[x] SEC exige User-Agent con contacto: --ua o env SEC_EDGAR_UA")
    fetch(args.ticker, Path(args.dest), ua, taxonomy=args.taxonomy, currency=currency,
          unmapped=args.unmapped)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
