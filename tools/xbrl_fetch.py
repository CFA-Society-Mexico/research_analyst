"""SEC XBRL companyfacts -> serie historica completa (anual + trimestral).

Resuelve la captura masiva de trimestres sin parsear HTML: la API publica
``data.sec.gov/api/xbrl/companyfacts`` trae TODOS los valores reportados de la
emisora (10-K y 10-Q, con periodo fiscal fp/fy). Este tool baja los conceptos
US GAAP mapeados a lineas canon y escribe un CSV largo que statement-mapper
convierte (con gate) en los canonical_*.csv de model/inputs/.

Uso:
    python tools/xbrl_fetch.py AAPL --dest <raiz>/AAPL/model/inputs \
        --ua "Nombre correo@dominio.com"

User-Agent obligatorio (--ua o env SEC_EDGAR_UA), igual que sec_fetch.
Consola ASCII-only.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.request
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

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
# de valores sin el sufijo "Debt" hasta FY2018. Se procesan DESPUES del mapa
# principal: en un empate de fecha de filing gana el concepto vigente.
LEGACY_CONCEPT_MAP = {
    "SalesRevenueNet": "is_ns_total",
    "PaymentsToAcquireAvailableForSaleSecurities": "cf_buy_securities",
    "ProceedsFromSaleOfAvailableForSaleSecurities": "cf_sell_securities",
}

# Flujos que se desacumulan y cuyo 4Q se deriva FY - acumulado a 3Q. La UPA no
# es aditiva entre trimestres (cambia el conteo de acciones): nunca se deriva.
FLOW_PREFIXES = ("is_", "cf_")
NON_ADDITIVE = frozenset({"is_eps_diluted"})
YTD_UNRESOLVED_TAG = "acumulado YTD sin desacumular"


BLOCKED_MSG = """[x] Sin acceso de red a SEC ({err}).
    El entorno bloquea data.sec.gov (proxy con allowlist - tipico en Claude
    Cowork). Opciones: 1) permitir www.sec.gov y data.sec.gov en el allowlist
    del entorno y reintentar; 2) correr este comando en una maquina con red y
    copiar el CSV a model/inputs/; 3) captura manual via statement-mapper.
    Nota: SEC tambien responde 403 si el User-Agent no trae contacto valido
    o si se excede la tasa de solicitudes."""

NO_FACTS_MSG = """[x] SEC no tiene companyfacts XBRL para este CIK (HTTP 404).
    Tipico en emisoras sin XBRL US GAAP (p. ej. FPI que reporta 20-F en
    IFRS) o en CIKs antiguos. No es un problema de red: captura manual via
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


def fetch(ticker: str, dest: Path, user_agent: str) -> Path:
    cik = resolve_cik(ticker, user_agent)
    data = _get_json(FACTS_URL.format(cik=cik), user_agent)
    facts = data.get("facts", {}).get("us-gaap", {})
    fye_month = detect_fye_month(facts)
    print(f"[ok] cierre fiscal detectado: mes {fye_month}")
    rows: dict[tuple[str, str], dict] = {}
    found: set[str] = set()
    for concept, canon in [*CONCEPT_MAP.items(), *LEGACY_CONCEPT_MAP.items()]:
        node = facts.get(concept)
        if not node:
            continue
        found.add(concept)
        for unit, items in node.get("units", {}).items():
            if unit not in ("USD", "USD/shares"):
                continue
            for item in items:
                label = period_label(item, fye_month)
                if not label:
                    continue
                months = duration_months(item)
                key = (canon, label)
                prev = rows.get(key)
                # dedup: gana el filed mas reciente (re-presentaciones); entre
                # duraciones distintas del mismo periodo gana la mas corta
                # (el hecho del trimestre, no el acumulado que tambien cierra ahi)
                if prev is not None:
                    pm, cm = prev.get("months"), months
                    if pm is not None and cm is not None and cm != pm:
                        if pm < cm:
                            continue
                    elif item.get("filed", "") <= prev["filed"]:
                        continue
                rows[key] = {
                    "canon": canon, "period": label,
                    "value": item.get("val"), "months": months,
                    "start": item.get("start", ""), "end": item.get("end", ""),
                    "form": item.get("form", ""), "filed": item.get("filed", ""),
                    "concept": concept, "unit": unit,
                    "source": f"XBRL companyfacts CIK{cik} {concept}",
                    "tag": "observado",
                }
    deacc = deaccumulate(rows)
    warnings = [n for n in deacc if n.startswith("[aviso]")]
    print(f"[ok] flujos desacumulados/derivados: {len(deacc) - len(warnings)}")
    if warnings:
        print(f"[aviso] {len(warnings)} filas acumuladas YTD que NO son trimestrales "
              f"(tag '{YTD_UNRESOLVED_TAG}'):")
        for note in warnings[:5]:
            print(f"    {note[len('[aviso] '):]}")
    dest.mkdir(parents=True, exist_ok=True)
    out = dest / f"xbrl_facts_{ticker.upper()}.csv"
    fieldnames = ["canon", "period", "value", "months", "start", "end",
                  "form", "filed", "concept", "unit", "source", "tag"]
    ordered = sorted(rows.values(), key=lambda r: (r["canon"], r["end"]))
    with out.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(ordered)
    quarters = sorted({r["period"] for r in ordered if "Q" in r["period"]})
    years = sorted({r["period"] for r in ordered if r["period"].startswith("FY")})
    print(f"[ok] {ticker.upper()}: {len(ordered)} observaciones -> {out.name}")
    print(f"[ok] trimestres cubiertos: {len(quarters)} ({quarters[0] if quarters else '-'}"
          f" .. {quarters[-1] if quarters else '-'})")
    print(f"[ok] anios cubiertos: {len(years)}")
    canons_with_data = {r["canon"] for r in ordered}
    missing = sorted(c for c, canon in CONCEPT_MAP.items()
                     if c not in found and canon not in canons_with_data)
    if missing:
        print(f"[!] conceptos sin datos en esta emisora ({len(missing)}): "
              + ", ".join(missing[:6]) + (" ..." if len(missing) > 6 else ""))
    print("[i] siguiente paso: statement-mapper convierte este CSV largo en los")
    print("    canonical_*.csv (mapeo canon con gate del analista). El 4Q de")
    print("    IS y CF ya viene derivado (FY - acumulado 3Q); la UPA de 4Q no.")
    return out


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="SEC XBRL companyfacts -> CSV largo")
    parser.add_argument("ticker")
    parser.add_argument("--dest", default=".", help="carpeta destino (model/inputs)")
    parser.add_argument("--ua", default=None, help="User-Agent 'Nombre correo' (o env SEC_EDGAR_UA)")
    args = parser.parse_args(argv[1:])
    ua = args.ua or os.environ.get("SEC_EDGAR_UA", "")
    if not ua or "@" not in ua:
        raise SystemExit("[x] SEC exige User-Agent con contacto: --ua o env SEC_EDGAR_UA")
    fetch(args.ticker, Path(args.dest), ua)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
