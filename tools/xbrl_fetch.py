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
from datetime import date
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


def period_label(item: dict, fye_month: int) -> Optional[str]:
    """Periodo FISCAL derivado de la FECHA DE CIERRE, no de fp/fy.

    Critico: en companyfacts, ``fy``/``fp`` describen el filing donde aparece
    el hecho, NO el periodo que mide — un dato con end 2024-12-28 aparece
    etiquetado fy=2026 y produce series corridas un anio o mas. El unico
    ancla confiable es ``end`` contra el cierre fiscal de la emisora.

    Con cierre fiscal en septiembre: end 2024-12-28 -> 1Q2025 (FY2025 corre
    de oct-2024 a sep-2025); end 2025-09-27 -> 4Q2025.
    """
    end = _parse_date(item.get("end"))
    if not end:
        return None
    fy = end.year + (1 if end.month > fye_month else 0)
    offset = (end.month - fye_month) % 12       # 0 = cierre de anio fiscal
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
                    counts[end.month] = counts.get(end.month, 0) + 1
    if not counts:
        return 12
    return max(counts.items(), key=lambda kv: kv[1])[0]


def deaccumulate(rows: dict[tuple[str, str], dict]) -> list[str]:
    """Convierte flujos YTD a TRIMESTRALES (in place).

    En los 10-Q los flujos son acumulados del anio fiscal: 2Q cubre 6 meses,
    3Q nueve, FY doce. Tomarlos como trimestrales infla cada trimestre y
    rompe el roll de caja. El trimestre real se obtiene por diferencia con
    el acumulado previo del MISMO anio fiscal; el 4Q se deriva de FY menos
    los tres primeros.
    """
    notes: list[str] = []
    by_canon_fy: dict[tuple[str, str], dict[str, dict]] = {}
    for (canon, period), row in rows.items():
        if not canon.startswith("cf_") or row.get("months") is None:
            continue
        if period.startswith("FY"):
            fy, q = period[2:], "FY"
        else:
            q, fy = period[0], period[2:]
        by_canon_fy.setdefault((canon, fy), {})[q] = row
    for (canon, fy), qs in by_canon_fy.items():
        cumulative = {}
        for q in ("1", "2", "3"):
            row = qs.get(q)
            if row and isinstance(row.get("value"), (int, float)):
                cumulative[q] = row
        # desacumular 3Q y 2Q (de mayor a menor para no usar valores ya netos)
        for q, prev_q in (("3", "2"), ("2", "1")):
            row, prev = cumulative.get(q), cumulative.get(prev_q)
            if not row or not prev:
                continue
            if (row.get("months") or 0) <= 4:
                continue                      # ya venia trimestral
            row["value"] = row["value"] - prev["value"]
            row["tag"] = "observado (desacumulado YTD)"
            row["months"] = 3
            notes.append(f"{canon} {q}Q{fy}")
        # 4Q = FY - (1Q + 2Q + 3Q), ya netos
        fy_row = qs.get("FY")
        q4 = qs.get("4")
        if fy_row and isinstance(fy_row.get("value"), (int, float)):
            partials = [cumulative.get(q) for q in ("1", "2", "3")]
            if all(p and isinstance(p.get("value"), (int, float)) for p in partials):
                derived = fy_row["value"] - sum(p["value"] for p in partials)
                if q4 is None or (q4.get("months") or 0) > 4:
                    rows[(canon, f"4Q{fy}")] = {
                        **fy_row,
                        "canon": canon, "period": f"4Q{fy}",
                        "value": derived, "months": 3,
                        "tag": "derivado (FY - 1Q - 2Q - 3Q)",
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
    for concept, canon in CONCEPT_MAP.items():
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
    print(f"[ok] flujos desacumulados/derivados: {len(deacc)}")
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
    missing = sorted(set(CONCEPT_MAP) - found)
    if missing:
        print(f"[!] conceptos sin datos en esta emisora ({len(missing)}): "
              + ", ".join(missing[:6]) + (" ..." if len(missing) > 6 else ""))
    print("[i] siguiente paso: statement-mapper convierte este CSV largo en los")
    print("    canonical_*.csv (mapeo canon con gate del analista); 4Q de flujos")
    print("    se deriva FY - (1Q+2Q+3Q).")
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
