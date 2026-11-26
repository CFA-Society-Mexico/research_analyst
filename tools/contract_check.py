"""Validador de contratos de la cobertura (checks K), por codigo.

Los contratos de ``templates/`` los escribe el agente. Este tool dice si los
de UNA cobertura cumplen lo que el template declara, antes de que un check
del modelo tropiece con ellos:

  K1  todo YAML parsea
  K2  issuer-profile.yaml
  K3  macro/macro-view.yaml
  K4  comps/*.yaml (el snapshot vigente de cada comparable)
  K5  model/inputs/canonical_*.csv

Solo lectura: jamas modifica un contrato.

Uso:
    python tools/contract_check.py <raiz>/<TICKER> [--macro macro-view.yaml]

Severidad: falla lo ESTRUCTURAL (no parsea, obligatorio vacio, valor fuera de
su lista, CSV sin columnas, llave duplicada); avisa lo TOLERABLE (campo que el
contrato no conoce, valor deprecado con equivalente, etiqueta valida con texto
extra).

PyYAML es OPCIONAL: sin el, K1-K4 salen [pendiente] y K5 (CSV, stdlib) corre.

Exit codes: 0 = verde; 1 = alguna falla; 3 = sin fallas pero con checks
[pendiente]; 2 = no existe la carpeta. Un [aviso] nunca cambia el exit code.

Consola ASCII-only.
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

from xlsx_builder import Finding  # noqa: E402

try:
    import yaml
except ImportError:  # dependencia opcional: sin ella K1-K4 quedan [pendiente]
    yaml = None

PYYAML_HINT = "PyYAML no instalado: python -m pip install pyyaml"

K1 = "K1 los YAML parsean"
K2 = "K2 issuer-profile"
K3 = "K3 macro-view"
K4 = "K4 comp-snapshots"
K5 = "K5 capa de captura (CSV)"

# ---------------------------------------------------------------------------
# Reglas del contrato (templates/*.yaml, templates/coverage-tree.md).
# tests/test_repo_contracts.py falla si un template gana o pierde un campo de
# primer nivel sin que estas tablas se enteren.
# ---------------------------------------------------------------------------

CAPTURE_COLUMNS = ("line_item", "period", "value", "source_doc", "source_ref", "tag")
CAPTURE_TAGS = ("observado", "derivado", "guidance")
CAPTURE_PERIODS = {"canonical_annual.csv": re.compile(r"FY\d{4}"),
                   "canonical_quarterly.csv": re.compile(r"[1-4]Q\d{4}")}

PROFILE_REQUIRED = ("ticker", "name", "as_of", "listing", "issuer_type",
                    "framework_governing", "currency_reporting", "units",
                    "fiscal_year_end", "model_periodicity")
PROFILE_ENUMS = {
    "listing": ("bmv", "sec", "dual", "private"),
    "issuer_type": ("non_financial", "bank", "insurer", "fibra", "private"),
    "framework_governing": ("ifrs", "us_gaap", "nif", "cnbv_criteria"),
    "model_periodicity": ("annual", "quarterly"),
    "life_cycle_stage": ("embryonic", "growth", "mature", "decline"),
}
# valor deprecado -> como se trata hoy
PROFILE_DEPRECATED = {"model_periodicity": {"annual_plus_quarterly": "quarterly"}}
PROFILE_CONFIRMATIONS = ("framework_confirmed_by_user", "periodicity_confirmed_by_user",
                         "mnpi_confirmed_public")
PROFILE_POSITIVE_INTS = ("quarterly_history_years", "industry_staleness_months")
PROFILE_KEYS = frozenset(PROFILE_REQUIRED + PROFILE_CONFIRMATIONS + PROFILE_POSITIVE_INTS + (
    "adr", "framework_notes", "currency_functional", "segments", "life_cycle_stage",
    "valuation_methods", "materiality", "retention_note", "sources"))


MACRO_KEYS = frozenset({"updated_at", "owner", "staleness_warn_months", "gate_status",
                        "rates", "fx", "inflation", "gdp", "commodity_decks",
                        "equity_risk_premium", "notes", "countries"})
# Campos que el contrato nombra en cada bloque de countries.<pais>. El bloque es
# ABIERTO: una seccion propia del pais (labor, tax) se permite sin aviso.
COUNTRY_KEYS = frozenset({"as_of", "gate_status", "rates", "fx", "inflation", "gdp",
                          "equity_risk_premium", "notes"})
# Secciones de primer nivel y de cada pais cuyos datos se validan como
# {value, source}; los campos de texto (owner, notes, gate_status) no.
MACRO_TEXT_KEYS = frozenset({"updated_at", "owner", "staleness_warn_months", "gate_status",
                             "notes", "countries", "as_of"})

COMP_KEYS = frozenset({"ticker", "name", "selected_because", "as_of", "stale_days",
                       "market_data", "balance_items", "income_items", "framework",
                       "comparability_notes"})
COMP_MARKET_REQUIRED = ("price", "price_currency", "shares_per_ads", "fx_to_reporting",
                        "shares_outstanding")
COMP_MARKET_KEYS = frozenset(COMP_MARKET_REQUIRED + ("market_cap",))
# Datos de mercado con forma {value, source}; price_currency es un texto.
COMP_MARKET_DATA = ("price", "shares_per_ads", "fx_to_reporting", "shares_outstanding",
                    "market_cap")
COMP_FRAMEWORKS = ("ifrs", "us_gaap", "nif")
# Datos {value, source} de cada seccion, y sus campos de texto.
COMP_SECTION_DATA = {
    "balance_items": ("total_debt", "cash_and_equivalents", "minority_interest",
                      "preferred_equity", "book_value_equity"),
    "income_items": ("revenue", "ebitda", "ebit", "net_income", "eps_diluted"),
}
COMP_SECTION_TEXT = {"balance_items": (), "income_items": ("period",)}
_ISO_4217 = re.compile(r"[A-Z]{3}")
_DATED_SUFFIX = re.compile(r"_\d{4}-\d{2}-\d{2}$")


_ISO_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_EXAMPLES = 6


def _ascii(text: object) -> str:
    """Texto para consola: lo que venga de un archivo puede traer acentos."""
    return str(text).encode("ascii", "replace").decode("ascii")


def _blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_number(value: object) -> bool:
    return (isinstance(value, (int, float)) and not isinstance(value, bool)
            and not (isinstance(value, float) and math.isnan(value)))


def _date_of(value: object) -> date | None:
    """Fecha de un campo: PyYAML ya convierte ``2026-11-20`` sin comillas a
    ``date``; entre comillas llega como texto ('2026-08-31 11:05 KST')."""
    if isinstance(value, date):
        return value
    found = _ISO_DATE.search(value) if isinstance(value, str) else None
    if found is None:
        return None
    try:
        return date(int(found.group(1)), int(found.group(2)), int(found.group(3)))
    except ValueError:
        return None


def _positive_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _unknown_fields(data: dict[str, Any], known: frozenset[str], prefix: str = "") -> list[str]:
    return [f"campo desconocido: {prefix}{key}" for key in sorted(map(str, set(data) - known))]


@dataclass
class Problems:
    """Problemas de un check: los duros lo fallan, los blandos lo avisan."""

    hard: list[str] = field(default_factory=list)
    soft: list[str] = field(default_factory=list)

    def finding(self, name: str, ok_detail: str) -> Finding:
        if self.hard:
            notes = self.hard + [f"aviso: {item}" for item in self.soft]
            return Finding(name, False, _ascii(_listing(notes)))
        if self.soft:
            return Finding(name, False, _ascii(_listing(self.soft)), warning=True)
        return Finding(name, True, ok_detail)


def _listing(items: Sequence[str]) -> str:
    extra = f" (+{len(items) - _EXAMPLES} mas)" if len(items) > _EXAMPLES else ""
    return "; ".join(items[:_EXAMPLES]) + extra


# ---------------------------------------------------------------------------
# Carga de YAML y K1
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Loaded:
    """Un contrato YAML leido: ``data`` es None si no parsea o no es un mapa."""

    path: Path
    data: dict[str, Any] | None
    error: str = ""


def load_yaml(path: Path) -> Loaded:
    """Lee un contrato YAML. Requiere PyYAML (el llamador ya lo comprobo)."""
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except UnicodeDecodeError:
        return Loaded(path, None, f"{path.name}: no es UTF-8")
    except yaml.YAMLError as exc:
        mark = getattr(exc, "problem_mark", None)
        where = f" linea {mark.line + 1}, columna {mark.column + 1}" if mark else ""
        problem = getattr(exc, "problem", None) or "YAML invalido"
        return Loaded(path, None, f"{path.name}{where}: {problem}")
    if not isinstance(data, dict):
        return Loaded(path, None, f"{path.name}: el documento no es un mapa de campos")
    return Loaded(path, data)


def check_yaml_parses(loaded: Sequence[Loaded]) -> Finding:
    """K1: todo contrato YAML de la cobertura es YAML valido."""
    problems = Problems(hard=[item.error for item in loaded if item.data is None])
    return problems.finding(K1, f"{len(loaded)} archivos YAML validos")


def _unparsed(name: str) -> Finding:
    return Finding(name, False, "no evaluado: el archivo no parsea (ver K1)", pending=True)


# ---------------------------------------------------------------------------
# K5 - capa de captura
# ---------------------------------------------------------------------------


def _capture_file_problems(path: Path, problems: Problems) -> int:
    """Valida un CSV canonico; devuelve cuantas filas leyo. Los problemas se
    cuentan por regla (con la primera linea) para no listar miles de filas."""
    period_format = CAPTURE_PERIODS[path.name]
    hard: dict[str, list[int]] = {}
    soft: dict[str, list[int]] = {}
    seen: set[tuple[str, str]] = set()
    rows = 0
    try:
        with path.open(encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh)
            missing = [c for c in CAPTURE_COLUMNS if c not in (reader.fieldnames or [])]
            if missing:
                problems.hard.append(f"{path.name}: faltan columnas {', '.join(missing)}")
                return 0
            for number, row in enumerate(reader, start=2):
                rows += 1
                line_item = (row["line_item"] or "").strip()
                period = (row["period"] or "").strip()
                tag = (row["tag"] or "").strip().lower()
                if not line_item:
                    hard.setdefault("line_item vacio", []).append(number)
                if not period_format.fullmatch(period):
                    hard.setdefault("period fuera de formato", []).append(number)
                try:
                    numeric = not math.isnan(float(row["value"]))
                except (TypeError, ValueError):
                    numeric = False
                if not numeric:
                    hard.setdefault("value no numerico", []).append(number)
                if (line_item, period) in seen:
                    hard.setdefault("llave (line_item, period) repetida", []).append(number)
                seen.add((line_item, period))
                if tag not in CAPTURE_TAGS:
                    if tag.startswith(CAPTURE_TAGS):
                        soft.setdefault("tag con texto extra", []).append(number)
                    else:
                        hard.setdefault(
                            f"tag fuera de ({', '.join(CAPTURE_TAGS)})", []).append(number)
                if tag.startswith("observado") and not (row["source_doc"] or "").strip():
                    hard.setdefault("observado sin source_doc", []).append(number)
    except UnicodeDecodeError:
        problems.hard.append(f"{path.name}: no es UTF-8")
        return rows
    for tally, target in ((hard, problems.hard), (soft, problems.soft)):
        for rule, lines in tally.items():
            target.append(f"{path.name}: {len(lines)} filas con {rule} (linea {lines[0]})")
    return rows


def check_capture(inputs_dir: Path) -> Finding:
    """K5: esquema de los CSV canonicos de la capa de captura."""
    files = [inputs_dir / name for name in CAPTURE_PERIODS if (inputs_dir / name).is_file()]
    if not files:
        return Finding(K5, True, "n/a (sin canonical_*.csv en model/inputs/)")
    problems = Problems()
    rows = sum(_capture_file_problems(path, problems) for path in files)
    return problems.finding(K5, f"{rows} filas en {len(files)} archivos, esquema en orden")


def audit_contracts(coverage: Path, macro_view: Path | None = None) -> list[Finding]:
    """Checks K sobre la carpeta de una cobertura (``<raiz>/<TICKER>``)."""
    capture = check_capture(coverage / "model" / "inputs")
    if yaml is None:
        return [*(Finding(name, False, PYYAML_HINT, pending=True) for name in (K1, K2, K3, K4)),
                capture]
    macro_path = macro_view or coverage.parent / "macro" / "macro-view.yaml"
    profile_path = coverage / "issuer-profile.yaml"
    comps_dir = coverage / "comps"
    profile = load_yaml(profile_path) if profile_path.is_file() else None
    macro = load_yaml(macro_path) if macro_path.is_file() else None
    comps = [load_yaml(p) for p in sorted(comps_dir.glob("*.yaml"))] if comps_dir.is_dir() else []
    loaded = [item for item in (profile, macro, *comps) if item is not None]
    return [check_yaml_parses(loaded), check_profile(profile), check_macro(macro),
            check_comps(comps), capture]


def check_profile(profile: Loaded | None) -> Finding:
    """K2: issuer-profile.yaml cumple su contrato."""
    if profile is None:
        return Finding(K2, False, "no existe issuer-profile.yaml en la cobertura")
    if profile.data is None:
        return _unparsed(K2)
    data, problems = profile.data, Problems()
    problems.hard += [f"{key} vacio" for key in PROFILE_REQUIRED if _blank(data.get(key))]
    for key, allowed in PROFILE_ENUMS.items():
        value = data.get(key)
        if _blank(value):
            continue
        replacement = PROFILE_DEPRECATED.get(key, {}).get(value)
        if replacement:
            problems.soft.append(f"{key}: {value} esta deprecado (se trata como {replacement})")
        elif value not in allowed:
            problems.hard.append(f"{key}: {value!r} no esta en ({', '.join(allowed)})")
    if not _blank(data.get("as_of")) and _date_of(data["as_of"]) is None:
        problems.hard.append("as_of sin fecha YYYY-MM-DD")
    fiscal_year_end = data.get("fiscal_year_end")
    if not _blank(fiscal_year_end) and not re.fullmatch(r"\d{2}-\d{2}", str(fiscal_year_end)):
        # Aviso y no falla: un anio fiscal de 52/53 semanas ("ultimo sabado de
        # septiembre") es legitimo y no cabe en MM-DD.
        problems.soft.append(f"fiscal_year_end no es MM-DD ({fiscal_year_end})")
    problems.hard += [f"{key} no esta en true (falta la confirmacion del analista)"
                      for key in PROFILE_CONFIRMATIONS if data.get(key) is not True]
    if "adr" in data and not isinstance(data["adr"], bool):
        problems.hard.append("adr no es booleano")
    problems.hard += [f"{key} debe ser un entero > 0" for key in PROFILE_POSITIVE_INTS
                      if key in data and not _positive_int(data[key])]
    for key, value in (data.get("materiality") or {}).items():
        if not _is_number(value) or value < 0:
            problems.hard.append(f"materiality.{key} debe ser un numero >= 0")
    for key, value in (data.get("valuation_methods") or {}).items():
        if not isinstance(value, bool):
            problems.hard.append(f"valuation_methods.{key} no es booleano")
    problems.soft += _unknown_fields(data, PROFILE_KEYS)
    return problems.finding(K2, "obligatorios, listas y confirmaciones en orden")


def check_macro(macro: Loaded | None) -> Finding:
    """K3: macro/macro-view.yaml cumple su contrato."""
    if macro is None:
        return Finding(K3, False, "sin macro-view.yaml en la carpeta macro/ del workspace",
                       warning=True)
    if macro.data is None:
        return _unparsed(K3)
    data, problems = macro.data, Problems()
    if _date_of(data.get("updated_at")) is None:
        problems.hard.append("updated_at vacio o sin fecha")
    if "staleness_warn_months" in data and not _positive_int(data["staleness_warn_months"]):
        problems.hard.append("staleness_warn_months debe ser un entero > 0")
    for key, value in data.items():
        if key not in MACRO_TEXT_KEYS:
            _macro_data_problems(value, str(key), False, problems)
    countries = data.get("countries")
    if countries is None:
        countries = {}
    if not isinstance(countries, dict):
        problems.hard.append("countries no es un mapa {pais: bloque}")
        countries = {}
    for name, block in countries.items():
        _country_problems(f"countries.{name}", block, problems)
    for key in sorted(map(str, set(data) - MACRO_KEYS)):
        if _looks_like_country(data[key]):
            problems.soft.append(f"campo desconocido: {key} (parece un bloque por pais: va "
                                 "bajo countries:, ver la migracion en templates/coverage-tree.md)")
        else:
            problems.soft.append(f"campo desconocido: {key}")
    plural = "pais" if len(countries) == 1 else "paises"
    return problems.finding(
        K3, f"updated_at, plazo y datos {{value, source}} en orden ({len(countries)} {plural})")


def _looks_like_country(value: object) -> bool:
    """Un mapa de primer nivel que trae secciones de un bloque por pais: lo que
    deja un workspace anterior a countries: (p. ej. ``mexico:`` suelto)."""
    return isinstance(value, dict) and bool(
        set(map(str, value)) & {"rates", "fx", "inflation", "gdp", "equity_risk_premium"})


def _country_problems(path: str, block: object, problems: Problems) -> None:
    """Un bloque de countries.<pais>: abierto, con as_of y los insumos del CAPM."""
    if not isinstance(block, dict):
        problems.hard.append(f"{path} no es un mapa de campos")
        return
    if "as_of" not in block:
        problems.soft.append(f"{path} sin as_of")
    elif _date_of(block["as_of"]) is None:
        problems.hard.append(f"{path}.as_of sin fecha YYYY-MM-DD")
    rates = block.get("rates")
    if not isinstance(rates, dict) or _empty_datum(rates.get("risk_free_10y")):
        problems.soft.append(f"{path} sin rates.risk_free_10y (insumo del CAPM)")
    if _empty_datum(block.get("equity_risk_premium")):
        problems.soft.append(f"{path} sin equity_risk_premium (insumo del CAPM)")
    for key, value in block.items():
        if key not in MACRO_TEXT_KEYS:
            _macro_data_problems(value, f"{path}.{key}", False, problems)


def _macro_data_problems(node: object, path: str, cited: bool, problems: Problems) -> None:
    """Recorre una seccion del macro-view. Todo mapa con ``value`` es un dato:
    ``value`` numerico y ``source`` propio o de un mapa que lo contiene (la cita
    de un ERP cubre sus componentes; la de una ruta, sus puntos). ``value: null``
    con source es un pendiente declarado: aviso. ``{}`` y ``[]`` no son datos."""
    if isinstance(node, list):
        for index, item in enumerate(node):
            _macro_data_problems(item, f"{path}[{index}]", cited, problems)
        return
    if not isinstance(node, dict):
        return
    cited = cited or not _blank(node.get("source"))
    if "value" in node:
        value = node["value"]
        if not cited:
            problems.hard.append(f"{path} sin source")
        elif value is None:
            problems.soft.append(f"{path}: value pendiente (null)")
        elif not _is_number(value):
            problems.hard.append(f"{path} sin value numerico")
    for key, child in node.items():
        if key != "value":
            _macro_data_problems(child, f"{path}.{key}", cited, problems)


def _empty_datum(value: object) -> bool:
    """Dato no capturado: ausente, texto vacio o ``{}``."""
    return _blank(value) or value == {}


def _datum_problem(name: str, value: object) -> str | None:
    """Un dato capturado lleva ``value`` numerico y ``source``."""
    if not isinstance(value, dict):
        return f"{name} no es {{value, source}}"
    if not _is_number(value.get("value")):
        return f"{name} sin value numerico"
    if _blank(value.get("source")):
        return f"{name} sin source"
    return None


def _snapshot_problems(data: dict[str, Any]) -> tuple[list[str], list[str]]:
    hard: list[str] = []
    soft: list[str] = []
    if _blank(data.get("ticker")):
        hard.append("ticker vacio")
    if _date_of(data.get("as_of")) is None:
        hard.append("as_of sin fecha")
    if "stale_days" in data and not _positive_int(data["stale_days"]):
        hard.append("stale_days debe ser un entero > 0")
    market = data.get("market_data")
    if not isinstance(market, dict):
        hard.append("sin market_data")
        market = {}
    else:
        missing = [key for key in COMP_MARKET_REQUIRED if _empty_datum(market.get(key))]
        if missing:
            hard.append("market_data sin " + ", ".join(missing))
        currency = market.get("price_currency")
        if not _blank(currency) and not _ISO_4217.fullmatch(str(currency)):
            hard.append("price_currency no es ISO 4217 (tres letras mayusculas)")
        for key in COMP_MARKET_DATA:
            if not _empty_datum(market.get(key)):
                problem = _datum_problem(f"market_data.{key}", market[key])
                if problem:
                    hard.append(problem)
        fx = market.get("fx_to_reporting")
        if isinstance(fx, dict) and fx and _blank(fx.get("as_of")):
            hard.append("market_data.fx_to_reporting sin as_of")
        soft += _unknown_fields(market, COMP_MARKET_KEYS, prefix="market_data.")
    for section, data_keys in COMP_SECTION_DATA.items():
        content = data.get(section) or {}
        if not isinstance(content, dict):
            hard.append(f"{section} no es un mapa de campos")
            continue
        for key in data_keys:
            if not _empty_datum(content.get(key)):
                problem = _datum_problem(f"{section}.{key}", content[key])
                if problem:
                    hard.append(problem)
        # Un campo de texto que el contrato no conoce (p. ej. `currency`) no es
        # un dato roto: se avisa como desconocido.
        known = frozenset(data_keys + COMP_SECTION_TEXT[section])
        soft += _unknown_fields(content, known, prefix=f"{section}.")
    framework = data.get("framework")
    if not _blank(framework) and framework not in COMP_FRAMEWORKS:
        soft.append(f"framework {framework!r} fuera de ({', '.join(COMP_FRAMEWORKS)})")
    soft += _unknown_fields(data, COMP_KEYS)
    return hard, soft


def _current_snapshots(comps: Sequence[Loaded]) -> list[Loaded]:
    """El snapshot vigente de cada comparable: el de ``as_of`` mas reciente
    (mismo criterio que D3). Los anteriores son historia y no se validan."""
    latest: dict[str, tuple[date, Loaded]] = {}
    for item in comps:
        if item.data is None:
            continue
        key = str(item.data.get("ticker") or _DATED_SUFFIX.sub("", item.path.stem))
        as_of = _date_of(item.data.get("as_of")) or date.min
        if key not in latest or as_of > latest[key][0]:
            latest[key] = (as_of, item)
    return [item for _, item in sorted(latest.values(), key=lambda pair: pair[1].path.name)]


def check_comps(comps: Sequence[Loaded]) -> Finding:
    """K4: el snapshot vigente de cada comparable cumple su contrato."""
    if not comps:
        return Finding(K4, True, "n/a (sin snapshots en comps/)")
    current = _current_snapshots(comps)
    if not current:
        return _unparsed(K4)
    problems = Problems()
    for item in current:
        hard, soft = _snapshot_problems(item.data or {})
        if hard:
            problems.hard.append(f"{item.path.name}: " + ", ".join(hard))
        if soft:
            problems.soft.append(f"{item.path.name}: " + ", ".join(soft))
    return problems.finding(K4, f"{len(current)} snapshots vigentes completos")


# ---------------------------------------------------------------------------
# Reporte y CLI
# ---------------------------------------------------------------------------


def print_report(findings: Sequence[Finding]) -> int:
    """Imprime el reporte. Exit: 1 fallas; 3 solo pendientes; 0 verde.
    Los avisos se listan y se cuentan, pero no cambian el exit code."""
    for finding in findings:
        print(f"{finding.mark} {finding.check}: {finding.detail}")
    failures = sum(1 for f in findings if f.blocking)
    pending = sum(1 for f in findings if f.pending)
    warnings = sum(1 for f in findings if f.warning and not f.pending)
    ok = sum(1 for f in findings if f.ok and not f.pending)
    print(f"Resumen: {ok} ok, {failures} fallas, {pending} pendientes, {warnings} avisos")
    if failures:
        print("Veredicto: FALLA")
        return 1
    if pending:
        print("Veredicto: PENDIENTE - no es verde; resuelve los pendientes y re-valida")
        return 3
    print("Veredicto: VERDE (contratos de la cobertura)")
    return 0


def main(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Checks K: contratos de una cobertura contra templates/ (solo lectura).")
    parser.add_argument("coverage", help="carpeta de la cobertura: <raiz>/<TICKER>")
    parser.add_argument("--macro",
                        help="macro-view.yaml (default: <raiz>/macro/macro-view.yaml)")
    args = parser.parse_args(list(argv[1:]))
    coverage = Path(args.coverage)
    if not coverage.is_dir():
        print(f"[x] no existe la carpeta de cobertura: {_ascii(args.coverage)}")
        return 2
    return print_report(audit_contracts(coverage, Path(args.macro) if args.macro else None))


if __name__ == "__main__":
    sys.exit(main(sys.argv))
