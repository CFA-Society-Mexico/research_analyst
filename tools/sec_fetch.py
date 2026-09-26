"""SEC EDGAR filing downloader for the research_analyst plugin.

Deterministic fetch step for /init-coverage and /update-quarter: when the
coverage folders lack filings for an SEC issuer, this tool downloads them from
EDGAR (free, public). It downloads RAW primary documents named by report date;
renaming to the coverage-tree convention (FYyyyy / #Qyyyy) is coverage-folders'
job — that skill owns naming, this tool only fetches.

SEC fair-access rules: a User-Agent identifying the requester is REQUIRED
(name + email). Pass it with --ua or the SEC_EDGAR_UA env var; it is never
hardcoded (public repo). Requests are throttled to ~4/s.

Usage:
    python tools/sec_fetch.py AAPL --dest <raiz>/AAPL/filings/sec \
        --ua "Nombre Apellido correo@dominio.com"
    python tools/sec_fetch.py AAPL --forms 10-K,10-Q --since 2020-01-01 --dry-run

The manifest (<TICKER>_filings_manifest.csv) is MERGED across runs: a run with
--since never drops the provenance of filings downloaded earlier.

Console output: ASCII only.
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
from pathlib import Path
from typing import Iterable, Optional

SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/{name}"
SEC_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accn}/{doc}"
SEC_FOLDER_INDEX_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accn}/index.json"

# Files never worth downloading from a filing folder (graphics, XBRL plumbing).
_EXHIBIT_SKIP_SUFFIXES = (".jpg", ".jpeg", ".gif", ".png", ".xml", ".xsd",
                          ".css", ".js", ".json", ".txt", ".zip", ".paper")

DEFAULT_FORMS = ("10-K", "10-Q", "8-K")
# Perspectiva de largo plazo: el default es TODA la historia electronica de
# EDGAR (arranca 1994) — ciclos completos de margenes y drivers. --since existe
# para que el ANALISTA recorte, no para proponer subconjuntos.
DEFAULT_SINCE = "1994-01-01"
THROTTLE_SECONDS = 0.25
# Errores transitorios: se reintenta con espera exponencial (1s, 2s, 4s).
RETRY_CODES = (429, 500, 502, 503, 504)
RETRIES = 3
MANIFEST_FIELDS = ["file", "form", "filed", "period", "accession", "source"]


@dataclass(frozen=True)
class FilingRow:
    """One EDGAR filing entry (metadata only)."""

    form: str
    filed: str
    period: str
    accession: str
    primary_doc: str


class SecHttpError(Exception):
    """Error HTTP de SEC que NO es bloqueo de red (404, 429 persistente, 5xx).

    Es una Exception normal (no SystemExit) a proposito: un exhibit o un
    filing que falla se reporta y la descarga sigue con el resto.
    """


BLOCKED_MSG = """[x] Sin acceso de red a SEC ({err}).
    El entorno bloquea www.sec.gov / data.sec.gov (proxy con allowlist -
    tipico en Claude Cowork y sandboxes corporativos). Opciones:
    1) Permitir los dominios www.sec.gov y data.sec.gov en el entorno
       (Cowork: configuracion de red/allowlist del espacio) y reintentar.
    2) Correr este mismo comando en una maquina con salida a internet y
       copiar los archivos resultantes a la carpeta destino.
    3) Descargar los filings a mano desde efts.sec.gov/LATEST/search-index
       y dejar que coverage-folders los archive.
    Nota: SEC tambien responde 403 si el User-Agent no trae contacto valido
    o si se excede la tasa de solicitudes; revisa --ua antes de tocar la red.
    NUNCA sustituyas el filing integro por contenido procesado de un lector
    web: el pipeline cita por documento y pagina."""


def _fetch_bytes(url: str, user_agent: str, timeout: int) -> bytes:
    """GET con clasificacion de errores.

    HTTPError se atrapa ANTES que URLError porque es su subclase: sin ese
    orden, un 404 o un 429 se reportaban como "red bloqueada".
    """
    req = urllib.request.Request(url, headers={"User-Agent": user_agent})
    for attempt in range(RETRIES + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except urllib.error.HTTPError as exc:
            if exc.code in (403, 407):
                raise SystemExit(BLOCKED_MSG.format(err=f"HTTP {exc.code}"))
            if exc.code in RETRY_CODES and attempt < RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise SecHttpError(f"HTTP {exc.code} en {url}") from exc
        except urllib.error.URLError as exc:
            raise SystemExit(BLOCKED_MSG.format(err=f"URLError {exc.reason}"))
        except (TimeoutError, ConnectionError) as exc:
            if attempt < RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise SecHttpError(f"sin respuesta ({type(exc).__name__}) en {url}") from exc
    raise SecHttpError(f"reintentos agotados en {url}")  # pragma: no cover


def _get_json(url: str, user_agent: str) -> dict:
    return json.loads(_fetch_bytes(url, user_agent, timeout=60).decode("utf-8"))


def _download(url: str, dest: Path, user_agent: str) -> None:
    """Escritura atomica: un corte a medias no deja un archivo truncado que la
    siguiente corrida daria por bueno ("ya existe")."""
    data = _fetch_bytes(url, user_agent, timeout=120)
    tmp = dest.with_name(dest.name + ".part")
    tmp.write_bytes(data)
    os.replace(tmp, dest)


def resolve_cik(ticker: str, user_agent: str) -> str:
    """Ticker -> zero-padded 10-digit CIK via SEC's public mapping."""
    data = _get_json(SEC_TICKERS_URL, user_agent)
    wanted = ticker.upper()
    for entry in data.values():
        if str(entry.get("ticker", "")).upper() == wanted:
            return f"{int(entry['cik_str']):010d}"
    raise SystemExit(f"[x] ticker {wanted} no encontrado en SEC company_tickers.json")


def list_filings(cik10: str, forms: Iterable[str], since: str,
                 user_agent: str, amendments: bool = False
                 ) -> tuple[list[FilingRow], int]:
    """Matching filings (oldest first) and how many amendments were left out.

    Las enmiendas (``10-K/A``...) solo se incluyen con ``amendments=True``;
    si no, se CUENTAN para avisar al analista en vez de omitirlas en silencio.
    """
    wanted = {f.upper() for f in forms}
    if amendments:
        wanted |= {f"{f}/A" for f in wanted if not f.endswith("/A")}
    root = _get_json(SEC_SUBMISSIONS_URL.format(name=f"CIK{cik10}.json"), user_agent)
    batches = [root["filings"]["recent"]]
    for extra in root["filings"].get("files", []):
        time.sleep(THROTTLE_SECONDS)
        batches.append(_get_json(SEC_SUBMISSIONS_URL.format(name=extra["name"]),
                                 user_agent))
    rows: list[FilingRow] = []
    omitted = 0
    for batch in batches:
        for i in range(len(batch["form"])):
            form = batch["form"][i]
            filed = batch["filingDate"][i]
            if filed < since:
                continue
            if form.upper() in wanted:
                rows.append(FilingRow(
                    form=form,
                    filed=filed,
                    period=batch["reportDate"][i] or filed,
                    accession=batch["accessionNumber"][i],
                    primary_doc=batch["primaryDocument"][i],
                ))
            elif form.upper().endswith("/A") and form.upper()[:-2] in wanted:
                omitted += 1
    return sorted(rows, key=lambda r: r.filed), omitted


def list_8k_exhibits(cik_short: str, accn_flat: str, primary_doc: str,
                     user_agent: str) -> list[str]:
    """Exhibit documents of one 8-K folder (press release EX-99.* lives here).

    Deterministic filter over the accession's index.json: every .htm that is
    not the primary wrapper nor an index page. 8-K folders are small (wrapper +
    1-2 exhibits + plumbing), so this stays precise without HTML parsing.
    """
    url = SEC_FOLDER_INDEX_URL.format(cik=cik_short, accn=accn_flat)
    data = _get_json(url, user_agent)
    items = data.get("directory", {}).get("item", [])
    exhibits: list[str] = []
    primary_lower = primary_doc.lower()
    for item in items:
        name = str(item.get("name", ""))
        low = name.lower()
        if not low.endswith(".htm") and not low.endswith(".html"):
            continue
        if low == primary_lower or "index" in low:
            continue
        if re.fullmatch(r"r\d+\.htm", low):  # XBRL viewer artifact, not a doc
            continue
        if low.endswith(_EXHIBIT_SKIP_SUFFIXES):
            continue
        exhibits.append(name)
    return exhibits


def _safe_form(form: str) -> str:
    """'10-K/A' -> '10-K-A': una diagonal en el nombre crearia una carpeta."""
    return form.replace("/", "-")


def read_manifest(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def write_manifest(path: Path, old: list[dict[str, str]],
                   new: list[dict[str, str]]) -> None:
    """MERGE por documento de origen (URL): la corrida nueva actualiza sus
    filas y jamas borra las de corridas anteriores (R5 — nada se borra). Una
    fila vieja que apuntaba al archivo equivocado (choque de nombres) queda
    reemplazada por la corregida, porque comparten URL de origen."""
    merged: dict[str, dict[str, str]] = {}
    for row in [*old, *new]:
        key = row.get("source") or f"{row.get('accession', '')}|{row.get('file', '')}"
        merged[key] = row
    fields = list(MANIFEST_FIELDS)
    for row in merged.values():
        fields.extend(k for k in row if k not in fields)
    rows = sorted(merged.values(), key=lambda r: (r.get("filed", ""), r.get("file", "")))
    tmp = path.with_name(path.name + ".part")
    with tmp.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    os.replace(tmp, path)


class _NameRegistry:
    """Asigna nombres unicos por accession.

    El nombre crudo es ``<TICKER>_<form>_<periodo>``. Dos filings distintos
    con el mismo periodo (dos 8-K del mismo dia de evento) chocaban: el
    segundo no se bajaba y el manifest apuntaba dos veces al primero. El
    primer dueno conserva el nombre corto; los demas llevan la accession.
    """

    def __init__(self, manifest_rows: list[dict[str, str]]) -> None:
        self._owner: dict[str, str] = {}
        for row in manifest_rows:
            self._owner.setdefault(row.get("file", ""), row.get("accession", ""))

    def claim(self, base: str, ext: str, accession: str) -> str:
        name = f"{base}{ext}"
        owner = self._owner.setdefault(name, accession)
        if owner != accession:
            name = f"{base}_{accession}{ext}"
            self._owner.setdefault(name, accession)
        return name


def fetch(ticker: str, dest: Path, forms: Iterable[str], since: str,
          user_agent: str, dry_run: bool = False,
          amendments: bool = False) -> tuple[Path, list[str]]:
    """Download filings + merge the manifest. Returns (manifest, failures)."""
    forms = list(forms)
    tick = ticker.upper()
    cik10 = resolve_cik(ticker, user_agent)
    cik_short = str(int(cik10))
    rows, omitted = list_filings(cik10, forms, since, user_agent, amendments)
    print(f"[ok] {tick} CIK {cik10}: {len(rows)} filings "
          f"({','.join(forms)} desde {since})")
    if omitted:
        print(f"[aviso] {omitted} enmiendas (/A) omitidas; usa --amendments "
              "para bajarlas (coverage-folders las archiva con _amended)")
    dest.mkdir(parents=True, exist_ok=True)
    manifest = dest / f"{tick}_filings_manifest.csv"
    old_rows = read_manifest(manifest)
    names = _NameRegistry(old_rows)
    new_rows: list[dict[str, str]] = []
    failures: list[str] = []

    def get(url: str, name: str, label: str) -> bool:
        path = dest / name
        if dry_run:
            print(f"[dry] {name}  <-  {url}")
            return True
        if path.exists():
            print(f"[ok] ya existe: {name}")
            return True
        print(f"[..] bajando{label}: {name}")
        try:
            _download(url, path, user_agent)
        except SecHttpError as exc:
            print(f"[x] {name}: {exc}")
            failures.append(name)
            return False
        finally:
            time.sleep(THROTTLE_SECONDS)
        return True

    try:
        for row in rows:
            accn_flat = row.accession.replace("-", "")
            url = SEC_ARCHIVES_URL.format(cik=cik_short, accn=accn_flat,
                                          doc=row.primary_doc)
            ext = Path(row.primary_doc).suffix or ".htm"
            # Raw name by report date; coverage-folders renames to FY/#Q convention.
            name = names.claim(f"{tick}_{_safe_form(row.form)}_{row.period}", ext,
                               row.accession)
            if get(url, name, ""):
                new_rows.append({
                    "file": name, "form": row.form, "filed": row.filed,
                    "period": row.period, "accession": row.accession, "source": url,
                })
            # 8-K: also fetch exhibits — EX-99.* press release carries the
            # quarter's quantitative guidance (transcripts do NOT live on EDGAR).
            if row.form.upper().startswith("8-K"):
                time.sleep(THROTTLE_SECONDS)
                try:
                    exhibits = list_8k_exhibits(cik_short, accn_flat,
                                                row.primary_doc, user_agent)
                except SecHttpError as exc:
                    print(f"[aviso] exhibits no listados para {row.accession}: {exc}")
                    exhibits = []
                for ex_doc in exhibits:
                    ex_url = SEC_ARCHIVES_URL.format(cik=cik_short, accn=accn_flat,
                                                     doc=ex_doc)
                    ex_base = (f"{tick}_{_safe_form(row.form)}_{row.period}"
                               f"_ex_{Path(ex_doc).stem}")
                    ex_name = names.claim(ex_base, Path(ex_doc).suffix, row.accession)
                    if get(ex_url, ex_name, " exhibit"):
                        new_rows.append({
                            "file": ex_name, "form": f"{row.form}-EX",
                            "filed": row.filed, "period": row.period,
                            "accession": row.accession, "source": ex_url,
                        })
    finally:
        # Siempre se escribe (tambien si la red se corta a medias): lo ya
        # descargado conserva su procedencia.
        if not dry_run and (new_rows or old_rows):
            write_manifest(manifest, old_rows, new_rows)
            print(f"[ok] manifest: {manifest} ({len(new_rows)} filas de esta corrida)")
    if failures:
        print(f"[x] {len(failures)} descargas fallaron (reintenta la corrida; "
              "lo ya bajado no se repite)")
    return manifest, failures


def _resolve_user_agent(cli_value: Optional[str]) -> str:
    ua = cli_value or os.environ.get("SEC_EDGAR_UA", "")
    if not ua or "@" not in ua:
        raise SystemExit(
            "[x] SEC exige User-Agent con contacto (nombre + email).\n"
            "    Pasa --ua \"Nombre correo@dominio.com\" o define SEC_EDGAR_UA.")
    return ua


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Descarga filings SEC EDGAR")
    parser.add_argument("ticker")
    parser.add_argument("--forms", default=",".join(DEFAULT_FORMS),
                        help="lista separada por comas (default: 10-K,10-Q,8-K)")
    parser.add_argument("--since", default=DEFAULT_SINCE, help="YYYY-MM-DD")
    parser.add_argument("--dest", default=".", help="carpeta destino")
    parser.add_argument("--ua", default=None,
                        help="User-Agent 'Nombre correo' (o env SEC_EDGAR_UA)")
    parser.add_argument("--amendments", action="store_true",
                        help="incluir enmiendas (10-K/A, 10-Q/A, 8-K/A)")
    parser.add_argument("--dry-run", action="store_true",
                        help="listar sin descargar")
    args = parser.parse_args(argv[1:])
    user_agent = _resolve_user_agent(args.ua)
    forms = [f.strip() for f in args.forms.split(",") if f.strip()]
    try:
        _, failures = fetch(args.ticker, Path(args.dest), forms, args.since,
                            user_agent, dry_run=args.dry_run,
                            amendments=args.amendments)
    except SecHttpError as exc:
        print(f"[x] SEC: {exc}")
        return 1
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
