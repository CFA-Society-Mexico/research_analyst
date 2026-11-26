"""SEC EDGAR filing downloader for the research_analyst plugin.

Deterministic fetch step for /init-coverage and /update-quarter: when the
coverage folders lack filings for an SEC issuer, this tool downloads them from
EDGAR (free, public). It downloads RAW documents; renaming to the coverage-tree
convention (FYyyyy / #Qyyyy) is coverage-folders' job — that skill owns naming,
this tool only fetches.

Raw names are deterministic (they depend only on the filing, never on the order
of the run nor on the manifest):
    <TICKER>_<form>_<reportDate>_<filingDate>_<accn6><ext>              primary
    <TICKER>_<form>_<reportDate>_<filingDate>_<accn6>_ex_<stem><ext>    exhibit
where accn6 = last 6 digits of the accession number. Several 8-K / 6-K of one
period therefore never share a file.

Forms: without --forms the issuer type is detected from its EDGAR submissions
(10-K,10-Q,8-K; foreign private issuer: 20-F,6-K or 40-F,6-K). Exhibits are
fetched for 8-K and 6-K (press release / quarterly report live there).

SEC fair-access rules: a User-Agent identifying the requester is REQUIRED
(name + email). Pass it with --ua or the SEC_EDGAR_UA env var; it is never
hardcoded (public repo). Requests are throttled to ~4/s (--throttle); transient
errors are retried with exponential backoff.

Usage:
    python tools/sec_fetch.py AAPL --dest <raiz>/AAPL/filings/sec \
        --ua "Nombre Apellido correo@dominio.com"
    python tools/sec_fetch.py AAPL --forms 10-K,10-Q --since 2020-01-01 --dry-run
    python tools/sec_fetch.py TSM --dest <raiz>/TSM/filings/sec --log sec_fetch.log

The manifest (<TICKER>_filings_manifest.csv) is MERGED across runs and persisted
DURING the run: a run with --since never drops the provenance of filings
downloaded earlier, and a killed run keeps the provenance of what it downloaded.

Console output: ASCII only.
"""

from __future__ import annotations

import argparse
import contextlib
import csv
import http.client
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Iterator, Optional, TextIO

SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/{name}"
SEC_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accn}/{doc}"
SEC_FOLDER_INDEX_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accn}/index.json"

DEFAULT_FORMS = ("10-K", "10-Q", "8-K")
# Emisora extranjera (foreign private issuer): anual 20-F / 40-F, corriente 6-K.
FOREIGN_ANNUAL_FORMS = ("20-F", "40-F")
FOREIGN_CURRENT_FORM = "6-K"
# Formas cuyo contenido vive en los exhibits (press release, reporte trimestral
# de una emisora extranjera): se baja el documento principal Y sus exhibits.
EXHIBIT_FORMS = ("8-K", "6-K")
EXHIBIT_SUFFIXES = (".htm", ".html", ".pdf")
# Perspectiva de largo plazo: el default es TODA la historia electronica de
# EDGAR (arranca 1994) — ciclos completos de margenes y drivers. --since existe
# para que el ANALISTA recorte, no para proponer subconjuntos.
DEFAULT_SINCE = "1994-01-01"
THROTTLE_SECONDS = 0.25
# Errores transitorios: hasta 6 reintentos con espera exponencial.
BACKOFF_SECONDS = (2, 4, 8, 16, 32, 64)
# 403 / URLError sin ningun exito previo en ese host = red bloqueada: un solo
# reintento corto y se sale con BLOCKED_MSG (no se espera 2 minutos a un proxy).
UNPROVEN_RETRIES = 1
# Documentos seguidos con reintentos agotados que detienen la corrida.
BREAKER_LIMIT = 3
MANIFEST_FIELDS = ["file", "form", "filed", "period", "accession", "source"]

# Unica via de espera del modulo: las pruebas la sustituyen y jamas duermen.
_sleep = time.sleep
# Hosts que ya respondieron bien en este proceso (red probada).
_proven_hosts: set[str] = set()


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


class SecRetriesExhausted(SecHttpError):
    """Error transitorio que siguio fallando tras todos los reintentos."""


class _CircuitOpen(Exception):
    """BREAKER_LIMIT documentos seguidos agotaron reintentos: se corta la corrida."""


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


# ---- HTTP -----------------------------------------------------------------

def _classify(exc: Exception) -> tuple[str, str]:
    """(clase, detalle) de un fallo de red.

    ``blocked``   proxy que exige autenticacion: nunca se reintenta.
    ``ambiguous`` 403 / URLError: bloqueo si el host aun no respondio bien,
                  transitorio (rate limit, corte) si ya lo hizo.
    ``transient`` 429, 5xx, timeout, conexion cortada: siempre se reintenta.
    ``fatal``     respuesta definitiva del servidor (404...): no se reintenta.

    HTTPError se revisa ANTES que URLError porque es su subclase: sin ese
    orden, un 404 o un 429 se reportaban como "red bloqueada".
    """
    if isinstance(exc, urllib.error.HTTPError):
        detail = f"HTTP {exc.code}"
        if exc.code == 407:
            return "blocked", detail
        if exc.code == 403:
            return "ambiguous", detail
        if exc.code == 429 or 500 <= exc.code < 600:
            return "transient", detail
        return "fatal", detail
    if isinstance(exc, urllib.error.URLError):
        # El texto del sistema operativo viene en el idioma local: a ASCII
        # antes de llegar a la consola (cp1252 / ascii estricto).
        reason = str(exc.reason).encode("ascii", "replace").decode("ascii")
        return "ambiguous", f"URLError {reason}"
    return "transient", type(exc).__name__


def _retry_wait(exc: Exception, attempt: int, url: str, proven: bool) -> tuple[int, str]:
    """Segundos a esperar antes del reintento ``attempt`` (base 0), o lanza.

    Red no probada + fallo ambiguo => ``SystemExit(BLOCKED_MSG)`` tras un
    reintento corto. Con la red probada jamas se sale por un solo documento:
    agotados los reintentos se lanza ``SecRetriesExhausted``.
    """
    kind, detail = _classify(exc)
    if kind == "blocked" or (kind == "ambiguous" and not proven
                             and attempt >= UNPROVEN_RETRIES):
        raise SystemExit(BLOCKED_MSG.format(err=detail)) from None
    if kind == "fatal":
        raise SecHttpError(f"{detail} en {url}") from exc
    if attempt >= len(BACKOFF_SECONDS):
        raise SecRetriesExhausted(
            f"{detail} en {url} ({len(BACKOFF_SECONDS)} reintentos agotados)") from exc
    return BACKOFF_SECONDS[attempt], detail


def _fetch_bytes(url: str, user_agent: str, timeout: int) -> bytes:
    """GET con reintentos; clasifica el fallo en bloqueo, transitorio o definitivo."""
    req = urllib.request.Request(url, headers={"User-Agent": user_agent})
    host = urllib.parse.urlsplit(url).netloc
    attempt = 0
    while True:
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = resp.read()
        except (OSError, http.client.HTTPException) as exc:
            if isinstance(exc, urllib.error.HTTPError):
                exc.close()  # suelta la respuesta de error antes de reintentar
            wait, detail = _retry_wait(exc, attempt, url, host in _proven_hosts)
            print(f"[..] {detail}: reintento {attempt + 1} en {wait} s")
            _sleep(wait)
            attempt += 1
        else:
            _proven_hosts.add(host)
            return data


def _get_json(url: str, user_agent: str) -> dict:
    return json.loads(_fetch_bytes(url, user_agent, timeout=60).decode("utf-8"))


def _download(url: str, dest: Path, user_agent: str) -> None:
    """Escritura atomica: un corte a medias no deja un archivo truncado que la
    siguiente corrida daria por bueno ("ya existe")."""
    data = _fetch_bytes(url, user_agent, timeout=120)
    tmp = dest.with_name(dest.name + ".part")
    tmp.write_bytes(data)
    os.replace(tmp, dest)


# ---- EDGAR listings -------------------------------------------------------

def resolve_cik(ticker: str, user_agent: str) -> str:
    """Ticker -> zero-padded 10-digit CIK via SEC's public mapping."""
    data = _get_json(SEC_TICKERS_URL, user_agent)
    wanted = ticker.upper()
    for entry in data.values():
        if str(entry.get("ticker", "")).upper() == wanted:
            return f"{int(entry['cik_str']):010d}"
    raise SystemExit(f"[x] ticker {wanted} no encontrado en SEC company_tickers.json")


def load_submissions(cik10: str, user_agent: str,
                     throttle: float = THROTTLE_SECONDS) -> list[dict]:
    """Every batch of the issuer's submissions index (recent + older pages).

    Fetched ONCE per run: the same batches feed form detection and the listing.
    """
    root = _get_json(SEC_SUBMISSIONS_URL.format(name=f"CIK{cik10}.json"), user_agent)
    batches = [root["filings"]["recent"]]
    for extra in root["filings"].get("files", []):
        _sleep(throttle)
        batches.append(_get_json(SEC_SUBMISSIONS_URL.format(name=extra["name"]),
                                 user_agent))
    return batches


def _base_form(form: str) -> str:
    """'6-K/A' -> '6-K'."""
    return form.upper().split("/")[0]


def detect_forms(batches: Iterable[dict]) -> tuple[tuple[str, ...], str]:
    """Forms to fetch when the analyst gave none, and what was detected.

    An issuer that files 20-F / 40-F and never a 10-K is a foreign private
    issuer: its annual report is the 20-F / 40-F and everything else (quarterly
    results, press releases) comes as 6-K. Anything else keeps the domestic
    default.
    """
    filed = {_base_form(str(form)) for batch in batches for form in batch["form"]}
    annual = tuple(form for form in FOREIGN_ANNUAL_FORMS if form in filed)
    if not annual:
        return DEFAULT_FORMS, "emisora domestica"
    seen = "/".join(annual)
    if any(form.startswith("10-K") for form in filed):
        return DEFAULT_FORMS, (f"emisora con 10-K y {seen} en su historia, se usa el "
                               f"default domestico (--forms para incluir {seen},6-K)")
    return (*annual, FOREIGN_CURRENT_FORM), f"emisora extranjera (presenta {seen}, sin 10-K)"


def select_filings(batches: Iterable[dict], forms: Iterable[str], since: str,
                   amendments: bool = False) -> tuple[list[FilingRow], int]:
    """Matching filings and how many amendments were left out.

    Las enmiendas (``10-K/A``, ``20-F/A``...) solo se incluyen con
    ``amendments=True``; si no, se CUENTAN para avisar al analista en vez de
    omitirlas en silencio. Orden (filed, accession): no depende de EDGAR.
    """
    wanted = {f.upper() for f in forms}
    if amendments:
        wanted |= {f"{f}/A" for f in wanted if not f.endswith("/A")}
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
    return sorted(rows, key=lambda r: (r.filed, r.accession)), omitted


def list_filings(cik10: str, forms: Iterable[str], since: str,
                 user_agent: str, amendments: bool = False
                 ) -> tuple[list[FilingRow], int]:
    """Fetch the submissions index and select from it (see select_filings)."""
    return select_filings(load_submissions(cik10, user_agent), forms, since, amendments)


def has_exhibits(form: str) -> bool:
    """True for the forms whose exhibits are fetched (8-K, 6-K and their /A)."""
    return _base_form(form) in EXHIBIT_FORMS


def list_exhibits(cik_short: str, accn_flat: str, primary_doc: str,
                  user_agent: str) -> list[str]:
    """Exhibit documents of one 8-K / 6-K folder.

    8-K: the EX-99.* press release. 6-K (foreign private issuer): the quarterly
    report and the press release, often as PDF. Deterministic filter over the
    accession's index.json: every .htm / .html / .pdf that is not the primary
    wrapper, an index page nor an XBRL viewer page. These folders are small
    (wrapper + a few exhibits + plumbing), so this stays precise without HTML
    parsing.
    """
    url = SEC_FOLDER_INDEX_URL.format(cik=cik_short, accn=accn_flat)
    data = _get_json(url, user_agent)
    items = data.get("directory", {}).get("item", [])
    exhibits: list[str] = []
    primary_lower = primary_doc.lower()
    for item in items:
        name = str(item.get("name", ""))
        low = name.lower()
        if not low.endswith(EXHIBIT_SUFFIXES):
            continue
        if low == primary_lower or "index" in low:
            continue
        if re.fullmatch(r"r\d+\.htm", low):  # XBRL viewer artifact, not a doc
            continue
        exhibits.append(name)
    return exhibits


list_8k_exhibits = list_exhibits  # nombre anterior (solo 8-K); se conserva


# ---- manifest -------------------------------------------------------------

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


# ---- naming ---------------------------------------------------------------

def _safe_form(form: str) -> str:
    """'10-K/A' -> '10-K-A': una diagonal en el nombre crearia una carpeta."""
    return form.replace("/", "-")


def _accn6(accession: str) -> str:
    """Last 6 digits of the accession number (its sequence within the year)."""
    return re.sub(r"\D", "", accession)[-6:]


def _prefix(tick: str, form: str, period: str, filed: str, tag: str) -> str:
    return f"{tick}_{_safe_form(form)}_{period}_{filed}_{tag}"


class _NameRegistry:
    """Prefijo crudo determinista por filing: ``<TICKER>_<form>_<periodo>_<filed>_<accn6>``.

    El nombre sale solo del filing: no depende del orden de la corrida ni del
    manifest, asi que "ya existe" nunca atribuye el archivo de otro filing. Si
    dos accessions distintas produjeran el mismo prefijo (mismo form, periodo,
    fecha y ultimos 6 digitos), la que ya lo usa en el manifest — o la primera
    en orden (filed, accession) — lo conserva y la otra lleva la accession
    completa: nunca dos origenes -> un archivo.
    """

    def __init__(self, tick: str, manifest_rows: Iterable[dict[str, str]]) -> None:
        self._tick = tick
        self._owner: dict[str, str] = {}
        for row in manifest_rows:
            accession = row.get("accession", "")
            short = _prefix(tick, row.get("form", "").removesuffix("-EX"),
                            row.get("period", ""), row.get("filed", ""), _accn6(accession))
            if row.get("file", "").startswith((f"{short}.", f"{short}_ex_")):
                self._owner.setdefault(short.lower(), accession)

    def prefix(self, row: FilingRow) -> str:
        short = _prefix(self._tick, row.form, row.period, row.filed, _accn6(row.accession))
        if self._owner.setdefault(short.lower(), row.accession) == row.accession:
            return short
        return _prefix(self._tick, row.form, row.period, row.filed, row.accession)


class _LegacyNames:
    """Nombres de versiones anteriores que el manifest viejo permite conservar.

    Un workspace existente no se vuelve a bajar: si el manifest tiene una fila
    con la misma URL de origen, su archivo sigue en disco y ese nombre lo
    reclama UNA sola accession, se sigue usando. Un nombre compartido por
    varias accessions es el bug viejo (todas las filas apuntaban al primer
    archivo): no se confia en el y el documento se baja con el nombre nuevo.
    """

    def __init__(self, dest: Path, manifest_rows: Iterable[dict[str, str]]) -> None:
        self._dest = dest
        self._file_by_source: dict[str, str] = {}
        self._claims: dict[str, set[str]] = {}
        for row in manifest_rows:
            file, source = row.get("file", ""), row.get("source", "")
            if file and source:
                self._file_by_source[source] = file
                self._claims.setdefault(file, set()).add(row.get("accession", ""))

    def resolve(self, url: str) -> tuple[Optional[str], bool]:
        """(legacy file to keep using or None, True if the legacy name is ambiguous)."""
        file = self._file_by_source.get(url)
        if file is None:
            return None, False
        if len(self._claims[file]) > 1:
            return None, True
        return (file if (self._dest / file).is_file() else None), False


# ---- run ------------------------------------------------------------------

@dataclass
class _Run:
    """State of one download run (one ticker, one destination folder)."""

    tick: str
    cik_short: str
    dest: Path
    user_agent: str
    throttle: float
    dry_run: bool
    manifest: Path
    old_rows: list[dict[str, str]]
    names: _NameRegistry
    legacy: _LegacyNames
    new_rows: list[dict[str, str]] = field(default_factory=list)
    failures: list[str] = field(default_factory=list)
    refetched: int = 0  # documentos con nombre legacy ambiguo bajados de nuevo
    _claimed: dict[str, str] = field(default_factory=dict)
    _streak: int = 0
    _unsaved: bool = False

    def fetch_filing(self, row: FilingRow) -> None:
        """Primary document, then exhibits (8-K / 6-K), then persist the manifest."""
        prefix = self.names.prefix(row)
        ext = Path(row.primary_doc).suffix or ".htm"
        self._document(row, row.form, row.primary_doc, f"{prefix}{ext}", "")
        if has_exhibits(row.form):
            # EX-99.* press release / reporte trimestral: trae el guidance
            # cuantitativo del periodo (los transcripts NO viven en EDGAR).
            for doc in self._exhibits(row):
                self._document(row, f"{row.form}-EX", doc,
                               f"{prefix}_ex_{Path(doc).stem}{Path(doc).suffix}", " exhibit")
        self._checkpoint()

    def save(self) -> bool:
        """Merge this run's rows into the manifest on disk. False = nothing to write."""
        if self.dry_run or not (self.new_rows or self.old_rows):
            return False
        write_manifest(self.manifest, self.old_rows, self.new_rows)
        self._unsaved = False
        return True

    def _checkpoint(self) -> None:
        """Manifest incremental: un corte duro pierde, a lo mas, el filing en curso."""
        if not self._unsaved:
            return
        try:
            self.save()
        except OSError as exc:  # p. ej. antivirus con el CSV abierto: no se aborta
            print(f"[aviso] manifest no persistido aun ({type(exc).__name__}); "
                  "se reintenta en el siguiente filing")

    def _url(self, row: FilingRow, doc: str) -> str:
        return SEC_ARCHIVES_URL.format(cik=self.cik_short,
                                       accn=row.accession.replace("-", ""), doc=doc)

    def _exhibits(self, row: FilingRow) -> list[str]:
        _sleep(self.throttle)
        try:
            docs = list_exhibits(self.cik_short, row.accession.replace("-", ""),
                                 row.primary_doc, self.user_agent)
        except SecRetriesExhausted as exc:
            print(f"[x] exhibits no listados para {row.accession}: {exc}")
            self.failures.append(f"exhibits de {row.accession}")
            self._note(exhausted=True)
            return []
        except SecHttpError as exc:
            print(f"[aviso] exhibits no listados para {row.accession}: {exc}")
            docs = []
        # Un listado que responde NO reinicia la racha: solo la reinicia un
        # documento bajado. Si el indice contesta pero los documentos agotan
        # reintentos (126 s cada uno), la corrida tiene que cortarse igual; con
        # el reinicio aqui, tres fallas seguidas de un mismo 6-K no la cortaban
        # (visto contra un servidor HTTP real, issue #48).
        return docs

    def _document(self, row: FilingRow, form: str, doc: str, new_name: str,
                  label: str) -> None:
        url = self._url(row, doc)
        legacy_name, ambiguous = self.legacy.resolve(url)
        name = legacy_name or new_name
        if self._claimed.setdefault(name.lower(), url) != url:
            print(f"[x] {name}: nombre ya usado por otro documento de esta corrida; "
                  f"no se baja {url}")
            self.failures.append(name)
            return
        if self._get(url, name, label, ambiguous):
            self.new_rows.append({
                "file": name, "form": form, "filed": row.filed,
                "period": row.period, "accession": row.accession, "source": url,
            })

    def _get(self, url: str, name: str, label: str, ambiguous: bool) -> bool:
        path = self.dest / name
        if self.dry_run:
            print(f"[dry] {name}  <-  {url}")
            self.refetched += ambiguous
            return True
        if path.exists():
            print(f"[ok] ya existe: {name}")
            return True
        print(f"[..] bajando{label}: {name}")
        try:
            _download(url, path, self.user_agent)
        except SecHttpError as exc:
            print(f"[x] {name}: {exc}")
            self.failures.append(name)
            self._note(exhausted=isinstance(exc, SecRetriesExhausted))
            return False
        finally:
            _sleep(self.throttle)
        self._note(exhausted=False)
        self.refetched += ambiguous
        self._unsaved = True
        return True

    def _note(self, exhausted: bool) -> None:
        """Circuit breaker: fallas SEGUIDAS con reintentos agotados (documento o
        listado de exhibits). Solo un documento bajado reinicia la racha."""
        self._streak = self._streak + 1 if exhausted else 0
        if self._streak >= BREAKER_LIMIT:
            raise _CircuitOpen


def _resolve_forms(forms: Optional[Iterable[str]], batches: list[dict]) -> list[str]:
    """Explicit forms always win; otherwise detect the issuer type and say so."""
    explicit = list(forms or [])
    if explicit:
        return explicit
    detected, label = detect_forms(batches)
    print(f"[ok] formas detectadas en EDGAR: {label} -> {','.join(detected)}")
    return list(detected)


def _report(run: _Run, aborted: bool) -> None:
    if run.refetched:
        verb = "se bajarian" if run.dry_run else "se bajaron"
        print(f"[aviso] {run.refetched} documentos tenian en el manifest un nombre legacy "
              f"compartido por varias accessions (ambiguo): {verb} de nuevo con el nombre "
              "determinista; el archivo viejo queda en disco (nada se borra)")
    if run.failures:
        print(f"[x] {len(run.failures)} descargas fallaron (reintenta la corrida; "
              "lo ya bajado no se repite)")
    if aborted:
        print(f"[x] corrida detenida: {BREAKER_LIMIT} solicitudes seguidas agotaron sus "
              "reintentos. Lo ya bajado y el manifest quedan en disco; espera unos "
              "minutos (SEC limita la tasa por IP) y repite el mismo comando")


def fetch(ticker: str, dest: Path, forms: Optional[Iterable[str]], since: str,
          user_agent: str, dry_run: bool = False, amendments: bool = False,
          throttle: float = THROTTLE_SECONDS) -> tuple[Path, list[str]]:
    """Download filings + merge the manifest. Returns (manifest, failures).

    ``forms=None`` detects the issuer type (domestic / foreign private issuer).
    """
    tick = ticker.upper()
    cik10 = resolve_cik(ticker, user_agent)
    batches = load_submissions(cik10, user_agent, throttle)
    wanted = _resolve_forms(forms, batches)
    rows, omitted = select_filings(batches, wanted, since, amendments)
    print(f"[ok] {tick} CIK {cik10}: {len(rows)} filings "
          f"({','.join(wanted)} desde {since})")
    if omitted:
        print(f"[aviso] {omitted} enmiendas (/A) omitidas; usa --amendments "
              "para bajarlas (coverage-folders las archiva con _amended)")
    dest.mkdir(parents=True, exist_ok=True)
    manifest = dest / f"{tick}_filings_manifest.csv"
    old_rows = read_manifest(manifest)
    run = _Run(tick=tick, cik_short=str(int(cik10)), dest=dest, user_agent=user_agent,
               throttle=throttle, dry_run=dry_run, manifest=manifest, old_rows=old_rows,
               names=_NameRegistry(tick, old_rows), legacy=_LegacyNames(dest, old_rows))
    aborted = False
    try:
        for row in rows:
            run.fetch_filing(row)
    except _CircuitOpen:
        aborted = True
    finally:
        # Siempre se escribe (tambien si la red se corta o matan el proceso a
        # medias): lo ya descargado conserva su procedencia.
        if run.save():
            print(f"[ok] manifest: {manifest} ({len(run.new_rows)} filas de esta corrida)")
    _report(run, aborted)
    return manifest, run.failures


# ---- CLI ------------------------------------------------------------------

class _Tee(io.TextIOBase):
    """Text stream that copies every write to several streams (console + log)."""

    def __init__(self, *streams: TextIO) -> None:
        self._streams = streams

    def write(self, text: str) -> int:
        for stream in self._streams:
            stream.write(text)
        return len(text)

    def flush(self) -> None:
        for stream in self._streams:
            stream.flush()


@contextlib.contextmanager
def _console_log(path: Optional[str]) -> Iterator[None]:
    """Append every console line of the run to ``path`` (UTF-8) while printing it."""
    if not path:
        yield
        return
    try:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        log = open(path, "a", encoding="utf-8", buffering=1)
    except OSError as exc:
        raise SystemExit(f"[x] no se pudo abrir --log {path}: {type(exc).__name__}")
    with log:
        # Unico uso del reloj local: fechar la corrida dentro de su propio log.
        log.write(f"=== sec_fetch {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
        with contextlib.redirect_stdout(_Tee(sys.stdout, log)):
            try:
                yield
            except SystemExit as exc:
                if isinstance(exc.code, str):  # mensaje de salida (va a stderr)
                    log.write(exc.code + "\n")
                raise


def _resolve_user_agent(cli_value: Optional[str]) -> str:
    ua = cli_value or os.environ.get("SEC_EDGAR_UA", "")
    if not ua or "@" not in ua:
        raise SystemExit(
            "[x] SEC exige User-Agent con contacto (nombre + email).\n"
            "    Pasa --ua \"Nombre correo@dominio.com\" o define SEC_EDGAR_UA.")
    return ua


def _seconds(value: str) -> float:
    seconds = float(value)
    if seconds < 0:
        raise argparse.ArgumentTypeError("debe ser >= 0")
    return seconds


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Descarga filings SEC EDGAR")
    parser.add_argument("ticker")
    parser.add_argument("--forms", default=None,
                        help="lista separada por comas; sin --forms se detecta en "
                             "EDGAR: 10-K,10-Q,8-K (emisora domestica) o 20-F,6-K / "
                             "40-F,6-K (emisora extranjera). Exhibits: 8-K y 6-K")
    parser.add_argument("--since", default=DEFAULT_SINCE, help="YYYY-MM-DD")
    parser.add_argument("--dest", default=".", help="carpeta destino")
    parser.add_argument("--ua", default=None,
                        help="User-Agent 'Nombre correo' (o env SEC_EDGAR_UA)")
    parser.add_argument("--amendments", action="store_true",
                        help="incluir enmiendas (10-K/A, 10-Q/A, 8-K/A, 20-F/A, 6-K/A)")
    parser.add_argument("--throttle", type=_seconds, default=THROTTLE_SECONDS,
                        metavar="SECONDS",
                        help="pausa entre solicitudes, en segundos (default: "
                             f"{THROTTLE_SECONDS}; SEC admite hasta 10 solicitudes/s)")
    parser.add_argument("--log", default=None, metavar="PATH",
                        help="anexar cada linea de consola de la corrida a este "
                             "archivo (UTF-8), ademas de mostrarla")
    parser.add_argument("--dry-run", action="store_true",
                        help="listar sin descargar")
    return parser.parse_args(argv)


def _run_cli(args: argparse.Namespace) -> int:
    user_agent = _resolve_user_agent(args.ua)
    forms = [f.strip() for f in (args.forms or "").split(",") if f.strip()]
    try:
        _, failures = fetch(args.ticker, Path(args.dest), forms or None, args.since,
                            user_agent, dry_run=args.dry_run,
                            amendments=args.amendments, throttle=args.throttle)
    except SecHttpError as exc:
        print(f"[x] SEC: {exc}")
        return 1
    return 1 if failures else 0


def main(argv: list[str]) -> int:
    args = _parse_args(argv[1:])
    with _console_log(args.log):
        return _run_cli(args)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
