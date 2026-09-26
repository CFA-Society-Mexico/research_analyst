"""FRED -> series macro historicas a macro/series/ (cadena macro del plugin).

Baja las series observadas (tasas, CPI, PIB, FX) para que la tab Macro del
modelo muestre HISTORICOS con fuente — no placeholders — y /update-macro
proponga macro-view.yaml desde datos, no desde memoria.

API key (gratuita, registro en https://fredaccount.stlouisfed.org/apikeys):
cascada --api-key -> env FRED_API_KEY -> macro/fred.key (decision del dueno
del plugin: archivo visible en el workspace) -> si falta, mensaje instruyendo
pegarla en el chat para que el agente la guarde en macro/fred.key.
La key JAMAS se imprime.

Uso:
    python tools/fred_fetch.py --dest <raiz>/macro/series
    python tools/fred_fetch.py --dest macro/series --series DGS10,DEXMXUS

Series extra: <raiz>/macro/fred-series.txt (un ID por linea; '#' comenta).
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
from pathlib import Path
from typing import Optional

API_URL = ("https://api.stlouisfed.org/fred/series/observations"
           "?series_id={sid}&api_key={key}&file_type=json")
META_URL = ("https://api.stlouisfed.org/fred/series"
            "?series_id={sid}&api_key={key}&file_type=json")

DEFAULT_SERIES = {
    "DGS10": "UST 10Y (riesgo libre USD)",
    "FEDFUNDS": "Fed Funds effective",
    "CPIAUCSL": "CPI US (indice)",
    "GDPC1": "PIB real US",
    "DEXMXUS": "USDMXN spot",
}

KEY_MSG = """[x] Falta la API key de FRED.
    Es GRATUITA: registrate en https://fredaccount.stlouisfed.org/apikeys
    (2 minutos, sin tarjeta). Luego, cualquiera de estas rutas:
    1) Pega la key en el chat y el agente la guarda en macro/fred.key.
    2) Crea el archivo macro/fred.key con la key como unico contenido.
    3) --api-key <key> o variable de entorno FRED_API_KEY."""

BLOCKED_MSG = """[x] Sin acceso de red a FRED ({err}).
    El entorno bloquea api.stlouisfed.org (proxy con allowlist - tipico en
    Claude Cowork). Opciones: 1) permitir api.stlouisfed.org en el allowlist
    del entorno; 2) correr este comando en una maquina con red y copiar
    macro/series/ al workspace."""


BAD_KEY_MSG = """[x] FRED rechazo la API key (HTTP 400).
    Verifica macro/fred.key (o --api-key / FRED_API_KEY). Registro gratuito:
    https://fredaccount.stlouisfed.org/apikeys"""

RETRY_CODES = (429, 500, 502, 503, 504)
RETRIES = 3


class FredSeriesError(Exception):
    """Falla de UNA serie (ID inexistente, 404, 5xx persistente).

    No es SystemExit a proposito: una serie mal escrita en fred-series.txt se
    reporta y el resto del lote se baja igual. Key invalida y red bloqueada
    si detienen todo (afectan a todas las series).
    """


def _redact(text: str, key: str) -> str:
    return text.replace(key, "***") if key else text


def _error_message(exc: urllib.error.HTTPError) -> str:
    try:
        body = json.loads(exc.read().decode("utf-8"))
        return str(body.get("error_message", exc.reason))
    except Exception:  # noqa: BLE001 - cuerpo no JSON: basta el reason
        return str(exc.reason)


def _get_json(url: str, key: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "research_analyst fred_fetch"})
    for attempt in range(RETRIES + 1):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code in (403, 407):
                raise SystemExit(BLOCKED_MSG.format(err=f"HTTP {exc.code}"))
            if exc.code in RETRY_CODES and attempt < RETRIES:
                time.sleep(2 ** attempt)
                continue
            message = _redact(_error_message(exc), key)
            if exc.code == 400 and "api_key" in message:
                raise SystemExit(BAD_KEY_MSG)
            raise FredSeriesError(f"HTTP {exc.code}: {message}") from exc
        except urllib.error.URLError as exc:
            raise SystemExit(BLOCKED_MSG.format(err=str(exc.reason)))
        except (TimeoutError, ConnectionError) as exc:
            if attempt < RETRIES:
                time.sleep(2 ** attempt)
                continue
            raise FredSeriesError(f"sin respuesta ({type(exc).__name__})") from exc
    raise FredSeriesError("reintentos agotados")  # pragma: no cover


def resolve_key(cli_key: Optional[str], dest: Path,
                key_file: Optional[str]) -> str:
    if cli_key:
        return cli_key.strip()
    env = os.environ.get("FRED_API_KEY", "").strip()
    if env:
        return env
    candidates = []
    if key_file:
        candidates.append(Path(key_file))
    candidates.append(dest.parent / "fred.key")   # macro/fred.key (dest=macro/series)
    candidates.append(dest / "fred.key")
    for path in candidates:
        try:
            content = path.read_text(encoding="utf-8").strip()
        except OSError:
            continue
        if content:
            return content.splitlines()[0].strip()
    raise SystemExit(KEY_MSG)


def series_list(dest: Path, cli_series: Optional[str]) -> dict[str, str]:
    if cli_series:
        return {s.strip().upper(): s.strip().upper()
                for s in cli_series.split(",") if s.strip()}
    out = dict(DEFAULT_SERIES)
    extra = dest.parent / "fred-series.txt"
    try:
        for line in extra.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                out.setdefault(line.upper(), line.upper())
    except OSError:
        pass
    return out


MANIFEST_FIELDS = ["series", "title", "units", "frequency", "last_observation",
                   "last_value", "source", "tag"]


def write_manifest(path: Path, new_rows: list[dict[str, str]]) -> None:
    """MERGE por serie: bajar solo DGS10 con --series actualiza su fila y
    conserva la procedencia de las demas series que siguen en disco."""
    merged: dict[str, dict[str, str]] = {}
    if path.exists():
        with path.open(newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                merged[row.get("series", "")] = row
    for row in new_rows:
        merged[row["series"]] = row
    fields = list(MANIFEST_FIELDS)
    for row in merged.values():
        fields.extend(k for k in row if k not in fields)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        writer.writerows(merged[s] for s in sorted(merged))


def fetch(dest: Path, key: str, wanted: dict[str, str]) -> list[str]:
    """Baja cada serie; una serie que falla se reporta y el lote sigue.
    Devuelve los IDs que fallaron."""
    dest.mkdir(parents=True, exist_ok=True)
    manifest_rows = []
    failed: list[str] = []
    for sid, title in wanted.items():
        try:
            data = _get_json(API_URL.format(sid=sid, key=key), key)
            meta = _get_json(META_URL.format(sid=sid, key=key), key)
        except FredSeriesError as exc:
            print(f"[x] {sid}: {exc} -- omitida, revisa el ID")
            failed.append(sid)
            continue
        obs = [(o["date"], o["value"]) for o in data.get("observations", [])
               if o.get("value") not in (".", "", None)]
        if not obs:
            print(f"[!] {sid}: sin observaciones -- omitida")
            continue
        info = (meta.get("seriess") or [{}])[0]
        out = dest / f"{sid}.csv"
        with out.open("w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            writer.writerow(["date", "value"])
            writer.writerows(obs)
        manifest_rows.append({
            "series": sid,
            "title": info.get("title", title),
            "units": info.get("units", ""),
            "frequency": info.get("frequency", ""),
            "last_observation": obs[-1][0],
            "last_value": obs[-1][1],
            "source": "FRED (Federal Reserve Bank of St. Louis)",
            "tag": "observado",
        })
        print(f"[ok] {sid}: {len(obs)} observaciones "
              f"({obs[0][0]} .. {obs[-1][0]}) -> {out.name}")
    if manifest_rows:
        mpath = dest / "manifest.csv"
        write_manifest(mpath, manifest_rows)
        print(f"[ok] manifest: {mpath}")
    if failed:
        print(f"[x] {len(failed)} series fallaron: {', '.join(failed)}")
    print("[i] siguiente paso: /update-macro propone macro-view.yaml desde")
    print("    estas series (gate por campo); la tab Macro del modelo las lee.")
    return failed


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="FRED -> macro/series/*.csv")
    parser.add_argument("--dest", default="macro/series",
                        help="carpeta destino (default: macro/series)")
    parser.add_argument("--api-key", default=None, help="key FRED (o env/archivo)")
    parser.add_argument("--key-file", default=None,
                        help="ruta a archivo con la key (default: macro/fred.key)")
    parser.add_argument("--series", default=None,
                        help="IDs separados por coma (default: set estandar + fred-series.txt)")
    args = parser.parse_args(argv[1:])
    dest = Path(args.dest)
    key = resolve_key(args.api_key, dest, args.key_file)
    failed = fetch(dest, key, series_list(dest, args.series))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
