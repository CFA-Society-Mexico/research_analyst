"""Recalculo y validacion del modelo con un motor de calculo real.

Doctrina: un check que nunca se evaluo no es un check verde. openpyxl escribe
formulas sin validarlas ni calcularlas, asi que ni el audit F ni los checks S
pueden saber si Excel ACEPTA el libro ni cuanto valen los checks de formula
viva. Este tool cierra ese hueco (issue #33):

  1. copia el modelo a una carpeta temporal y lo recalcula con Excel (COM, via
     PowerShell - sin dependencias) o con LibreOffice headless;
  2. R1: compara las formulas del original contra las del libro recalculado.
     Una formula que desaparece es una formula que el motor DESCARTO (Excel en
     modo reparacion: "encontramos un problema con el contenido");
  3. R2: ninguna celda calculada queda en error (#NAME?, #VALUE!, #REF!...);
  4. lee los checks vivos de las tabs Checks y Cover ya evaluados.

Solo lectura sobre el modelo: trabaja sobre la copia. ``--in-place`` guarda los
valores calculados en el modelo (lo que antes se hacia a mano con Excel COM)
UNICAMENTE si R1 esta verde: jamas se sobreescribe un modelo con un libro al
que el motor le borro formulas.

Uso:
    python tools/recalc_validate.py <modelo.xlsx> [--engine auto|excel|libreoffice]
                                    [--in-place]

Exit codes (los mismos del audit): 0 = verde; 1 = alguna falla; 3 = no se pudo
recalcular (sin Excel ni LibreOffice): [pendiente], los checks vivos quedan
SIN evaluar y eso NUNCA es verde.

Consola ASCII-only.
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from enum import Enum
from pathlib import Path
from typing import Optional, Sequence

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_builder as B  # noqa: E402
from xlsx_builder import Finding  # noqa: E402

RECALC_TIMEOUT_SECONDS = 600
R1_NAME = "R1 libro abre sin reparacion"
R2_NAME = "R2 sin errores de calculo"

_SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)

# Abre, recalcula TODO y guarda en otra ruta. Un libro que Excel no acepta
# (formula invalida) NO abre por automatizacion -ni con CorruptLoad en modo
# reparar, verificado con Excel 16-: a mano mostraria "encontramos un problema
# con el contenido" y borraria formulas. Ese fallo de apertura ES la falla R1.
# El script cierra SOLO la instancia de Excel que el mismo creo (por PID),
# nunca una del usuario.
_EXCEL_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
Add-Type -Namespace RA -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);'
$xl = New-Object -ComObject Excel.Application
$procId = [uint32]0
[void][RA.Win]::GetWindowThreadProcessId([System.IntPtr]$xl.Hwnd, [ref]$procId)
$wb = $null
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.AutomationSecurity = 3
    $wb = $xl.Workbooks.Open($env:RA_RECALC_SRC, 0, $false)
    $xl.CalculateFullRebuild()
    $wb.SaveAs($env:RA_RECALC_OUT, 51)
    $wb.Close($false)
} finally {
    if ($wb -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    try { $xl.Quit() } catch { }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if ($procId -ne 0) {
        $own = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($own -ne $null -and -not $own.WaitForExit(5000)) { $own.Kill() }
    }
}
"""


class Engine(Enum):
    EXCEL = "excel"
    LIBREOFFICE = "libreoffice"


class RecalcError(Exception):
    """El motor no pudo abrir, recalcular o guardar el libro."""


def _excel_available() -> bool:
    if sys.platform != "win32" or shutil.which("powershell") is None:
        return False
    import winreg
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
    except OSError:
        return False
    return True


def _soffice_path() -> Optional[str]:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    return next((c for c in _SOFFICE_CANDIDATES if Path(c).is_file()), None)


def find_engine(preferred: str = "auto") -> Optional[Engine]:
    """Motor disponible. ``auto`` prefiere Excel: es el que abre el analista."""
    excel = _excel_available() if preferred in ("auto", "excel") else False
    if excel:
        return Engine.EXCEL
    if preferred in ("auto", "libreoffice") and _soffice_path():
        return Engine.LIBREOFFICE
    return None


def recalculate(source: Path, workdir: Path, engine: Engine) -> Path:
    """Recalcula una COPIA de ``source`` y devuelve la ruta del libro calculado.

    El nombre del archivo se conserva (S8 audita el nombre versionado)."""
    copy_dir, out_dir = workdir / "copia", workdir / "recalculado"
    copy_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    copy = copy_dir / source.name
    shutil.copy2(source, copy)
    out = out_dir / source.name
    if engine is Engine.EXCEL:
        env = {**os.environ, "RA_RECALC_SRC": str(copy), "RA_RECALC_OUT": str(out)}
        command = ["powershell", "-NoProfile", "-NonInteractive", "-Command", _EXCEL_SCRIPT]
    else:
        env = dict(os.environ)
        command = [str(_soffice_path()), "--headless", "--calc", "--convert-to", "xlsx",
                   "--outdir", str(out_dir), str(copy)]
    try:
        done = subprocess.run(command, env=env, capture_output=True, text=True,
                              timeout=RECALC_TIMEOUT_SECONDS, check=False)
    except subprocess.TimeoutExpired as exc:
        raise RecalcError(f"{engine.value}: sin respuesta en {RECALC_TIMEOUT_SECONDS}s") from exc
    except OSError as exc:
        raise RecalcError(f"{engine.value}: no se pudo ejecutar ({exc})") from exc
    if done.returncode != 0 or not out.is_file():
        reason = (done.stderr or done.stdout or "sin salida").strip().splitlines()
        raise RecalcError(f"{engine.value} no pudo abrir o guardar el libro "
                          f"({reason[0][:160] if reason else 'sin detalle'})")
    return out


def formula_cells(path: Path) -> set[str]:
    """Celdas con formula ('Hoja!A1'), incluidas las de arreglo."""
    cells: set[str] = set()
    wb = load_workbook(str(path), data_only=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, ArrayFormula) or (
                        isinstance(value, str) and value.startswith("=")):
                    cells.add(f"{ws.title}!{cell.coordinate}")
    return cells


def rejected_formulas(path: Path) -> list[str]:
    """Celdas cuya formula el lint F21 senala (candidatas de un libro que no abre)."""
    hits: list[str] = []
    wb = load_workbook(str(path), data_only=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                is_array = isinstance(cell.value, ArrayFormula)
                text = cell.value.text if is_array else cell.value
                problems = B.lint_formula(text, array=is_array)
                if problems:
                    hits.append(f"{ws.title}!{cell.coordinate} ({problems[0].split(':')[0]})")
    return hits


def error_cells(path: Path) -> list[str]:
    """Celdas cuyo valor calculado es un error de Excel."""
    hits: list[str] = []
    wb = load_workbook(str(path), data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip() in B._EXCEL_ERRORS:
                    hits.append(f"{ws.title}!{cell.coordinate} = {cell.value.strip()}")
    return hits


def _examples(hits: Sequence[str], limit: int = 8) -> str:
    ordered = sorted(hits)
    extra = f" (+{len(ordered) - limit} mas)" if len(ordered) > limit else ""
    return ", ".join(ordered[:limit]) + extra


def validate(model: Path, workdir: Path,
             engine_name: str = "auto") -> tuple[list[Finding], Optional[Path]]:
    """R1, R2 y checks vivos. Devuelve (hallazgos, libro recalculado o None)."""
    engine = find_engine(engine_name)
    if engine is None:
        return [Finding(
            R1_NAME, False,
            "sin Excel (COM) ni LibreOffice en esta maquina: el libro NO se recalculo y "
            "los checks vivos quedan SIN evaluar. Abre el modelo en Excel, recalcula "
            "(CalculateFullRebuild), guarda y vuelve a auditar", pending=True)], None
    try:
        recalculated = recalculate(model, workdir, engine)
    except RecalcError as exc:
        suspects = rejected_formulas(model)
        hint = ("; formulas que Excel no acepta (F21): " + _examples(suspects)
                if suspects else "; el lint F21 no encontro la formula culpable")
        return [Finding(R1_NAME, False,
                        f"{exc}: a mano abriria en modo reparacion y borraria formulas. "
                        "Checks vivos SIN evaluar" + hint)], None
    discarded = formula_cells(model) - formula_cells(recalculated)
    r1 = Finding(
        R1_NAME, not discarded,
        (f"{engine.value} descarto {len(discarded)} formulas al abrir (libro reparado): "
         + _examples(list(discarded))) if discarded
        else f"{engine.value} conserva todas las formulas del modelo")
    errors = error_cells(recalculated)
    r2 = Finding(R2_NAME, not errors,
                 (f"{len(errors)} celdas en error: " + _examples(errors)) if errors
                 else "ninguna celda calculada en error")
    return [r1, r2, B.live_checks_finding(str(recalculated))], recalculated


def save_in_place(model: Path, recalculated: Path) -> None:
    """Sustituye el modelo por su version con valores calculados (atomico)."""
    staged = model.with_name(model.name + ".part")
    shutil.copy2(recalculated, staged)
    os.replace(staged, model)


def main(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Recalcula el modelo con Excel o LibreOffice y valida que el motor "
                    "lo acepte (R1), que no haya errores (R2) y los checks vivos.")
    parser.add_argument("model", help="ruta del modelo .xlsx")
    parser.add_argument("--engine", choices=("auto", "excel", "libreoffice"), default="auto")
    parser.add_argument("--in-place", action="store_true",
                        help="guardar los valores calculados en el modelo (solo si R1 esta verde)")
    args = parser.parse_args(list(argv[1:]))
    model = Path(args.model)
    if not model.is_file():
        print(f"[x] no existe el modelo: {args.model}")
        return 2
    with tempfile.TemporaryDirectory(prefix="ra_recalc_") as tmp:
        findings, recalculated = validate(model, Path(tmp), args.engine)
        code = B._print_report(findings, label="Resumen recalculo")
        if args.in_place:
            if recalculated is not None and findings[0].ok:
                save_in_place(model, recalculated)
                print(f"[ok] valores calculados guardados en {model.name}")
            else:
                print("[x] --in-place omitido: el modelo no se toca si no se pudo "
                      "recalcular o si el motor descarto formulas")
    return code


if __name__ == "__main__":
    sys.exit(main(sys.argv))
