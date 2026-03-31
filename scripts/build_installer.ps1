Param(
    [string]$VenvPython = ".\.venv\Scripts\python.exe"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $VenvPython)) {
    throw "Python executable not found in venv: $VenvPython"
}

& $VenvPython -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('cx_Freeze') else 1)"
if ($LASTEXITCODE -ne 0) {
    throw "cx_Freeze is not installed in venv. Install it with: .\\.venv\\Scripts\\python.exe -m pip install cx_Freeze"
}

$distDir = Join-Path (Get-Location) "dist_installer"
if (-not (Test-Path $distDir)) {
    New-Item -ItemType Directory -Path $distDir | Out-Null
}
Get-ChildItem -Path $distDir -Filter *.msi -ErrorAction SilentlyContinue | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
}

# Keep temporary files inside project workspace for predictable cleanup.
# Use a unique folder per run to avoid stale lock issues.
$tmpRoot = Join-Path (Get-Location) ".tmp_cx_run"
if (-not (Test-Path $tmpRoot)) {
    New-Item -ItemType Directory -Path $tmpRoot | Out-Null
}
$tmpDir = Join-Path $tmpRoot ([Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmpDir | Out-Null

$env:TEMP = $tmpDir
$env:TMP = $tmpDir

& $VenvPython ".\scripts\cx_setup.py" bdist_msi --dist-dir "$distDir"

$msi = Get-ChildItem -Path $distDir -Filter *.msi | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $msi) {
    throw "MSI file was not generated."
}

# Rename installer file to stable English slug.
$renameScript = @'
from pathlib import Path
import re
import sys

src = Path(sys.argv[1])
app = "SeasonalPrice"
match = re.search(r"(\d+\.\d+\.\d+)", src.stem)
version = match.group(1) if match else "0.1.0"
target = src.with_name(f"{app}-{version}-win64{src.suffix}")
if target.exists():
    target.unlink()
src.replace(target)
print(target.resolve())
'@

$renamedMsi = $renameScript | & $VenvPython - "$($msi.FullName)"
Write-Output "Installer created: $renamedMsi"

# Best-effort cleanup of the per-run temp folder.
Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
