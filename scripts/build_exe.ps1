Param(
    [string]$VenvPython = ".\.venv\Scripts\python.exe"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $VenvPython)) {
    throw "Python executable not found in venv: $VenvPython"
}

& $VenvPython -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('PyInstaller') else 1)"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller is not installed in venv. Install it with: .\\.venv\\Scripts\\python.exe -m pip install pyinstaller"
}

# Build GUI-only executable with a stable ASCII name first.
& $VenvPython -m PyInstaller `
    --clean `
    --noconfirm `
    --name SeasonalPrice `
    --onedir `
    --windowed `
    --paths src `
    src\seasonal_price\presentation\gui.py

Write-Output "Build completed."
Write-Output "Run: dist\\SeasonalPrice\\SeasonalPrice.exe"
