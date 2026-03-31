Param(
    [string]$VenvPython = ".\\.venv\\Scripts\\python.exe",
    [switch]$Clean,
    [switch]$SkipPostCleanup
)

$ErrorActionPreference = "Stop"
$rootPath = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$targetScript = Join-Path $rootPath "scripts\\build_pyinstaller.ps1"

if (-not (Test-Path -LiteralPath $targetScript)) {
    throw "Build script not found: $targetScript"
}

$args = @{
    VenvPython = $VenvPython
}
if ($SkipPostCleanup) {
    $args.SkipPostCleanup = $true
}
if ($Clean) {
    $args.Clean = $true
}

& $targetScript @args
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
