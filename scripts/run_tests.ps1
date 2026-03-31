Param(
    [string]$VenvPython = ".\\.venv\\Scripts\\python.exe",
    [switch]$SkipPostCleanup
)

$ErrorActionPreference = "Stop"
$rootPath = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$cleanupScript = Join-Path $rootPath "scripts\\cleanup_workspace.ps1"

function Resolve-PythonExe {
    Param(
        [string]$Candidate,
        [string]$RootPath
    )

    $candidatePath = $Candidate
    if (-not [System.IO.Path]::IsPathRooted($candidatePath)) {
        $candidatePath = Join-Path $RootPath $candidatePath
    }
    if (Test-Path -LiteralPath $candidatePath) {
        return (Resolve-Path -LiteralPath $candidatePath).Path
    }

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        Write-Host "Using python from PATH: $($pythonCmd.Source)"
        return $pythonCmd.Source
    }

    throw "Python executable not found: $Candidate"
}

$pythonExe = Resolve-PythonExe -Candidate $VenvPython -RootPath $rootPath

try {
    Push-Location $rootPath

    $env:PYTHONPATH = "src"
    $env:PYTHONDONTWRITEBYTECODE = "1"

    & $pythonExe -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) {
        throw "Tests failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
    if (-not $SkipPostCleanup -and (Test-Path -LiteralPath $cleanupScript)) {
        & $cleanupScript -Root $rootPath
    }
}
