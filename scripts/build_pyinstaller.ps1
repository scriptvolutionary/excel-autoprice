Param(
    [string]$VenvPython = ".\\.venv\\Scripts\\python.exe",
    [switch]$Clean,
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

& $pythonExe -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('PyInstaller') else 1)"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller is not installed for: $pythonExe"
}

try {
    Push-Location $rootPath

    $env:PYTHONDONTWRITEBYTECODE = "1"

    $legacyDir = Join-Path $rootPath "dist\\SeasonalPrice"
    if (Test-Path -LiteralPath $legacyDir) {
        Remove-Item -LiteralPath $legacyDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    $onefileExe = Join-Path $rootPath "dist\\SeasonalPrice.exe"
    if (Test-Path -LiteralPath $onefileExe) {
        Remove-Item -LiteralPath $onefileExe -Force -ErrorAction SilentlyContinue
    }

    $pyInstallerArgs = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--name", "SeasonalPrice",
        "--onefile",
        "--windowed",
        "--paths", "src",
        "src\\seasonal_price\\presentation\\gui.py"
    )
    if ($Clean) {
        $pyInstallerArgs = @("-m", "PyInstaller", "--clean") + $pyInstallerArgs[2..($pyInstallerArgs.Length - 1)]
    }

    & $pythonExe @pyInstallerArgs

    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed with exit code $LASTEXITCODE."
    }

    Write-Output "Build completed."
    if ($Clean) {
        Write-Output "Mode: clean build"
    }
    else {
        Write-Output "Mode: fast incremental build"
    }
    Write-Output "Run: dist\\SeasonalPrice.exe"
}
finally {
    Pop-Location
    if (-not $SkipPostCleanup -and (Test-Path -LiteralPath $cleanupScript)) {
        & $cleanupScript -Root $rootPath -IncludeBuildArtifacts
    }
}
