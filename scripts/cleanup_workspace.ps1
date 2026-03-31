Param(
    [string]$Root = ".",
    [switch]$IncludeBuildArtifacts
)

$ErrorActionPreference = "Stop"

$rootPathRaw = (Resolve-Path -LiteralPath $Root).Path
$rootPath = [System.IO.Path]::GetFullPath($rootPathRaw).TrimEnd('\\')
$rootWithSep = $rootPath + [System.IO.Path]::DirectorySeparatorChar

if (-not (Test-Path -LiteralPath (Join-Path $rootPath ".git"))) {
    throw "Cleanup aborted: '$rootPath' is not a project root (missing .git)."
}

function Remove-SafePath {
    Param(
        [string]$PathToRemove,
        [string]$RootPath,
        [string]$RootWithSep
    )

    if (-not (Test-Path -LiteralPath $PathToRemove)) {
        return
    }

    $resolved = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PathToRemove).Path).TrimEnd('\\')
    if (-not $resolved.StartsWith($RootWithSep, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe cleanup target blocked: $resolved"
    }
    if ($resolved -eq $RootPath) {
        throw "Unsafe cleanup target blocked: root folder"
    }

    Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
}

$cacheDirs = @(
    ".tmp",
    ".tmp_tests",
    ".tmp_cx",
    ".tmp_cx_run",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache"
)

foreach ($dir in $cacheDirs) {
    $target = Join-Path $rootPath $dir
    Remove-SafePath -PathToRemove $target -RootPath $rootPath -RootWithSep $rootWithSep
}

if ($IncludeBuildArtifacts) {
    Remove-SafePath -PathToRemove (Join-Path $rootPath "build") -RootPath $rootPath -RootWithSep $rootWithSep
}

$scanRoots = @("src", "tests", "scripts")
foreach ($scanRoot in $scanRoots) {
    $scanPath = Join-Path $rootPath $scanRoot
    if (-not (Test-Path -LiteralPath $scanPath)) {
        continue
    }

    Get-ChildItem -LiteralPath $scanPath -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    ForEach-Object {
        Remove-SafePath -PathToRemove $_.FullName -RootPath $rootPath -RootWithSep $rootWithSep
    }

    Get-ChildItem -LiteralPath $scanPath -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -in @('.pyc', '.pyo') } |
    ForEach-Object {
        Remove-SafePath -PathToRemove $_.FullName -RootPath $rootPath -RootWithSep $rootWithSep
    }
}

Write-Output "Cleanup completed: cache folders/files removed."

