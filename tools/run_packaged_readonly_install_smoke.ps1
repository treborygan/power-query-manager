$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$RepositoryRoot = Split-Path -Parent $PSScriptRoot
$SourcePackage = Join-Path $RepositoryRoot 'dist\PowerQueryManager'
$TempRoot = [IO.Path]::GetFullPath($env:TEMP)
$TestRoot = Join-Path $TempRoot ('PowerQueryManager-readonly-' + [guid]::NewGuid().ToString('N'))
$PackageCopy = Join-Path $TestRoot 'PowerQueryManager'
$Identity = (& whoami).Trim()

if (-not (Test-Path -LiteralPath (Join-Path $SourcePackage 'PowerQueryManager.exe'))) {
    throw "Published GUI package is missing: $SourcePackage"
}
if (-not $TestRoot.StartsWith($TempRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing test path outside the temporary directory: $TestRoot"
}

New-Item -ItemType Directory -Path $TestRoot | Out-Null
try {
    Copy-Item -LiteralPath $SourcePackage -Destination $PackageCopy -Recurse
    $Grant = '{0}:(RX)' -f $Identity
    & icacls.exe $PackageCopy /inheritance:r /T /C /grant:r $Grant | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Could not set a read/execute-only ACL on the disposable package (icacls exit $LASTEXITCODE)."
    }

    $WriteProbe = Join-Path $PackageCopy '.write-probe'
    $WriteWasDenied = $false
    try {
        [IO.File]::WriteAllText($WriteProbe, 'disposable')
    } catch [UnauthorizedAccessException] {
        $WriteWasDenied = $true
    } catch [IO.IOException] {
        $WriteWasDenied = $true
    }
    if (-not $WriteWasDenied) {
        if (Test-Path -LiteralPath $WriteProbe) {
            Remove-Item -LiteralPath $WriteProbe -Force
        }
        throw 'The install-directory write probe unexpectedly succeeded; refusing to call this a read-only test.'
    }

    $Python = Get-Command python -ErrorAction Stop
    $SmokeRunner = Join-Path $PSScriptRoot 'run_packaged_gui_import_smoke.py'
    $ExeCopy = Join-Path $PackageCopy 'PowerQueryManager.exe'
    & $Python.Source -B $SmokeRunner --exe $ExeCopy
    if ($LASTEXITCODE -ne 0) {
        throw "Packaged GUI import acceptance failed with exit $LASTEXITCODE."
    }
} finally {
    if (Test-Path -LiteralPath $TestRoot) {
        & icacls.exe $TestRoot /reset /T /C /Q | Out-Null
        Remove-Item -LiteralPath $TestRoot -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $TestRoot) {
            Write-Warning "Disposable read-only package copy remains at $TestRoot"
        }
    }
}
