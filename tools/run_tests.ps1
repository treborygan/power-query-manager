$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$App = Join-Path $Root 'app'
$LocalPython = Join-Path $App '.venv\Scripts\python.exe'
$Py = if (Test-Path $LocalPython) { $LocalPython } elseif (Get-Command python -ErrorAction SilentlyContinue) { 'python' } else { 'py' }
$PythonArgs = @()
if ($Py -eq 'py') {
    $found = $false
    foreach ($version in @('-3.13', '-3.12', '-3.11')) {
        & py $version --version *> $null
        if ($LASTEXITCODE -eq 0) { $PythonArgs = @($version); $found = $true; break }
    }
    if (-not $found) { throw 'Python 3.11, 3.12 or 3.13 is needed to run current v1.3.0 tests.' }
}
$env:PYTHONPATH = Join-Path $App 'src'
Push-Location $App
try {
    & $Py @PythonArgs -m pytest -q
    if ($LASTEXITCODE -ne 0) { throw "pytest failed ($LASTEXITCODE)" }
    Write-Host 'Unit tests passed. This DOES NOT verify Windows Excel COM or packaged EXE.'
} finally { Pop-Location }
