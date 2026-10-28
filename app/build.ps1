$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$ProjectRoot = $PSScriptRoot
Set-Location $ProjectRoot

function Invoke-Checked {
    param([string]$Program, [string[]]$Arguments)
    & $Program @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE`: $Program $($Arguments -join ' ')"
    }
}

if (-not (Test-Path '.venv\Scripts\python.exe')) {
    $PythonCommand = Get-Command py -ErrorAction SilentlyContinue
    if ($PythonCommand) {
        Invoke-Checked 'py' @('-3.11', '-m', 'venv', '.venv')
    } else {
        $PythonCommand = Get-Command python -ErrorAction SilentlyContinue
        if (-not $PythonCommand) { throw 'Python 3.11 is required. Install it, then rerun this script.' }
        Invoke-Checked 'python' @('-m', 'venv', '.venv')
    }
}
$Python = Join-Path $ProjectRoot '.venv\Scripts\python.exe'
$Version = & $Python --version
if ($LASTEXITCODE -ne 0) { throw 'Unable to run virtual-environment Python.' }
if ($Version -notmatch 'Python 3\.(11|12|13)\.') { throw "Unsupported Python version: $Version. Use Python 3.11, 3.12, or 3.13." }

Invoke-Checked $Python @('-m', 'pip', 'install', '--disable-pip-version-check', '-r', 'requirements-win.txt')
Invoke-Checked $Python @('-m', 'pip', 'install', '--no-deps', '-e', '.')
# CI diagnostic mode provides a named, short first failure instead of hanging
# for the entire job. Normal/release runs must execute the complete suite.
if ($env:PQM_CI_DIAGNOSTICS -eq '1') {
    Invoke-Checked $Python @('-m', 'pytest', '-vv', '--tb=short', '--maxfail=15', '--timeout=75', '--timeout-method=thread')
} else {
    Invoke-Checked $Python @('-m', 'pytest', '-q')
}
$PreviousPythonHashSeed = $env:PYTHONHASHSEED
$PreviousSourceDateEpoch = $env:SOURCE_DATE_EPOCH
try {
    # Stabilize hash-dependent module ordering and PE timestamps for repeatable
    # onedir packages built from the same source and pinned toolchain.
    $env:PYTHONHASHSEED = '0'
    $env:SOURCE_DATE_EPOCH = '1704067200'
    Invoke-Checked $Python @('-m', 'PyInstaller', '--noconfirm', '--clean', 'PowerQueryInstaller.spec')
} finally {
    if ($null -eq $PreviousPythonHashSeed) { Remove-Item Env:PYTHONHASHSEED -ErrorAction SilentlyContinue }
    else { $env:PYTHONHASHSEED = $PreviousPythonHashSeed }
    if ($null -eq $PreviousSourceDateEpoch) { Remove-Item Env:SOURCE_DATE_EPOCH -ErrorAction SilentlyContinue }
    else { $env:SOURCE_DATE_EPOCH = $PreviousSourceDateEpoch }
}
$Executable = Join-Path $ProjectRoot 'dist\PowerQueryManager\PowerQueryManager.exe'
$CliExecutable = Join-Path $ProjectRoot 'dist\PowerQueryManager\PowerQueryManagerCLI.exe'
if (-not (Test-Path $Executable)) { throw "Build completed without expected executable: $Executable" }
if (-not (Test-Path $CliExecutable)) { throw "Build completed without expected CLI executable: $CliExecutable" }
Write-Host "Build succeeded: $Executable and $CliExecutable"
Write-Host 'Launch it on Windows with desktop Excel installed to complete the manual COM acceptance checks.'
