$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$App = Join-Path $Root 'app'
if ($env:OS -ne 'Windows_NT') { throw 'Windows required.' }
$Build = Join-Path $App 'build.ps1'
if (-not (Test-Path $Build)) { throw "Build script missing: $Build" }
Write-Host 'Running project build (dependencies, unit tests, PyInstaller onedir).'
& $Build
if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { throw "Build failed ($LASTEXITCODE)" }
$Packager = Join-Path $Root 'tools\package_release.py'
& python $Packager
if ($LASTEXITCODE -ne 0) { throw "Release package staging failed ($LASTEXITCODE)" }
$PackageVerifier = Join-Path $Root 'tools\verify_release_package.py'
& python $PackageVerifier
if ($LASTEXITCODE -ne 0) { throw "Release package verification failed ($LASTEXITCODE)" }
$Exe = Join-Path $Root 'dist\PowerQueryManager\PowerQueryManager.exe'
$Cli = Join-Path $Root 'dist\PowerQueryManager\PowerQueryManagerCLI.exe'
$Internal = Join-Path $Root 'dist\PowerQueryManager\_internal'
if (-not (Test-Path $Exe)) { throw "Launcher missing: $Exe" }
if (-not (Test-Path $Cli)) { throw "CLI planner missing: $Cli" }
if (-not (Test-Path $Internal)) { throw "onedir runtime missing: $Internal" }
Write-Host "Built EXE: $Exe"
Write-Host "Built CLI: $Cli"
Write-Host "SHA256: $((Get-FileHash $Exe -Algorithm SHA256).Hash)"
& $Cli --help
if ($LASTEXITCODE -ne 0) { throw "CLI startup smoke failed ($LASTEXITCODE)" }
Write-Host 'Next step: smoke startup requires visible desktop. No test workbook will be opened automatically.'
$Proceed = Read-Host 'Launch the packaged application for manual visual smoke? Type YES'
if ($Proceed -ne 'YES') { Write-Host 'Not launched; packaged startup UNVERIFIED.'; exit 0 }
$Process = Start-Process -FilePath $Exe -WorkingDirectory (Split-Path $Exe) -PassThru
try {
    Write-Host "Launched PID $($Process.Id). Verify UI, version, resize, source discovery, close app normally."
    $Result = Read-Host 'After checking and closing window: enter PASS/FAIL/BLOCKED'
    Write-Host "Manual startup reported: $Result. Capture supporting screenshot/log in evidence."
    if ($Result -ne 'PASS') { throw 'Packaged startup not marked as passing.' }
} finally {
    try {
        $Process.Refresh()
        if (-not $Process.HasExited) {
            Write-Warning 'The launched test process is still open; close it manually. This script will not terminate other Excel/Office processes.'
        }
    } catch {}
}
Write-Host 'Do separate disposable workbook COM tests from docs/QA_WINDOWS_MATRIX.md before release.'
