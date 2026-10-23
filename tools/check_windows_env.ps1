$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Write-Host 'Power Query Manager Windows prerequisites (read-only check)'
Write-Host "OS: $([System.Environment]::OSVersion.VersionString)"
if ($env:OS -ne 'Windows_NT') { throw 'Windows required for actual Excel COM tests and build.' }
$Commands = @('py','python','powershell')
foreach ($name in $Commands) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($cmd) { Write-Host "$name : $($cmd.Source)" } else { Write-Host "$name : unavailable" }
}
$ExcelClass = Get-Item 'Registry::HKEY_CLASSES_ROOT\Excel.Application' -ErrorAction SilentlyContinue
if ($ExcelClass) { Write-Host 'Excel.Application ProgID: registered (not yet verified to launch).' }
else { Write-Warning 'Excel.Application ProgID: not found. Install/activate desktop Excel or confirm policy.' }
Write-Host 'This script does not open Excel, modify a workbook, or install any software.'
