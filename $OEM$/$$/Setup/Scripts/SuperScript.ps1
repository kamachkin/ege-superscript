# SuperScript Launcher - finds and runs the real script from soft/ folder
$ErrorActionPreference = 'Continue'
$found = $false
foreach ($drive in Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.Root -and (Test-Path $_.Root) }) {
    $scriptPath = Join-Path $drive.Root 'soft\superscript_modified.ps1'
    if (Test-Path $scriptPath) {
        Write-Host "Found script: $scriptPath" -ForegroundColor Green
        Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs -Wait
        $found = $true
        break
    }
}
if (-not $found) {
    Write-Host "ERROR: Could not find soft\superscript_modified.ps1 on any drive!" -ForegroundColor Red
    Write-Host "Please ensure the USB drive with 'soft' folder is connected." -ForegroundColor Yellow
    Start-Sleep -Seconds 30
}