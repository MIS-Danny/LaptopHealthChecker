
$ErrorActionPreference = 'Stop'

$base = 'https://raw.githubusercontent.com/MIS-Danny/LaptopHealthChecker/main'
$folder = Join-Path $env:TEMP 'LaptopHealthChecker'

try {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null

    Write-Host 'Downloading Laptop Health Checker...' -ForegroundColor Cyan

    foreach ($file in @('Dashboard.ps1', 'Worker.ps1')) {
        $stamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $url = "${base}/${file}?v=$stamp"
        $destination = Join-Path $folder $file

        Invoke-WebRequest -Uri $url -OutFile $destination
    }

    Write-Host 'Download complete. Starting dashboard...' -ForegroundColor Green

    $dashboard = Join-Path $folder 'Dashboard.ps1'

    $process = Start-Process `
        -FilePath 'powershell.exe' `
        -ArgumentList @(
            '-NoProfile'
            '-ExecutionPolicy', 'Bypass'
            '-File', "`"$dashboard`""
        ) `
        -PassThru

    if (-not $process) {
        throw 'Unable to start the dashboard.'
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
}
