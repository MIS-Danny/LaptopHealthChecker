$ErrorActionPreference = 'Stop'

$base = 'https://raw.githubusercontent.com/MIS-Danny/LaptopHealthChecker/main'
$folder = Join-Path $env:TEMP 'LaptopHealthChecker'

try {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null

    Write-Host 'Downloading Laptop Health Checker...' -ForegroundColor Cyan

    foreach ($file in @('Dashboard.ps1', 'Worker.ps1')) {
        $url = "$base/$file"
        $destination = Join-Path $folder $file

        Invoke-WebRequest -Uri $url -OutFile $destination
    }

    Write-Host 'Download complete. Starting dashboard...' -ForegroundColor Green

    & "$folder\Dashboard.ps1"
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host 'Press Enter to exit'
}
