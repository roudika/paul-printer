# --- CONFIGURATION ---
$ZipUrl      = "https://github.com/roudika/paul-printer/archive/refs/tags/v1.0.zip"
$PrinterIP   = "192.168.1.50"
$PrinterName = "Sharp Hall Printer [Ladenburg]"
$DriverName  = "SHARP MX-3061 PCL6"

# --- WORKSPACE ---
$TempDir = "$env:TEMP\SharpPrinter"
$DownloadPath = "$TempDir\github_download.zip"

if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir | Out-Null

# 1. Download the GitHub Release ZIP
Write-Host "Downloading from GitHub..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $ZipUrl -OutFile $DownloadPath

# 2. Extract the GitHub Archive (the paul-printer-1.0 folder)
Write-Host "Extracting GitHub Archive..." -ForegroundColor Cyan
Expand-Archive -Path $DownloadPath -DestinationPath $TempDir -Force

# 3. Find and Extract the INTERNAL Driver ZIP (SHARP-MX3061.zip)
$InternalZip = Get-ChildItem -Path $TempDir -Filter "SHARP-MX3061.zip" -Recurse | Select-Object -First 1
if ($null -eq $InternalZip) {
    Write-Error "Could not find SHARP-MX3061.zip inside the GitHub download."
    exit
}

Write-Host "Extracting internal driver files..." -ForegroundColor Cyan
$DriverExtractPath = "$TempDir\DriverFiles"
Expand-Archive -Path $InternalZip.FullName -DestinationPath $DriverExtractPath -Force

# 4. Find the .inf file inside the newly extracted driver folder
$InfFile = Get-ChildItem -Path $DriverExtractPath -Filter "su2emenu.inf" -Recurse | Select-Object -First 1
if ($null -eq $InfFile) {
    Write-Error "Still could not find su2emenu.inf. Check the contents of SHARP-MX3061.zip"
    exit
}

# 5. Final Installation
Write-Host "Installing driver: $($InfFile.FullName)" -ForegroundColor Yellow
pnputil.exe /add-driver "$($InfFile.FullName)" /install
Add-PrinterDriver -Name $DriverName
Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"

Write-Host "Installation Finished Successfully!" -ForegroundColor Green
