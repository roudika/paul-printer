# --- CONFIGURATION ---
$ZipUrl      = "https://github.com/roudika/paul-printer/archive/refs/tags/v1.0.zip"
$PrinterIP   = "192.168.1.50"
$PrinterName = "Sharp Hall Printer [Ladenburg]"
$DriverName  = "SHARP MX-3061 PCL6"

# --- WORKSPACE ---
$TempDir = "$env:TEMP\SharpPrinter"
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir | Out-Null

# 1. Download the main GitHub ZIP
Write-Host "Downloading from GitHub..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\main.zip"

# 2. Extract the main ZIP (This creates the 'paul-printer-1.0' folder)
Write-Host "Extracting GitHub folder..." -ForegroundColor Cyan
Expand-Archive -Path "$TempDir\main.zip" -DestinationPath $TempDir -Force

# 3. Find and extract the internal SHARP-MX3061.zip
$InternalZip = Get-ChildItem -Path $TempDir -Filter "SHARP-MX3061.zip" -Recurse | Select-Object -First 1
if ($null -eq $InternalZip) {
    Write-Error "Could not find SHARP-MX3061.zip inside the download."
    exit
}

Write-Host "Extracting internal driver ZIP..." -ForegroundColor Cyan
$DriverFilesPath = "$TempDir\DriverUnzipped"
Expand-Archive -Path $InternalZip.FullName -DestinationPath $DriverFilesPath -Force

# 4. Find the .inf file anywhere inside that new folder
$InfFile = Get-ChildItem -Path $DriverFilesPath -Filter "su2emenu.inf" -Recurse | Select-Object -First 1
if ($null -eq $InfFile) {
    Write-Error "Could not find su2emenu.inf inside the driver ZIP."
    exit
}

$FullInfPath = $InfFile.FullName
Write-Host "Found INF at: $FullInfPath" -ForegroundColor Green

# 5. Final Installation Commands
Write-Host "Installing driver to Windows Store..." -ForegroundColor Yellow
pnputil.exe /add-driver "$FullInfPath" /install

Write-Host "Adding Driver to Print Subsystem..." -ForegroundColor Yellow
Add-PrinterDriver -Name $DriverName

Write-Host "Creating Port and Printer..." -ForegroundColor Yellow
Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"

Write-Host "Success! Printer installed." -ForegroundColor Green
