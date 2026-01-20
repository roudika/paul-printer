# --- CONFIGURATION ---
$ZipUrl      = "https://github.com/roudika/paul-printer/archive/refs/tags/v1.0.zip"
$PrinterIP   = "192.168.1.50"
$PrinterName = "Sharp Hall Printer [Ladenburg]"
$DriverName  = "SHARP MX-3061 PCL6"

# This matches the internal path of your ZIP structure
$InfRelativePath = "paul-printer-1.0\SHARP-MX3061\English\PCL6\64bit\su2emenu.inf"

# --- PREPARATION ---
$TempDir = "$env:TEMP\SharpPrinter"
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir | Out-Null

# 1. Download the ZIP from GitHub
Write-Host "Downloading drivers from GitHub..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\drivers.zip"

# 2. Extract the ZIP
Write-Host "Extracting files..." -ForegroundColor Cyan
Expand-Archive -Path "$TempDir\drivers.zip" -DestinationPath $TempDir

# 3. Locate the INF file
$FullInfPath = Join-Path $TempDir $InfRelativePath

if (-not (Test-Path $FullInfPath)) {
    Write-Error "Could not find INF at: $FullInfPath. Check your ZIP folder structure."
    return
}

# 4. Windows Installation Commands
Write-Host "Installing driver and creating printer..." -ForegroundColor Yellow
pnputil.exe /add-driver "$FullInfPath" /install
Add-PrinterDriver -Name $DriverName
Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"

Write-Host "Success! '$PrinterName' is installed." -ForegroundColor Green
