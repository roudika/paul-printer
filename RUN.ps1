# --- CONFIGURATION ---
$PrinterIP   = "192.168.1.50"
$PrinterName = "Sharp Hall Printer [Ladenburg]"
$DriverName  = "SHARP MX-3061 PCL6"

# Detect the folder where THIS script is currently saved
$CurrentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Build the relative path to the INF file
$InfPath = Join-Path $CurrentDir "SHARP-MX3061\English\PCL6\64bit\su2emenu.inf"

# Verify the file exists before starting
if (-not (Test-Path $InfPath)) {
    Write-Error "Driver file not found at: $InfPath"
    exit
}

Write-Host "Installing driver from: $InfPath" -ForegroundColor Cyan

# 1. Add the driver to the Windows Driver Store
pnputil.exe /add-driver "$InfPath" /install

# 2. Install the driver into the Print Subsystem
Add-PrinterDriver -Name $DriverName

# 3. Create the Network Port
Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP

# 4. Install the Printer
Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"

Write-Host "Installation Complete!" -ForegroundColor Green