# --- CONFIGURATION ---
$ZipUrl      = "https://github.com/roudika/paul-printer/archive/refs/tags/v1.0.zip"
$PrinterIP   = "192.168.1.50"
$PrinterName = "Sharp Hall Printer [Ladenburg]"
$DriverName  = "SHARP MX-3061 PCL6"

# --- WORKSPACE ---
$TempDir = "$env:TEMP\SharpPrinter"
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir | Out-Null

# 1. Download and Extract
Write-Host "Downloading and preparing files..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\main.zip"
Expand-Archive -Path "$TempDir\main.zip" -DestinationPath $TempDir -Force

$InternalZip = Get-ChildItem -Path $TempDir -Filter "SHARP-MX3061.zip" -Recurse | Select-Object -First 1
$DriverFilesPath = "$TempDir\DriverUnzipped"
Expand-Archive -Path $InternalZip.FullName -DestinationPath $DriverFilesPath -Force
$InfFile = Get-ChildItem -Path $DriverFilesPath -Filter "su2emenu.inf" -Recurse | Select-Object -First 1

# 2. USER PROMPT FOR CLEANUP
$UserResponse = Read-Host "Do you want to DELETE existing printer/port for a clean install? (y/n)"

if ($UserResponse -eq 'y') {
    Write-Host "Cleaning up existing configuration..." -ForegroundColor Yellow
    
    # Remove Printer (Must happen before port removal)
    if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing printer: $PrinterName" -ForegroundColor Gray
        Remove-Printer -Name $PrinterName
    }

    # Remove Port
    if (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing port: IP_$PrinterIP" -ForegroundColor Gray
        Remove-PrinterPort -Name "IP_$PrinterIP"
    }
} else {
    Write-Host "Skipping cleanup. Proceeding with standard installation..." -ForegroundColor Gray
}

# 3. INSTALLATION
Write-Host "Installing driver to Windows Store..." -ForegroundColor Yellow
pnputil.exe /add-driver "$($InfFile.FullName)" /install

Write-Host "Adding Driver to Print Subsystem..." -ForegroundColor Yellow
Add-PrinterDriver -Name $DriverName

# Add Port (Only if it doesn't exist)
if (-not (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
}

# Add Printer
if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"
    Write-Host "Success! Printer installed." -ForegroundColor Green
    
    # Optional: Set as Default
    $SetDefault = Read-Host "Set $PrinterName as the default printer? (y/n)"
    if ($SetDefault -eq 'y') {
        (Get-WmiObject -Query "Select * from Win32_Printer Where Name = '$PrinterName'").SetDefaultPrinter()
        Write-Host "Set as default printer." -ForegroundColor Green
    }
} else {
    Write-Host "Printer already exists. No changes made." -ForegroundColor Cyan
}

# Clean up temp files
Remove-Item $TempDir -Recurse -Force
