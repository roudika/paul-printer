# --- CONFIGURATION ---
$ZipUrl = "https://github.com/roudika/paul-printer/archive/refs/tags/v1.0.zip"
$PrinterIP = "192.168.1.50"
$PrinterName = "Sharp 1 - Hall Printer [Ladenburg]"
$DriverName = "SHARP MX-3061 PCL6"

# --- WORKSPACE ---
$TempDir = "$env:TEMP\SharpPrinter"
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir | Out-Null

# 0. Admin Check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script MUST be run as Administrator." -ForegroundColor Red
    exit
}

# 1. Download and Extract
Write-Host "Downloading and preparing files..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\main.zip"
Expand-Archive -Path "$TempDir\main.zip" -DestinationPath $TempDir -Force

$InternalZip = Get-ChildItem -Path $TempDir -Filter "SHARP-MX3061.zip" -Recurse | Select-Object -First 1
$DriverFilesPath = "$TempDir\DriverUnzipped"
Expand-Archive -Path $InternalZip.FullName -DestinationPath $DriverFilesPath -Force
$InfFile = Get-ChildItem -Path $DriverFilesPath -Filter "su2emenu.inf" -Recurse | Select-Object -First 1

# 2. CONFLICT DETECTION & CLEANUP
$ExistingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
$ExistingPort = Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue

if ($ExistingPrinter -or $ExistingPort) {
    Write-Host "`n[!] CONFLICT DETECTED" -ForegroundColor Yellow
    if ($ExistingPrinter) { Write-Host " - Printer '$PrinterName' already exists." -ForegroundColor Gray }
    if ($ExistingPort) { Write-Host " - Port 'IP_$PrinterIP' already exists." -ForegroundColor Gray }
    
    $UserResponse = Read-Host "`nDo you want to REMOVE existing configuration for a clean install? (y/n)"
    
    if ($UserResponse -eq 'y') {
        Write-Host "Cleaning up..." -ForegroundColor Yellow
        
        # Find ALL printers using this port (even those with different names)
        $PrintersUsingPort = Get-Printer | Where-Object { $_.PortName -eq "IP_$PrinterIP" -or $_.Name -eq $PrinterName }
        
        foreach ($P in $PrintersUsingPort) {
            Write-Host " - Removing printer: $($P.Name)" -ForegroundColor Gray
            Remove-Printer -Name $P.Name
        }

        # Small pause to let the Spooler release port handles
        Start-Sleep -Seconds 1

        if (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue) { 
            Write-Host " - Removing port: IP_$PrinterIP" -ForegroundColor Gray
            try {
                Remove-PrinterPort -Name "IP_$PrinterIP" -ErrorAction Stop
            }
            catch {
                Write-Host " ! Warning: Could not remove port. It might still be in use by another driver." -ForegroundColor Yellow
            }
        }
        Write-Host "Cleanup complete.`n" -ForegroundColor Green
        
        # Clear variables for installation
        $ExistingPrinter = $null
        $ExistingPort = $null
    }
    else {
        Write-Host "Proceeding with existing configuration...`n" -ForegroundColor Gray
    }
}


# 3. INSTALLATION
Write-Host "Installing driver to Windows Store..." -ForegroundColor Yellow
pnputil.exe /add-driver "$($InfFile.FullName)" /install

Write-Host "Adding Driver to Print Subsystem..." -ForegroundColor Yellow
Add-PrinterDriver -Name $DriverName

# Add Port (Only if it doesn't exist)
if (-not (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue)) {
    Write-Host "Creating Port: IP_$PrinterIP" -ForegroundColor Gray
    Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
}

# Add or Update Printer
if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"
    Write-Host "Success! Printer installed." -ForegroundColor Green
}
else {
    Write-Host "Printer '$PrinterName' already exists. Updating configuration..." -ForegroundColor Cyan
    Set-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"
    Write-Host "Success! Printer updated." -ForegroundColor Green
}

# 4. OPTIONAL: Set as Default
$SetDefault = Read-Host "`nSet $PrinterName as the default printer? (y/n)"
if ($SetDefault -eq 'y') {
    Set-DefaultPrinter -Name $PrinterName
    Write-Host "Set as default printer." -ForegroundColor Green
}


# Clean up temp files
Remove-Item $TempDir -Recurse -Force
