# --- CONFIGURATION ---
$ZipUrl = "https://github.com/roudika/paul-printer/releases/download/BP-51C26/SHARP_BP-51C26.zip"
$PrinterIP = "10.50.30.50"
$OldPrinterIP = "10.50.30.30"
$OldPrinterName = "SHARP MX-2651"
$PrinterName = "SHARP BP-51C26 - Prod 1"
$DriverName = "SHARP BP-51C26 PCL6"

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
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\driver.zip"
Expand-Archive -Path "$TempDir\driver.zip" -DestinationPath $TempDir -Force

# Automatically find the INF file (usually in English\PCL6\64bit\)
$InfFile = Get-ChildItem -Path $TempDir -Filter "sw1emenu.inf" -Recurse | Select-Object -First 1

if (-not $InfFile) {
    Write-Host "ERROR: Could not find sw1emenu.inf in the extracted files." -ForegroundColor Red
    exit
}

# 1.5. Clean up old printer (10.50.30.30 and SHARP MX-2651) if exists
Write-Host "Checking for old printer configurations..." -ForegroundColor Cyan
$OldPortName = "IP_$OldPrinterIP"

# Find printers by IP or Name
$OldPrinters = Get-Printer | Where-Object { $_.PortName -eq $OldPortName -or $_.Name -eq $OldPrinterName }

if ($OldPrinters -or (Get-PrinterPort -Name $OldPortName -ErrorAction SilentlyContinue)) {
    Write-Host "Old printer/port detected. Removing..." -ForegroundColor Yellow
    
    foreach ($P in $OldPrinters) {
        Write-Host " - Removing printer: $($P.Name)" -ForegroundColor Gray
        Remove-Printer -Name $P.Name -ErrorAction SilentlyContinue
    }
    
    # Wait for spooler to update
    Start-Sleep -Seconds 2
    
    if (Get-PrinterPort -Name $OldPortName -ErrorAction SilentlyContinue) {
        Write-Host " - Removing port: $OldPortName" -ForegroundColor Gray
        Remove-PrinterPort -Name $OldPortName -ErrorAction SilentlyContinue
    }
    Write-Host "Cleanup of old devices complete.`n" -ForegroundColor Green
}


# 2. CONFLICT DETECTION & CLEANUP
$ExistingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
$ExistingPort = Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue

if ($ExistingPrinter -or $ExistingPort) {
    Write-Host "`n[!] CONFLICT DETECTED" -ForegroundColor Yellow
    if ($ExistingPrinter) { Write-Host " - Printer '$PrinterName' already exists." -ForegroundColor Gray }
    if ($ExistingPort) { Write-Host " - Port 'IP_$PrinterIP' already exists." -ForegroundColor Gray }
    
    $UserResponse = Read-Host "`nDo you want to REMOVE existing configuration for a clean install? [Y/n]"
    if ([string]::IsNullOrWhiteSpace($UserResponse)) { $UserResponse = 'y' }
    
    if ($UserResponse -eq 'y') {
        Write-Host "Cleaning up..." -ForegroundColor Yellow
        
        # Find ALL printers using this port (even those with different names)
        $PrintersUsingPort = Get-Printer | Where-Object { $_.PortName -eq "IP_$PrinterIP" -or $_.Name -eq $PrinterName }
        
        if ($PrintersUsingPort.Count -gt 0) {
            foreach ($P in $PrintersUsingPort) {
                Write-Host " - Removing printer: $($P.Name)" -ForegroundColor Gray
                Remove-Printer -Name $P.Name -ErrorAction SilentlyContinue
            }

            # Wait for printers to be fully purged (up to 10 seconds)
            Write-Host " - Waiting for spooler to release port..." -ForegroundColor Gray
            for ($i = 0; $i -lt 10; $i++) {
                $Remaining = Get-Printer | Where-Object { $_.PortName -eq "IP_$PrinterIP" -or $_.Name -eq $PrinterName }
                if (-not $Remaining) { break }
                Start-Sleep -Seconds 1
            }
        }

        if (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue) { 
            Write-Host " - Removing port: IP_$PrinterIP" -ForegroundColor Gray
            # Attempt port removal multiple times
            for ($j = 0; $j -lt 5; $j++) {
                try {
                    Remove-PrinterPort -Name "IP_$PrinterIP" -ErrorAction Stop
                    $PortRemoved = $true
                    break
                }
                catch {
                    Start-Sleep -Seconds 1
                }
            }
            if (-not $PortRemoved) {
                Write-Host " ! Note: Port is busy. Will attempt to reconfigure it in-place." -ForegroundColor Yellow
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
Write-Host "Installing driver to Windows Driver Store..." -ForegroundColor Yellow
pnputil.exe /add-driver "$($InfFile.FullName)" /install | Out-Null

Write-Host "Adding Driver to Print Subsystem..." -ForegroundColor Yellow
Add-PrinterDriver -Name $DriverName

# Add Port (Only if it doesn't exist)
if (-not (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue)) {
    Write-Host "Creating Port: IP_$PrinterIP" -ForegroundColor Gray
    Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
}

# Add or Update Printer
# We use a combined approach to handle 'ghost' printers that Get-Printer might miss
Write-Host "Configuring Printer: $PrinterName" -ForegroundColor Yellow
$OperationSuccess = $false

# Try 1: Add fresh
try {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction Stop
    Write-Host "Success! Printer installed." -ForegroundColor Green
    $OperationSuccess = $true
}
catch {
    # If it fails (e.g. already exists), try to update it
    try {
        Set-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction Stop
        Write-Host "Success! Printer updated." -ForegroundColor Green
        $OperationSuccess = $true
    }
    catch {
        Write-Host " ! Warning: Initial config failed. Retrying in 2 seconds..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Set-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction SilentlyContinue
        if ($?) { 
            Write-Host "Success! Printer configured." -ForegroundColor Green
            $OperationSuccess = $true
        }
    }
}

if (-not $OperationSuccess) {
    Write-Host "ERROR: Could not configure printer. Please check if another printer is using the name or port." -ForegroundColor Red
}


# 4. OPTIONAL: Set as Default
$SetDefault = Read-Host "`nSet $PrinterName as the default printer? [Y/n]"
if ([string]::IsNullOrWhiteSpace($SetDefault)) { $SetDefault = 'y' }

if ($SetDefault -eq 'y') {
    (New-Object -ComObject WScript.Network).SetDefaultPrinter($PrinterName)
    Write-Host "Set as default printer." -ForegroundColor Green
}


# Clean up temp files
Remove-Item $TempDir -Recurse -Force
