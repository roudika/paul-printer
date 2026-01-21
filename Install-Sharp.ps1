# --- CONFIGURATION ---
# Script Version: 1.2
$ZipUrl = "https://github.com/roudika/paul-printer/releases/download/BP-51C26/SHARP_BP-51C26.zip"
$PrinterIP = "10.50.30.50"
$OldPrinterIP = "10.50.30.30"
$OldPrinterName = "SHARP MX-2651"
$PrinterName = "SHARP BP-51C26 - Prod 1"
$DriverName = "SHARP BP-51C26 PCL6"

# Helper to prevent "Invalid Handle" errors
function Write-Log ($Message, $Color = "Cyan") {
    try {
        Write-Host $Message -ForegroundColor $Color
    } catch {
        Write-Output "[$Color] $Message"
    }
}

Write-Log "--- Sharp Printer Installer v1.2 ---" "Magenta"

# --- WORKSPACE ---
$TempDir = "$env:TEMP\SharpPrinter"
if (Test-Path $TempDir) { try { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {} }
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

# 0. Admin Check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "ERROR: This script MUST be run as Administrator." "Red"
    pause
    exit
}

# 0.5 Ensure Print Spooler is running
try {
    $Spooler = Get-Service -Name Spooler -ErrorAction Stop
    if ($Spooler.Status -ne 'Running') {
        Write-Log "Print Spooler is not running. Starting it now..." "Yellow"
        Start-Service -Name Spooler
        Start-Sleep -Seconds 2
    }
} catch {
    Write-Log "Warning: Could not check Print Spooler service status." "Yellow"
}

# 1. Download and Extract
Write-Log "Preparing driver files..." "Cyan"
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\driver.zip"
Expand-Archive -Path "$TempDir\driver.zip" -DestinationPath $TempDir -Force

$InfFile = Get-ChildItem -Path $TempDir -Filter "sw1emenu.inf" -Recurse | Select-Object -First 1
if (-not $InfFile) {
    Write-Log "ERROR: Could not find sw1emenu.inf." "Red"
    exit
}

# 1.5. Clean up old printer configurations
Write-Log "Searching for old device ($OldPrinterIP / $OldPrinterName)..." "Cyan"
$OldPrinters = $null
$OldPortExists = $false

try {
    # Check for printers
    $OldPrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object { 
        $_.PortName -match $OldPrinterIP -or $_.Name -like "*$OldPrinterName*" 
    }
    
    # Check for ports
    if (Get-PrinterPort -Name "IP_$OldPrinterIP" -ErrorAction SilentlyContinue) { $OldPortExists = $true }
    if (Get-PrinterPort -Name $OldPrinterIP -ErrorAction SilentlyContinue) { $OldPortExists = $true }
} catch {
    Write-Log " ! Warning: System query failed, but attempt to clean up will continue." "Yellow"
}

if ($OldPrinters -or $OldPortExists) {
    Write-Log "Old device detected. Removing..." "Yellow"
    foreach ($P in $OldPrinters) {
        Write-Log " - Removing printer: $($P.Name)" "Gray"
        Remove-Printer -Name $P.Name -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
    Remove-PrinterPort -Name "IP_$OldPrinterIP" -ErrorAction SilentlyContinue
    Remove-PrinterPort -Name $OldPrinterIP -ErrorAction SilentlyContinue
    Write-Log "Cleanup of old devices complete.`n" "Green"
} else {
    Write-Log "No old devices found matching $OldPrinterIP or $OldPrinterName.`n" "Gray"
}


# 2. CONFLICT DETECTION & CLEANUP
$ExistingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
$ExistingPort = Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue

if ($ExistingPrinter -or $ExistingPort) {
    Write-Log "`n[!] CONFLICT DETECTED" "Yellow"
    if ($ExistingPrinter) { Write-Log " - Printer '$PrinterName' already exists." "Gray" }
    if ($ExistingPort) { Write-Log " - Port 'IP_$PrinterIP' already exists." "Gray" }
    
    $UserResponse = Read-Host "`nDo you want to REMOVE existing configuration for a clean install? [Y/n]"
    if ([string]::IsNullOrWhiteSpace($UserResponse)) { $UserResponse = 'y' }
    
    if ($UserResponse -eq 'y') {
        Write-Log "Cleaning up..." "Yellow"
        
        $PrintersUsingPort = Get-Printer | Where-Object { $_.PortName -eq "IP_$PrinterIP" -or $_.Name -eq $PrinterName }
        
        foreach ($P in $PrintersUsingPort) {
            Write-Log " - Removing printer: $($P.Name)" "Gray"
            Remove-Printer -Name $P.Name -ErrorAction SilentlyContinue
        }

        Start-Sleep -Seconds 2

        if (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue) { 
            Write-Log " - Removing port: IP_$PrinterIP" "Gray"
            try { Remove-PrinterPort -Name "IP_$PrinterIP" -ErrorAction Stop } catch {
                Write-Log " ! Note: Port is busy. Will attempt to reconfigure it in-place." "Yellow"
            }
        }
        Write-Log "Cleanup complete.`n" "Green"
        $ExistingPrinter = $null
        $ExistingPort = $null
    }
}


# 3. INSTALLATION
Write-Log "Installing driver to Windows Driver Store..." "Yellow"
pnputil.exe /add-driver "$($InfFile.FullName)" /install | Out-Null

Write-Log "Adding Driver to Print Subsystem..." "Yellow"
try {
    Add-PrinterDriver -Name $DriverName -ErrorAction Stop
} catch {
    Write-Log "Note: Add-PrinterDriver info: $($_.Exception.Message)" "Gray"
}

# Add Port
if (-not (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue)) {
    Write-Log "Creating Port: IP_$PrinterIP" "Gray"
    Add-PrinterPort -Name "IP_$IP_$PrinterIP" -PrinterHostAddress $PrinterIP -ErrorAction SilentlyContinue
    if (-not $?) {
        # Fallback if specific naming fails
        Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
    }
}

# Add or Update Printer
Write-Log "Configuring Printer: $PrinterName" "Yellow"
$OperationSuccess = $false

try {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction Stop
    Write-Log "Success! Printer installed." "Green"
    $OperationSuccess = $true
}
catch {
    try {
        Set-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction Stop
        Write-Log "Success! Printer updated." "Green"
        $OperationSuccess = $true
    }
    catch {
        Write-Log "ERROR: Could not configure printer. Please check if another printer is using the name or port." "Red"
    }
}

# 4. OPTIONAL: Set as Default
$SetDefault = Read-Host "`nSet $PrinterName as the default printer? [Y/n]"
if ([string]::IsNullOrWhiteSpace($SetDefault)) { $SetDefault = 'y' }

if ($SetDefault -eq 'y' -and $OperationSuccess) {
    try {
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($PrinterName)
        Write-Log "Set as default printer." "Green"
    } catch {
        Write-Log "Could not set default printer automatically." "Yellow"
    }
}

# Clean up temp files
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Log "`nAll steps completed." "Green"

