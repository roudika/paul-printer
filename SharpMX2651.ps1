# --- CONFIGURATION ---
# Script Version: 1.0
# Sharp MX-2651 installer - download drivers from GitHub, remove old printers (EPSON/WF-C869R), add new one.

$ZipUrl = "https://github.com/roudika/paul-printer/releases/download/MX2651/MX2651.zip"
$PrinterIP = "192.168.10.69"
$PrinterName = "THA-Sharp"
$DriverName = "SHARP MX-2651 PCL6"

# IPs and names to remove before install: EPSON (any), WF-C869R (any), 192.168.10.69, 192.168.19.69
$OldPrinterIPs = @("192.168.10.69", "192.168.19.69")
$OldPrinterNamePatterns = @("*EPSON*", "*WF-C869R*")
$InfFileName = "su2emenu.inf"

# Helper to prevent "Invalid Handle" errors
function Write-Log ($Message, $Color = "Cyan") {
    try {
        Write-Host $Message -ForegroundColor $Color
    } catch {
        Write-Output "[$Color] $Message"
    }
}

Write-Log "--- Sharp MX-2651 Printer Installer v1.0 ---" "Magenta"

# --- WORKSPACE ---
$TempDir = "$env:TEMP\SharpMX2651"
if (Test-Path $TempDir) { try { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {} }
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

# 0. Admin Check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "ERROR: This script MUST be run as Administrator." "Red"
    pause
    exit 1
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
Write-Log "Downloading driver package..." "Cyan"
try {
    Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\driver.zip" -UseBasicParsing
} catch {
    Write-Log "ERROR: Failed to download driver zip. Check URL and network. $($_.Exception.Message)" "Red"
    exit 1
}

Write-Log "Extracting driver files..." "Cyan"
Expand-Archive -Path "$TempDir\driver.zip" -DestinationPath $TempDir -Force

$InfFile = Get-ChildItem -Path $TempDir -Filter $InfFileName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $InfFile) {
    Write-Log "ERROR: Could not find $InfFileName in the extracted package." "Red"
    exit 1
}

# 2. Remove old printers: EPSON, WF-C869R (by name), 192.168.10.69, 192.168.19.69 (by port)
Write-Log "Checking for old printers to remove (EPSON, WF-C869R, 192.168.10.69, 192.168.19.69)..." "Cyan"
$ToRemove = @()

try {
    $AllPrinters = Get-Printer -ErrorAction Stop
    foreach ($P in $AllPrinters) {
        $remove = $false
        foreach ($pattern in $OldPrinterNamePatterns) {
            if ($P.Name -like $pattern) { $remove = $true; break }
        }
        foreach ($ip in $OldPrinterIPs) {
            if ($P.PortName -eq "IP_$ip" -or $P.PortName -eq $ip) { $remove = $true }
        }
        if ($remove) { $ToRemove += $P }
    }
} catch {
    Write-Log "Warning: Could not enumerate printers. Cleanup will continue." "Yellow"
}

if ($ToRemove.Count -gt 0) {
    Write-Log "Removing $($ToRemove.Count) old printer(s)..." "Yellow"
    foreach ($P in $ToRemove) {
        Write-Log " - Removing printer: $($P.Name)" "Gray"
        Remove-Printer -Name $P.Name -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
}

# Remove ports for old IPs (so we can re-add 192.168.19.69 clean)
foreach ($ip in $OldPrinterIPs) {
    foreach ($PortName in @("IP_$ip", $ip)) {
        if (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue) {
            Write-Log " - Removing port: $PortName" "Gray"
            Remove-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
        }
    }
}
if ($ToRemove.Count -gt 0) { Write-Log "Cleanup of old devices complete.`n" "Green" }

# 3. Conflict check: if same printer/port still exists, offer to remove
$ExistingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
$ExistingPort = Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue

if ($ExistingPrinter -or $ExistingPort) {
    Write-Log "`n[!] Printer or port already exists." "Yellow"
    if ($ExistingPrinter) { Write-Log " - Printer '$PrinterName' exists." "Gray" }
    if ($ExistingPort) { Write-Log " - Port 'IP_$PrinterIP' exists." "Gray" }
    $UserResponse = Read-Host "`nRemove existing and reinstall? [Y/n]"
    if ([string]::IsNullOrWhiteSpace($UserResponse)) { $UserResponse = 'y' }
    if ($UserResponse -eq 'y') {
        $PrintersUsingPort = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq "IP_$PrinterIP" -or $_.Name -eq $PrinterName }
        foreach ($P in $PrintersUsingPort) {
            Write-Log " - Removing printer: $($P.Name)" "Gray"
            Remove-Printer -Name $P.Name -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 2
        Remove-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue
        Remove-PrinterPort -Name $PrinterIP -ErrorAction SilentlyContinue
        Write-Log "Cleanup complete.`n" "Green"
    }
}

# 4. Install driver
Write-Log "Installing driver to Windows Driver Store..." "Yellow"
pnputil.exe /add-driver "$($InfFile.FullName)" /install 2>&1 | Out-Null

Write-Log "Adding driver to Print subsystem..." "Yellow"
try {
    Add-PrinterDriver -Name $DriverName -ErrorAction Stop
} catch {
    Write-Log "Note: $($_.Exception.Message)" "Gray"
}

# 5. Add port and printer
if (-not (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue)) {
    Write-Log "Creating port: IP_$PrinterIP" "Gray"
    Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP -ErrorAction SilentlyContinue
    if (-not $?) {
        Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP
    }
}

Write-Log "Adding printer: $PrinterName on $PrinterIP" "Yellow"
$OperationSuccess = $false
try {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction Stop
    Write-Log "Success! Printer installed." "Green"
    $OperationSuccess = $true
} catch {
    try {
        Set-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP" -ErrorAction Stop
        Write-Log "Success! Printer updated." "Green"
        $OperationSuccess = $true
    } catch {
        Write-Log "ERROR: Could not add or update printer. Check name/port/driver." "Red"
    }
}

# 6. Optional: Set as default
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

# Cleanup
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Log "`nDone." "Green"
