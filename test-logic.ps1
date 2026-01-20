$ZipUrl = 'https://github.com/roudika/paul-printer/archive/refs/tags/v1.0.zip'
$TempDir = Join-Path 'c:\Test\Ladenburg\paul-printer' 'TestTemp'

if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
New-Item -ItemType Directory -Path $TempDir | Out-Null

Write-Host "Downloading..."
Invoke-WebRequest -Uri $ZipUrl -OutFile "$TempDir\main.zip"

Write-Host "Extracting main..."
Expand-Archive -Path "$TempDir\main.zip" -DestinationPath $TempDir -Force

Write-Host "Looking for SHARP-MX3061.zip..."
$InternalZip = Get-ChildItem -Path $TempDir -Filter "SHARP-MX3061.zip" -Recurse | Select-Object -First 1

if ($InternalZip) {
    Write-Host "Found Internal Zip: $($InternalZip.FullName)"
    $DriverFilesPath = "$TempDir\DriverUnzipped"
    Expand-Archive -Path $InternalZip.FullName -DestinationPath $DriverFilesPath -Force
    
    $InfFile = Get-ChildItem -Path $DriverFilesPath -Filter "su2emenu.inf" -Recurse | Select-Object -First 1
    if ($InfFile) {
        Write-Host "Found INF File: $($InfFile.FullName)"
        Write-Host "Download and Extraction Test: SUCCESS" -ForegroundColor Green
    }
    else {
        Write-Error "INF file 'su2emenu.inf' not found in extracted driver."
    }
}
else {
    Write-Error "Internal zip 'SHARP-MX3061.zip' not found in repo archive."
}

# Cleanup
# Remove-Item $TempDir -Recurse -Force
