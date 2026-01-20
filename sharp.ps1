# 1. Define your GitHub ZIP URL (Replace with your actual URL)
$repoZip = "https://github.com/USER/REPO/archive/refs/heads/main.zip"
$dest = "$env:TEMP\PrinterSetup"

# 2. Download and Extract everything
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Invoke-WebRequest -Uri $repoZip -OutFile "$env:TEMP\repo.zip"
Expand-Archive -Path "$env:TEMP\repo.zip" -DestinationPath $dest

# 3. Find and run the script inside the extracted folder
$script = Get-ChildItem -Path $dest -Filter "Install-Printer.ps1" -Recurse | Select-Object -ExpandProperty FullName
powershell.exe -ExecutionPolicy Bypass -File $script
