# Decksterity PowerShell Installer Creator
# This script creates a simple MSI installer for the Decksterity PowerPoint add-in

param(
    [string]$OutputPath = ".\DecksteritySetup.msi",
    [string]$SourcePath = ".\bin\Release"
)

Write-Host "Creating Decksterity PowerPoint Add-in Installer..." -ForegroundColor Green

# Check if source files exist
if (!(Test-Path "$SourcePath\decksterity.dll")) {
    Write-Error "Build files not found. Please build the project in Release mode first."
    exit 1
}

# Create temporary directory for installer content
$TempDir = New-TemporaryFile | %{ rm $_; mkdir $_ }
$InstallDir = Join-Path $TempDir "Install"
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

Write-Host "Copying files..." -ForegroundColor Yellow

# Copy main files
Copy-Item "$SourcePath\decksterity.dll" $InstallDir
Copy-Item "$SourcePath\decksterity.dll.manifest" $InstallDir
Copy-Item "$SourcePath\decksterity.vsto" $InstallDir

# Copy dependencies if they exist
if (Test-Path "$SourcePath\Microsoft.Office.Tools.Common.v4.0.Utilities.dll") {
    Copy-Item "$SourcePath\Microsoft.Office.Tools.Common.v4.0.Utilities.dll" $InstallDir
}

# Create installer script
$InstallerScript = @"
# Decksterity PowerPoint Add-in Installer
Write-Host "Installing Decksterity PowerPoint Add-in..." -ForegroundColor Green

# Define installation path
`$InstallPath = "`$env:ProgramFiles\Decksterity"
if (!(Test-Path `$InstallPath)) {
    New-Item -ItemType Directory -Path `$InstallPath -Force | Out-Null
}

# Copy files
Write-Host "Copying files to `$InstallPath..." -ForegroundColor Yellow
Copy-Item ".\*" `$InstallPath -Recurse -Force

# Register VSTO add-in
Write-Host "Registering PowerPoint add-in..." -ForegroundColor Yellow
`$RegPath = "HKLM:\SOFTWARE\Microsoft\Office\PowerPoint\Addins\decksterity"

if (!(Test-Path `$RegPath)) {
    New-Item -Path `$RegPath -Force | Out-Null
}

Set-ItemProperty -Path `$RegPath -Name "Description" -Value "Enhanced slide elements and layout tools for PowerPoint"
Set-ItemProperty -Path `$RegPath -Name "FriendlyName" -Value "Decksterity PowerPoint Add-in"
Set-ItemProperty -Path `$RegPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path `$RegPath -Name "Manifest" -Value "`$InstallPath\decksterity.vsto|vstolocal"

Write-Host "Installation completed successfully!" -ForegroundColor Green
Write-Host "You can now start PowerPoint to use the Decksterity add-in." -ForegroundColor Cyan

Read-Host "Press Enter to continue..."
"@

# Save installer script
Set-Content -Path (Join-Path $InstallDir "Install.ps1") -Value $InstallerScript

# Create uninstaller script
$UninstallerScript = @"
# Decksterity PowerPoint Add-in Uninstaller
Write-Host "Uninstalling Decksterity PowerPoint Add-in..." -ForegroundColor Yellow

# Remove registry entries
`$RegPath = "HKLM:\SOFTWARE\Microsoft\Office\PowerPoint\Addins\decksterity"
if (Test-Path `$RegPath) {
    Remove-Item -Path `$RegPath -Recurse -Force
    Write-Host "Registry entries removed." -ForegroundColor Green
}

# Remove installation directory
`$InstallPath = "`$env:ProgramFiles\Decksterity"
if (Test-Path `$InstallPath) {
    Remove-Item -Path `$InstallPath -Recurse -Force
    Write-Host "Installation files removed." -ForegroundColor Green
}

Write-Host "Uninstallation completed successfully!" -ForegroundColor Green
Read-Host "Press Enter to continue..."
"@

Set-Content -Path (Join-Path $InstallDir "Uninstall.ps1") -Value $UninstallerScript

# Create README
$ReadmeContent = @"
# Decksterity PowerPoint Add-in Installation

## Installation
1. Right-click on Install.ps1
2. Select "Run with PowerShell"
3. If prompted about execution policy, choose "Yes" or "Run anyway"
4. Follow the on-screen instructions

## Requirements
- Microsoft PowerPoint 2016 or later
- .NET Framework 4.7.2 or higher
- Visual Studio Tools for Office Runtime

## Uninstallation
1. Right-click on Uninstall.ps1
2. Select "Run with PowerShell"
3. Follow the on-screen instructions

## Features
- Harvey Balls (progress indicators)
- Directional arrows
- Icons (check, cross, plus, minus, question, ellipsis)
- Stoplight indicators  
- Advanced alignment and distribution tools

## Support
For support or questions, visit: https://github.com/avirut/decksterity
"@

Set-Content -Path (Join-Path $InstallDir "README.txt") -Value $ReadmeContent

Write-Host "Creating ZIP package..." -ForegroundColor Yellow

# Create ZIP file instead of MSI for simplicity
$ZipPath = $OutputPath -replace '\.msi$', '.zip'
if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

# Compress to ZIP
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($InstallDir, $ZipPath)

Write-Host "Installer package created: $ZipPath" -ForegroundColor Green

# Cleanup
Remove-Item $TempDir -Recurse -Force

Write-Host "Done! Distribute the ZIP file for installation." -ForegroundColor Cyan