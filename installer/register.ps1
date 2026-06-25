# Development Registration Script
# Run as administrator in PowerShell after building in Debug mode.
# This registers the add-in for the current user without needing the MSI.
#
# Usage: .\installer\register.ps1

param(
    [switch]$Unregister
)

$addinName = "OutlookPurviewColumn"
$dllPath = Join-Path $PSScriptRoot "..\src\OutlookPurviewColumn\bin\Debug\OutlookPurviewColumn.dll"
$vstoPath = Join-Path $PSScriptRoot "..\src\OutlookPurviewColumn\bin\Debug\OutlookPurviewColumn.vsto"
$manifestPath = Join-Path $PSScriptRoot "..\src\OutlookPurviewColumn\Properties\OutlookPurviewColumn.dll.manifest"

$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$addinName"

if ($Unregister) {
    Write-Host "Unregistering $addinName..."
    Remove-Item -Path $regPath -Recurse -ErrorAction SilentlyContinue
    Write-Host "Done. Restart Outlook."
    return
}

# Verify build output exists
if (-not (Test-Path $dllPath)) {
    Write-Error "Build output not found at: $dllPath"
    Write-Host "Build the project in Debug mode first."
    return
}

Write-Host "Registering $addinName..."

# Create registry keys
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "Description" -Value "Adds Purview sensitivity label column to Outlook" -Type String
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Purview Label Column" -Type String
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $regPath -Name "Manifest" -Value "file:///$($vstoPath -replace '\\','/')|vstolocal" -Type String

Write-Host "Add-in registered. Restart Outlook."
Write-Host "Registry path: $regPath"
