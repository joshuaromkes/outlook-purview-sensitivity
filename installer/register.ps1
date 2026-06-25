# Development Registration Script
# Run as administrator in PowerShell after building in Debug mode.
# Registers the raw COM add-in for the current user.
#
# Usage: .\installer\register.ps1
#        .\installer\register.ps1 -Unregister

param(
    [switch]$Unregister
)

$addinName = "OutlookPurviewColumn.AddIn"
$dllPath = Join-Path $PSScriptRoot "..\src\OutlookPurviewColumn\bin\Debug\OutlookPurviewColumn.dll"

$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$addinName"

if ($Unregister) {
    Write-Host "Unregistering $addinName..."

    # Unregister COM
    $regasm = "${env:SystemRoot}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    & $regasm $dllPath /unregister /silent

    # Remove Outlook add-in registry key
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

# Register COM (creates CLSID entries)
$regasm = "${env:SystemRoot}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
& $regasm $dllPath /codebase

# Create Outlook add-in registry key
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "Description" -Value "Adds Purview sensitivity label column to Outlook" -Type String
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Purview Label Column" -Type String
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord

Write-Host "Add-in registered. Restart Outlook."
Write-Host "Registry path: $regPath"
