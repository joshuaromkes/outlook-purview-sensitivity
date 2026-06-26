<#
.SYNOPSIS
    Installs or uninstalls the Outlook Purview Sensitivity VSTO add-in.

.DESCRIPTION
    Copies the published VSTO files to a permanent location and registers
    the add-in under HKCU\Software\Microsoft\Office\Outlook\AddIns so
    Outlook loads it with LoadBehavior=3 (startup).

.PARAMETER InstallPath
    Path to the published output folder containing the .vsto and .application files.
    Required for -Install.

.PARAMETER Uninstall
    Remove the add-in from the registry and delete installed files.

.EXAMPLE
    .\Install-AddIn.ps1 -InstallPath .\bin\Release\app.publish

.EXAMPLE
    .\Install-AddIn.ps1 -Uninstall
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$InstallPath,

    [Parameter(Mandatory = $false)]
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$AddInName = "Outlook-Purview-Sensitivity"
$RegKeyPath = "HKCU:\Software\Microsoft\Office\Outlook\AddIns\$AddInName"
$InstallDir = "$env:LocalAppData\Outlook-Purview-Sensitivity"

function Test-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-OutlookInstalled {
    $path = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Office\Outlook\AddIns" -ErrorAction SilentlyContinue
    return $null -ne $path
}

function Uninstall-AddIn {
    Write-Host "Uninstalling $AddInName..." -ForegroundColor Yellow

    if (Test-Path $RegKeyPath) {
        Remove-Item -Path $RegKeyPath -Recurse -Force
        Write-Host "  Registry key removed." -ForegroundColor Green
    }
    else {
        Write-Host "  Registry key not found — already uninstalled." -ForegroundColor Gray
    }

    if (Test-Path $InstallDir) {
        Remove-Item -Path $InstallDir -Recurse -Force
        Write-Host "  Installed files removed from $InstallDir" -ForegroundColor Green
    }

    Write-Host "Uninstall complete. Restart Outlook if it is running." -ForegroundColor Yellow
}

function Install-AddIn {
    if (-not $InstallPath) {
        Write-Host "ERROR: -InstallPath is required for installation." -ForegroundColor Red
        Write-Host "Usage: .\Install-AddIn.ps1 -InstallPath path-to-publish-folder" -ForegroundColor Gray
        exit 1
    }

    if (-not (Test-Path $InstallPath)) {
        Write-Host "ERROR: InstallPath '$InstallPath' does not exist." -ForegroundColor Red
        exit 1
    }

    Write-Host "Installing $AddInName from $InstallPath" -ForegroundColor Cyan

    # Step 1: Check prerequisites
    if (-not (Test-OutlookInstalled)) {
        Write-Host "WARNING: Outlook does not appear to be installed on this machine." -ForegroundColor Yellow
    }

    # Step 2: Stop Outlook if running
    $outlook = Get-Process outlook -ErrorAction SilentlyContinue
    if ($outlook) {
        Write-Host "  Outlook is running — please close Outlook before installing." -ForegroundColor Yellow
        Stop-Process -Name outlook -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }

    # Step 3: Remove any existing install
    Uninstall-AddIn

    # Step 4: Copy files
    Write-Host "  Copying files to $InstallDir ..."
    New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path "$InstallPath\*" -Destination $InstallDir -Recurse -Force

    # Find the .vsto file
    $vstoFile = Get-ChildItem -Path $InstallDir -Filter *.vsto -Recurse | Select-Object -First 1
    if (-not $vstoFile) {
        Write-Host "ERROR: No .vsto file found in $InstallDir. Was the project published?" -ForegroundColor Red
        Write-Host "Run 'msbuild /t:Publish /p:PublishProfile=FolderProfile /p:Configuration=Release' first." -ForegroundColor Gray
        exit 1
    }

    $manifestUrl = "file:///" + $vstoFile.FullName.Replace('\', '/') + "|vstolocal"
    Write-Host "  Manifest: $manifestUrl" -ForegroundColor Gray

    # Step 5: Register in registry
    Write-Host "  Registering add-in in HKCU..." -ForegroundColor Gray
    New-Item -Path $RegKeyPath -Force | Out-Null
    Set-ItemProperty -Path $RegKeyPath -Name "Description" -Value "Displays Microsoft Purview sensitivity labels in Outlook" -Type String
    Set-ItemProperty -Path $RegKeyPath -Name "FriendlyName" -Value "Outlook Purview Sensitivity" -Type String
    Set-ItemProperty -Path $RegKeyPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $RegKeyPath -Name "Manifest" -Value $manifestUrl -Type String

    Write-Host "  Add-in registered with LoadBehavior=3 (load at startup)" -ForegroundColor Green

    # Step 6: Verify registration
    $verify = Get-ItemProperty -Path $RegKeyPath -ErrorAction SilentlyContinue
    if ($verify -and $verify.LoadBehavior -eq 3) {
        Write-Host "Installation successful." -ForegroundColor Green
        Write-Host "Start Outlook to see the PurviewLabel column in your message list." -ForegroundColor Cyan
    }
    else {
        Write-Host "WARNING: Registry verification failed. Check key: $RegKeyPath" -ForegroundColor Red
    }
}

# Main
if ($Uninstall) {
    Uninstall-AddIn
}
else {
    Install-AddIn
}
