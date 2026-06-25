# Intune Deployment Guide

## Prerequisites

- Code signing certificate (EV or standard) in `.pfx` format
- The MSI built from `installer/Product.wxs`
- Intune admin access

## Step 1: Build & Sign

```powershell
# Build Release configuration in Visual Studio
msbuild src\OutlookPurviewColumn.sln /p:Configuration=Release

# Sign the DLL
signtool sign /fd SHA256 /f cert.pfx /p <password> `
    /tr http://timestamp.digicert.com `
    src\OutlookPurviewColumn\bin\Release\OutlookPurviewColumn.dll

# Build the MSI via WiX
cd installer
candle Product.wxs -o Product.wixobj
light Product.wixobj -o OutlookPurviewColumn.msi

# Sign the MSI
signtool sign /fd SHA256 /f cert.pfx /p <password> `
    /tr http://timestamp.digicert.com `
    OutlookPurviewColumn.msi
```

## Step 2: Package for Intune

Use the [Microsoft Win32 Content Prep Tool](https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool):

```powershell
IntuneWinAppUtil.exe -c .\installer -s OutlookPurviewColumn.msi `
    -o .\installer -q
```

This creates `OutlookPurviewColumn.intunewin`.

## Step 3: Deploy in Intune

1. **Apps** → **Windows** → **Add** → **Windows app (Win32)**
2. Upload `OutlookPurviewColumn.intunewin`
3. Configure:
   - **Install command:** `msiexec /i OutlookPurviewColumn.msi /quiet /norestart`
   - **Uninstall command:** `msiexec /x {ProductCode} /quiet /norestart`
   - **Install behavior:** System
   - **Device restart behavior:** No specific action
4. **Detection rules:**
   - Rule type: Registry
   - Key path: `HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\OutlookPurviewColumn`
   - Value name: `LoadBehavior`
   - Detection method: Integer comparison
   - Operator: Equals
   - Value: `3`
5. **Assignments:** target your user group

## Step 4: Verify

After deployment, users should:
1. Restart Outlook (or reboot)
2. Navigate to any mail folder
3. See the "Purview Label" column in the view
4. Emails with Purview labels show the label name (e.g., "PII High")

If the column doesn't appear:
- Right-click the column header → **Field Chooser**
- Under "User-defined fields in folder", drag "PurviewLabel" into the view

## Uninstall

Standard Intune uninstall or `msiexec /x` removes registry keys and files.
No user data is affected — all label data lives in Outlook's own UserProperties.
