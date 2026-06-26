# Outlook Purview Sensitivity

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![.NET Framework 4.7.2](https://img.shields.io/badge/.NET-4.7.2-purple.svg)](https://dotnet.microsoft.com/)
[![Outlook 64-bit](https://img.shields.io/badge/Outlook-2016%2B%2064--bit-orange.svg)](https://www.microsoft.com/microsoft-365/outlook/)

> **🤖 AI-Generated Code Disclosure:** This project was developed with AI assistance ([Hermes Agent](https://github.com/NousResearch/hermes-agent) + [OpenCode](https://opencode.ai)), using a structured development workflow of research, planning, implementation, and iterative debugging. All code has been tested and verified in real Outlook environments. See [Contributing](#contributing) for guidelines on submitting changes.

Displays Microsoft Purview sensitivity labels as a sortable **"Sensitivity Label"** column in classic Outlook (desktop). Works across all folders — Inbox, Sent Items, subfolders, and shared mailboxes.

Built as a VSTO COM add-in. No modification of email subjects, categories, or body content.

## Features

- Adds a **Sensitivity Label** column to every folder automatically
- Reads the `msip_labels` MAPI property via `PropertyAccessor` to extract the label name
- Stamps labels into a user-defined field so they appear as a sortable/filterable column
- New emails stamped on arrival via `ItemAdd` event — no manual refresh needed
- Handles shared mailboxes, delegate mailboxes, and PST files gracefully
- Startup health check with debug logging (`[PS]` / `[CM]` prefixes)
- PowerShell installer with per-user HKCU registration (no admin rights needed)
- ClickOnce publish profile included for enterprise deployment via Intune / SCCM

## Tips for Users

Outlook already has everything you need to work with the Sensitivity Label column — no extra features required.

**Sort by label:** Click the **Sensitivity Label** column header. Click again to reverse.

**Find all emails with a specific label:** Press **Ctrl+Shift+F** → **Advanced** tab → **Field** → **User-defined fields in folder** → **Sensitivity Label**. Set your condition and click Find Now.

**Group by label:** **View** → **View Settings** → **Group By** → select **Sensitivity Label**. All emails with the same label are grouped together with collapsible headers.

## System Requirements

| Component | Requirement |
|-----------|-------------|
| Windows | Windows 10 or later (x64) |
| Outlook | Microsoft Outlook 2016+ (64-bit) |
| .NET | .NET Framework 4.7.2 |
| VSTO Runtime | Visual Studio 2010 Tools for Office Runtime 4.0 |
| Build tooling | Visual Studio 2022 (Community or higher) |

> The VSTO runtime ships with Office 2016+ and Visual Studio. If needed, `setup.exe` downloads it automatically.

## How It Works

1. On startup, the add-in logs a health check and waits for an Explorer window
2. When a folder is opened, it creates the **Sensitivity Label** user-defined property on that folder (if missing) and adds it as a column to the current view
3. The first 50 labeled items in the folder are stamped automatically
4. A `FolderSwitch` event handler ensures every folder gets processed as you navigate
5. An `ItemAdd` event handler stamps new labeled emails the moment they arrive
6. Labels are read from the `msip_labels` MAPI header — no modification of the email itself

## Quick Start

```powershell
git clone https://github.com/joshuaromkes/outlook-purview-sensitivity.git
cd outlook-purview-sensitivity
```

Open `Outlook-Purview-Sensitivity.slnx` in Visual Studio 2022 and press **F5**.

## Installation & Uninstall

VSTO add-ins don't have a native install/uninstall mechanism — they rely on ClickOnce or registry registration. This project supports both.

### Step 1: Build & Publish

First, publish the add-in from Visual Studio to produce the `.vsto` deployment files:

```powershell
# In Visual Studio: Build → Publish → select "FolderProfile" → Publish
#
# Or from command line:
msbuild /t:Publish /p:PublishProfile=FolderProfile /p:Configuration=Release
```

This creates `bin\Release\app.publish\` containing the `.vsto` manifest, `.application` deployment file, and `setup.exe` bootstrapper.

### Step 2: Install

```
setup.exe /install
```

### Uninstall

Manually:

```
setup.exe /uninstall
```

Or via Windows Settings → Apps & Features → Outlook Purview Sensitivity → Uninstall.

### Intune / SCCM Deployment

1. Publish the add-in using Step 1 above
2. Deploy `setup.exe /install` as a user-context script or package

### Creating a Release (for maintainers)

```powershell
# 1. Publish the add-in
msbuild /t:Publish /p:PublishProfile=FolderProfile /p:Configuration=Release

# 2. Create the release zip (includes installer script + published output)
Compress-Archive -Path .\bin\Release\app.publish\* -DestinationPath .\Outlook-Purview-Sensitivity-v1.0.0.zip
```

Upload the zip to a [GitHub release](https://github.com/joshuaromkes/outlook-purview-sensitivity/releases). End users unzip and run `setup.exe` from the extracted folder.

## Troubleshooting

### Column doesn't appear in a folder

The column is added automatically when you open the folder for the first time. If it's missing:
1. Switch the folder to **Messages** view (View → Change View → Messages)
2. Restart Outlook to re-trigger folder discovery

### Add-in doesn't load

1. Check `HKCU\Software\Microsoft\Office\Outlook\AddIns\Outlook-Purview-Sensitivity` — `LoadBehavior` must be `3`
2. File → Options → Add-ins → Manage: Disabled Items → Go → re-enable if listed
3. Re-run `setup.exe` from the published folder

### "Save failed" errors in debug output

Expected on PST files, delegate mailboxes, or shared mailboxes without write access. The add-in logs and continues. Use [DebugView](https://learn.microsoft.com/en-us/sysinternals/downloads/debugview) to see trace output — filter for `[PS]` (startup/health) or `[CM]` (column operations).

## Architecture

| File | Responsibility |
|------|---------------|
| `ThisAddIn.cs` | Startup/shutdown, Explorer event wiring, health checks, `ItemAdd`/`FolderSwitch` handlers |
| `ColumnManager.cs` | User-defined property creation, view column management, item stamping |
| `LabelReader.cs` | Reads `msip_labels` from MAPI `PropertyAccessor` |
| `LabelResolver.cs` | Parses human-readable label name from the `msip_labels` string |
| `Properties/PublishProfiles/FolderProfile.pubxml` | ClickOnce publish profile for folder-based deployment |

## Contributing

This project is AI-assisted but human-maintained. Contributions are welcome — here's how:

1. **Issues:** Found a bug or have a feature idea? [Open an issue](https://github.com/joshuaromkes/outlook-purview-sensitivity/issues) first to discuss it
2. **Pull Requests:** Fork the repo, create a feature branch, and submit a PR against `main`
3. **AI-assisted contributions:** If you use AI tools to generate code for a PR, please disclose it in the PR description
4. **Style:** Follow the existing patterns — `for` loops over COM collections (never `foreach`), `Marshal.ReleaseComObject` on every COM reference, `Debug.WriteLine` for error paths, no `RegexOptions.Compiled`

### Development Guidelines

- **Never edit `ThisAddIn.Designer.cs`** — it's auto-generated by Visual Studio
- **COM references must be released** in reverse acquisition order — the project won't build without proper cleanup
- **No `RegexOptions.Compiled`** — causes VSTO load failures (dynamic assembly at type-load time)
- **Test with DebugView** — add `Debug.WriteLine` traces with `[PS]` or `[CM]` prefixes for your code

## Known Limitations

- Classic Outlook only — new Outlook and OWA don't support COM/VSTO add-ins
- 64-bit Outlook required — 32-bit build targets a different interop assembly
- Sent Items and some system folders default to non-TableView views (CardView, Sent To) — the column is created on the folder, but you may need to switch to Messages view to see it
- The `MAPIFolder.CurrentView` property is read-only in the Outlook PIA — view switching is handled via the Explorer object

## License

MIT — see [LICENSE](LICENSE) for details.
