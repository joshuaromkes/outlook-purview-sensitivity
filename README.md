# Outlook Purrrview Sensitivity

Outlook COM add-in (VSTO) that adds a **"Purview Label" column** to your inbox and folders, showing Microsoft Purview sensitivity labels at a glance.

## The Problem

Microsoft Purview sensitivity labels work great — but Outlook's built-in "Sensitivity" column shows the legacy Normal/Personal/Private/Confidential flag, NOT your Purview labels. Users can only see labels one email at a time by opening the message.

## The Fix

This add-in reads the actual Purview sensitivity label from each email's metadata (`msip_labels` MAPI property) and exposes it as a sortable, filterable column in every Outlook folder view.

| | Before | After |
|---|---|---|
| Inbox column | "Normal" (wrong) | "PII High" / "PII Med" / "PII Low" |
| Sort by label | ❌ | ✅ |
| Filter by label | ❌ | ✅ |
| At-a-glance visibility | Open each email | Column visible in folder view |

## How It Works

1. Reads the `msip_labels` x-header from every email via MAPI `PropertyAccessor`
2. Parses the human-readable label name (e.g., "PII High")
3. Injects a custom "PurviewLabel" user-defined field into every folder
4. Stamps each email with its label so it appears as a column

No external APIs, no cloud dependency — labels are embedded in the email metadata.

## Compatibility

- Outlook 2016, 2019, 2021, Microsoft 365 (classic, desktop)
- Windows 10, Windows 11
- .NET Framework 4.8
- Classic Outlook only (new Outlook and OWA not supported — COM add-ins don't run there)

## Install

### For End Users (via Intune)

Download the MSI from [Releases](https://github.com/joshuaromkes/outlook-purrrview-sensitivity/releases) and deploy as a Win32 app. See [docs/intune-deploy.md](docs/intune-deploy.md).

### For Developers

```powershell
# Build in Visual Studio 2022 Community (free)
# Requires: Office/SharePoint development workload

# Or register for debugging:
.\installer\register.ps1
```

## License

MIT — open source, use it anywhere.
