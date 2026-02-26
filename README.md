# 🚀 Winget-AutoUpdate Intune Remediation Scripts

This repository contains Microsoft Intune Remediation scripts that automatically detect and update Windows applications installed via Winget, with **Blacklist** (exclude apps) or **Whitelist** (include only) modes.

Based on update detection logic from [Winget-AutoUpdate (WAU)](https://github.com/Romanitho/Winget-AutoUpdate), adapted for Intune's two-script remediation model.

## 📋 Overview

The remediation consists of two PowerShell scripts:
- 🔍 **Detection Script** (`detection.ps1`) - Checks if any applications have available updates
- 🔧 **Remediation Script** (`remediation.ps1`) - Performs the actual application updates

## ✨ Features

- 🚫 **Blacklist** mode (default): update all apps except those listed
- ⬜ **Whitelist** mode: update only specified apps
- 🎯 Wildcard pattern matching for app IDs (e.g. `Mozilla.Firefox*`, `*Microsoft*`)
- 🔄 **Dual-scope detection**: queries both default (machine) and `--scope user` to find all updatable apps
- 🧠 **Smart scope upgrade**: upgrades without `--scope` first; falls back to `--scope user` only on "no applicable installer" for user-detected apps
- 📡 **Source health**: runs `winget source update` once at script start before the upgrade loop
- 📊 **Exit code mapping**: ~45 Winget exit codes categorised into `Success`, `RetryScope`, `RetryLater`, `Fail` (includes MSI 3010 = success)
- ⏳ **In-progress wait loop**: if another install is running, waits 30s and retries up to 5 times
- ⏸️ **Deferred updates**: transient errors (disk full, no network, app in use) return exit 0 so Intune retries next cycle
- 🔁 **Pending reboot detection**: logs when a reboot is pending on the system
- 📝 Comprehensive logging to `%ProgramData%\IntuneLogs\Scripts\Winget-AppUpdate\`
- 🌐 UTF-8 encoding handling for proper Unicode/CJK character support
- 💻 [PowerShell approved verbs](https://learn.microsoft.com/en-us/powershell/scripting/developer/cmdlet/approved-verbs-for-windows-powershell-commands) throughout
- 📏 camelCase variable naming

## 📦 Requirements

- 🪟 Windows 10/11 with Winget installed
- 🔗 Microsoft Entra joined or hybrid joined device
- 📱 Intune Management Extension installed
- 💻 PowerShell 5.1 or later
- 🔐 System context execution (scripts run as SYSTEM)

## 🚀 Deployment Guide

### Step 1: ⚙️ Configure List Mode

Open both `detection.ps1` and `remediation.ps1` and edit the Config section at the top.

**Blacklist mode** (default) - update all apps except those listed:

```powershell
$listMode = 'Blacklist'
$blacklistApps = @(
    'Microsoft.Edge*',
    'Microsoft.Teams*',
    'Microsoft.Office',
    'Mozilla.Firefox*',
    'Adobe.Acrobat.Reader.64-bit'
)
```

**Whitelist mode** - update only the listed apps:

```powershell
$listMode = 'Whitelist'
$whitelistApps = @(
    '7zip.7zip',
    'Google.Chrome',
    'Microsoft.DotNet*',
    'Microsoft.VCRedist*',
    'Notepad++.Notepad++'
)
```

⚠️ Both scripts must have the **same** `$listMode` and list configuration.

### Step 2: 📝 Configure Logging (Optional)

For production, disable debug logging:

```powershell
$logDebug = $false
```

### Step 3: ☁️ Upload Scripts to Intune

1. 🌐 Go to [Microsoft Intune Admin Center](https://endpoint.microsoft.com) > **Devices** > **Remediations**
2. ➕ Click **+ Create script package**
3. 📤 Upload `detection.ps1` as the detection script
4. 📤 Upload `remediation.ps1` as the remediation script
5. ⚙️ Configure:
   - **Run this script using the logged-on credentials**: No (system context)
   - **Enforce script signature check**: No (unless you sign the scripts)
   - **Run script in 64-bit PowerShell**: Yes
6. 👥 Assign to device groups
7. ⏰ Schedule (e.g. daily)

### Step 4: 📊 Monitor Execution

- 📈 **Intune Portal**: Devices > Remediations > select package > Device status
- 📂 **On-device logs**: `%ProgramData%\IntuneLogs\Scripts\Winget-AppUpdate\`
  - `detection.log`
  - `remediation.log`

## 🔧 How It Works

### Scope Detection Strategy

The scripts query Winget twice for available updates:

1. **`winget upgrade --source winget`** (no `--scope` flag) - returns machine-scoped apps by default
2. **`winget upgrade --source winget --scope user`** - returns additional user-scoped apps

Each app is tagged with its detected scope. Results are deduplicated by AppId, preferring the default (no-scope) entry when an app appears in both.

### Upgrade Strategy

When upgrading, the script always tries **without** `--scope` first (letting Winget auto-detect the correct scope). This works for the majority of apps. If Winget returns "no applicable installer for current scope" (`-1978335216`) **and** the app was originally detected as user-scoped, the script retries once with `--scope user`.

### Exit Code Categories

| Category | Meaning | Script Action |
|----------|---------|---------------|
| ✅ Success | App is at desired state | Count as success |
| 🔄 RetryScope | No installer for detected scope | Retry with `--scope user` if user-detected |
| ⏸️ RetryLater | Transient error (app in use, network, disk) | Return `$null` (deferred); Intune retries next cycle |
| ❌ Fail | Unrecoverable (policy, hash mismatch, unsupported) | Return `$false`; logged as error |
| ❓ Unknown | Unmapped code | Treated as Fail; hex value logged for lookup |

Notable: MSI exit code **3010** (`ERROR_SUCCESS_REBOOT_REQUIRED`) is treated as Success since the install actually completed.

### Three-State Result Model

`Update-Application` returns one of three values:

- ✅ `$true` - upgrade succeeded
- ❌ `$false` - hard failure (unrecoverable)
- ⏸️ `$null` - deferred (transient error; Intune retries next cycle)

Only hard failures (`$false`) cause the script to exit with code 1. Deferred updates exit 0 so Intune does not flag them as failures.

## 🚫 Blacklist / Whitelist Configuration

### Wildcard Support

Patterns use PowerShell's `-like` operator:

| Pattern | Matches | Does Not Match |
|---------|---------|----------------|
| `Mozilla.Firefox*` | `Mozilla.Firefox`, `Mozilla.Firefox.ESR` | `Other.Firefox` |
| `*Firefox` | `Mozilla.Firefox`, `Other.Firefox` | `Firefox.Standalone` |
| `Microsoft.VCLibs.*` | `Microsoft.VCLibs.140.00` | `Microsoft.VCLibs` |
| `Microsoft.Office` | `Microsoft.Office` (exact) | `Microsoft.Office365` |

### Finding App IDs

```powershell
winget list                          # All installed apps
winget search "AppName"              # Search by name
winget list | Select-String "Chrome"  # Filter installed apps
```

## 📝 Logging

### Log Tags

| Tag | Meaning |
|-----|---------|
| 🚀 Start/End | Script lifecycle |
| 🔍 Get | Data retrieval (Winget queries, parsing) |
| ⚙️ Run | Execution (upgrade commands, source update) |
| ℹ️ Info | Progress, configuration |
| ✅ Success | Successful operations |
| ❌ Error | Failures |
| 🐛 Debug | Verbose detail (disable with `$logDebug = $false`) |

### Example Log Output

```
2026-02-26 10:00:00 [  Start   ] ======== Script Started ========
2026-02-26 10:00:00 [  Info    ] ComputerName: PC01 | User: SYSTEM | Script: Winget-AppUpdate
2026-02-26 10:00:00 [  Run     ] Refreshing winget sources...
2026-02-26 10:00:02 [  Debug   ] Source refresh complete.
2026-02-26 10:00:02 [  Get     ] Checking for available updates (default + user scope)
2026-02-26 10:00:05 [  Debug   ] Default scope: found 3 app(s) with updates
2026-02-26 10:00:08 [  Debug   ] User scope: found 1 app(s) with updates
2026-02-26 10:00:08 [  Get     ] Found 4 unique apps with available updates
2026-02-26 10:00:08 [  Info    ] Using blacklist mode with 17 entries
2026-02-26 10:00:08 [  Get     ] Filtered to 3 apps requiring updates
2026-02-26 10:00:08 [  Run     ] Starting update process for 3 application(s)
2026-02-26 10:00:08 [  Info    ] [1/3] Processing: 7zip.7zip (24.08 -> 24.09)
2026-02-26 10:00:15 [  Success ] Successfully updated: 7zip.7zip
2026-02-26 10:00:18 [  Info    ] [2/3] Processing: Google.Chrome (131.0.0 -> 132.0.0) [user]
2026-02-26 10:00:25 [  Info    ] No applicable installer without scope; retrying Google.Chrome with --scope user.
2026-02-26 10:00:32 [  Success ] Successfully updated: Google.Chrome (with --scope user)
2026-02-26 10:00:35 [  Info    ] [3/3] Processing: SomeApp (1.0 -> 2.0)
2026-02-26 10:00:40 [  Info    ] Transient error for SomeApp (Application is currently running); will retry next Intune cycle.
2026-02-26 10:00:40 [  Success ] Successfully updated: 2 application(s)
2026-02-26 10:00:40 [  Info    ] Deferred (transient): 1 application(s) - SomeApp
2026-02-26 10:00:40 [  Success ] All updates completed or deferred successfully
2026-02-26 10:00:40 [  Info    ] Exit Code: 0
```

## 🔧 Troubleshooting

### ❌ Winget Not Found

Script fails with "Winget is not available or not working properly".

- ✅ Ensure `Microsoft.DesktopAppInstaller` is installed
- 📂 Check `%ProgramFiles%\WindowsApps\Microsoft.DesktopAppInstaller_*`
- 🔍 Run `winget --version` manually to verify
- 🔐 Ensure the script runs in system context

### 🔍 Updates Not Detected

Detection script returns exit code 0 but updates are available.

- ✅ Verify app IDs match Winget package IDs (`winget list`)
- 📝 Check blacklist/whitelist for typos
- 📋 Review detection logs for parsing errors
- ⚠️ Check if apps have "Unknown" versions (skipped by design)

### ⬆️ Updates Fail During Remediation

Remediation script returns exit code 1.

- 📋 Check remediation logs for the exit code and category
- 🔄 **RetryScope**: the script already retries with `--scope user` for user-detected apps; if it still fails the manifest may not support upgrade
- ⏸️ **RetryLater / deferred**: not counted as failures; Intune retries next cycle
- ❌ **Fail**: requires manual intervention (policy change, dependency install, etc.)
- ❓ **Unknown**: the log includes hex value for lookup on [MS return codes](https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md)

### 🌐 Encoding Issues

App names show garbled characters. The scripts set `[Console]::OutputEncoding` to UTF-8 before calling Winget. Ensure you are using the latest version of the scripts.

## 📄 License

This project follows the same MIT license as Winget-AutoUpdate.

## 🙏 Credits

Based on the excellent work by [Romanitho/Winget-AutoUpdate](https://github.com/Romanitho/Winget-AutoUpdate).
