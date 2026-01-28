# 🚀 Winget-AutoUpdate Intune Remediation Scripts

This repository contains Microsoft Intune Remediation scripts that replicate the update logic from [Winget-AutoUpdate (WAU)](https://github.com/Romanitho/Winget-AutoUpdate). These scripts automatically detect and update Windows applications installed via Winget, while respecting a configurable blacklist.

## 📋 Overview

The remediation consists of two PowerShell scripts:
- 🔍 **Detection Script** (`detection.ps1`): Checks if any applications have available updates
- 🔧 **Remediation Script** (`remediation.ps1`): Performs the actual application updates

## ✨ Features

- ✅ Replicates WAU's exact update detection and execution logic
- 🚫 Blacklist support (excluded apps) - configured directly in script
- 🎯 Wildcard pattern matching for app IDs
- 📝 Comprehensive logging to `%ProgramData%\IntuneLogs\Scripts\`
- 🌐 UTF-8 encoding handling for proper Unicode character support
- 💻 Uses PowerShell approved verbs
- 📏 Strict camelCase/PascalCase variable naming
- 🔐 System context execution
- ⚠️ Handles "Unknown" version apps gracefully
- 📊 Progress tracking with detailed logging

## 📦 Requirements

- 🪟 Windows 10/11 with Winget installed
- 🔗 Microsoft Entra joined or hybrid joined device
- 📱 Intune Management Extension installed
- 💻 PowerShell 5.1 or later
- 🔐 System context execution (scripts run as SYSTEM)

## 🚀 Deployment Guide

### Step 1: ⚙️ Configure Blacklist

Before deploying, configure which applications should **NOT** be updated automatically:

1. 📂 Open both `detection.ps1` and `remediation.ps1`
2. 🔍 Locate the **Config** section at the top of each script (around line 17)
3. ✏️ Edit the `$blacklistApps` array:

```powershell
# ---------------------------[ Config ]---------------------------
$blacklistApps = @(
    'Microsoft.Edge*',              # Excludes all Edge variants
    'Microsoft.Teams*',             # Excludes all Teams variants
    'Microsoft.Office',              # Exact match for Office
    'Mozilla.Firefox*',              # Excludes all Firefox channels
    'Adobe.Acrobat.Reader.64-bit'   # Specific version exclusion
)
```

⚠️ **Important**: Both scripts must have the **same** blacklist configuration!

### Step 2: 📝 Configure Logging (Optional)

By default, debug logging is enabled. For production, you may want to disable it:

1. 🔍 Locate the **Logging Setup** section (around line 28)
2. ⚙️ Set `$logDebug = $false` to reduce log verbosity:

```powershell
$logDebug = $false    # Set to $false for production (less verbose)
```

### Step 3: ☁️ Upload Scripts to Intune

1. 🌐 **Navigate to Intune Portal**:
   - Go to [Microsoft Intune Admin Center](https://endpoint.microsoft.com)
   - Navigate to **Devices** > **Remediations**

2. ➕ **Create New Remediation**:
   - Click **+ Create script package**
   - Enter a name: `Winget Application Updates`
   - (Optional) Add description: `Automatically updates Winget applications while respecting blacklist`

3. 🔍 **Upload Detection Script**:
   - In the **Detection script** section, click **Select a file**
   - Upload `detection.ps1`
   - ✅ Verify it appears in the editor

4. 🔧 **Upload Remediation Script**:
   - In the **Remediation script** section, click **Select a file**
   - Upload `remediation.ps1`
   - ✅ Verify it appears in the editor

5. ⚙️ **Configure Settings**:
   - **Run this script using the logged-on credentials**: **No** (system context)
   - **Enforce script signature check**: **No** (unless you sign the scripts)
   - **Run script in 64-bit PowerShell**: **Yes** (recommended)

6. 👥 **Assign to Devices/Groups**:
   - Click **Next** to go to **Scope tags** (optional)
   - Click **Next** to go to **Assignments**
   - Click **+ Select groups to include**
   - Choose your target device groups (e.g., "All Windows Devices", specific security groups)
   - (Optional) Add exclusion groups
   - Click **Next**

7. ✅ **Review and Create**:
   - Review all settings
   - Click **Create**

### Step 4: ⏰ Schedule Remediation (Optional)

By default, Intune runs remediations based on the detection script results. You can configure a schedule:

1. 📂 Go to **Devices** > **Remediations**
2. 🎯 Select your remediation package
3. ⚙️ Click **Properties** > **Schedule**
4. 🔧 Configure:
   - **Run detection script every**: Choose frequency (e.g., Daily, Weekly)
   - **Run remediation script if detection returns non-zero**: **Yes**

### Step 5: 📊 Monitor Execution

1. 📈 **View Remediation Status**:
   - Go to **Devices** > **Remediations**
   - Select your remediation package
   - View **Device status** tab to see execution results

2. 📝 **Check Logs on Devices**:
   - Logs are written to: `%ProgramData%\IntuneLogs\Scripts\Winget-AppUpdate-Blacklist\`
   - Detection logs: `detection.log`
   - Remediation logs: `remediation.log`

## 🎭 Script Behavior

### 🔍 Detection Script (`detection.ps1`)

**Purpose**: Determines if any applications require updates.

**Execution Flow**:
1. 🚀 **Initialize**: Sets up logging and script metadata
2. 🔍 **Check Winget Availability**: 
   - Locates `winget.exe` in system or user context
   - Verifies Winget is working by checking version
   - Exits with code 1 if Winget is unavailable
3. 📦 **Get Available Updates**:
   - Executes `winget upgrade --source winget` (no accept flags for listing)
   - Parses tabular output using WAU's exact column-position-based parsing
   - Handles UTF-8 encoding properly (fixes encoding issues)
   - Replaces ellipsis characters (`\u2026`) for proper column alignment
   - Filters out apps with "Unknown" versions
   - Extracts: App Name, App ID, Current Version, Available Version
4. 🚫 **Filter by Blacklist**:
   - Compares each app ID against blacklist patterns
   - Supports wildcard matching (e.g., `Microsoft.Edge*`)
   - Removes blacklisted apps from update list
5. ✅ **Return Result**:
   - **Exit Code 0**: No updates needed (all apps up to date or filtered out)
   - **Exit Code 1**: Updates available (remediation required)

**Example Output** (when updates available):
```
[  Get     ] Found 4 apps with available updates
[  Info    ] Using blacklist with 15 entries
[  Get     ] Filtered to 4 apps requiring updates
[  Info    ] Apps requiring updates:
[  Info    ]   - Microsoft.DotNet.DesktopRuntime.8: 8.0.22 -> 8.0.23
[  Info    ]   - Microsoft.VCRedist.2015+.x64: 14.42.34433.0 -> 14.50.35719.0
[  Success ] Detection complete: 4 app(s) need updating
[  Info    ] Exit Code: 1
```

**Example Output** (no updates):
```
[  Get     ] Found 0 apps with available updates
[  Success ] No updates available - all apps are up to date
[  Info    ] Exit Code: 0
```

### 🔧 Remediation Script (`remediation.ps1`)

**Purpose**: Performs the actual application updates.

**Execution Flow**:
1. 🚀 **Initialize**: Sets up logging and script metadata
2. 🔍 **Check Winget Availability**: Same as detection script
3. 📦 **Get Available Updates**: Same parsing logic as detection script
4. 🚫 **Filter by Blacklist**: Same filtering logic as detection script
5. ⬆️ **Update Each Application**:
   - Iterates through filtered update list
   - For each app:
     - Logs progress: `[X/Total] Processing: AppId (current -> available)`
     - Executes: `winget upgrade --id <AppId> -e --accept-package-agreements --accept-source-agreements -s winget -h`
       - `-e`: Exact match
       - `-h`: Silent mode
       - `-s winget`: Source specification
     - Filters output to remove progress indicators
     - Checks exit code for success/failure
     - Waits 2 seconds between updates (to avoid overwhelming system)
6. ✅ **Return Result**:
   - **Exit Code 0**: All updates successful
   - **Exit Code 1**: One or more updates failed

**Example Output** (successful run):
```
[  Run     ] Starting update process for 4 application(s)
[  Info    ] [1/4] Processing: Microsoft.DotNet.DesktopRuntime.8 (8.0.22 -> 8.0.23)
[  Success ] Successfully updated: Microsoft.DotNet.DesktopRuntime.8
[  Info    ] [2/4] Processing: Microsoft.VCRedist.2015+.x64 (14.42.34433.0 -> 14.50.35719.0)
[  Success ] Successfully updated: Microsoft.VCRedist.2015+.x64
[  Success ] Successfully updated: 4 application(s)
[  Success ] All updates completed successfully
[  Info    ] Exit Code: 0
```

**Example Output** (partial failure):
```
[  Run     ] Starting update process for 3 application(s)
[  Info    ] [1/3] Processing: App1 (1.0.0 -> 2.0.0)
[  Success ] Successfully updated: App1
[  Info    ] [2/3] Processing: App2 (1.0.0 -> 2.0.0)
[  Error   ] Failed to update App2 - Exit code: -1978335212
[  Error   ] Error: No installed package found matching input criteria.
[  Info    ] [3/3] Processing: App3 (1.0.0 -> 2.0.0)
[  Success ] Successfully updated: App3
[  Success ] Successfully updated: 2 application(s)
[  Error   ] Failed to update: 1 application(s)
[  Error   ] Failed apps: App2
[  Info    ] Exit Code: 1
```

### 🔄 How Detection and Remediation Work Together

1. 🔍 **Intune runs Detection Script**:
   - Executes `detection.ps1` on schedule or on-demand
   - Script checks for available updates
   - Returns exit code based on results

2. ⬆️ **If Detection Returns Exit Code 1**:
   - Intune recognizes updates are needed
   - Intune automatically runs `remediation.ps1`
   - Remediation script performs the updates

3. ✅ **If Detection Returns Exit Code 0**:
   - Intune recognizes no updates needed
   - Remediation script is **not** executed
   - Process repeats on next schedule

4. 📊 **Remediation Result**:
   - If remediation succeeds (exit code 0), Intune marks as successful
   - If remediation fails (exit code 1), Intune may retry based on policy
   - Detection will run again on next schedule to check for remaining updates

## 🚫 Blacklist Configuration

### ⚙️ Basic Configuration

Edit the `$blacklistApps` array in the **Config** section at the top of both scripts:

```powershell
# ---------------------------[ Config ]---------------------------
$blacklistApps = @(
    'Microsoft.Edge*',              # Excludes all Edge variants
    'Microsoft.Teams*',             # Excludes all Teams variants  
    'Microsoft.Office',              # Exact match only
    'Microsoft.OneDrive',
    'Mozilla.Firefox*',              # Excludes all Firefox channels
    'Opera.Opera*',
    'Brave.Brave*',
    'Adobe.Acrobat.Reader.64-bit'   # Specific version
)
```

### 🎯 Wildcard Support

The blacklist supports three types of wildcard patterns:

1. 🔚 **End Wildcard** (`AppName*`):
   - `Mozilla.Firefox*` matches:
     - `Mozilla.Firefox`
     - `Mozilla.Firefox.ESR`
     - `Mozilla.Firefox.DeveloperEdition`
   - Does NOT match: `Other.Firefox`

2. 🔜 **Start Wildcard** (`*AppName`):
   - `*Firefox` matches:
     - `Mozilla.Firefox`
     - `Other.Firefox`
   - Does NOT match: `Firefox.Standalone`

3. 🔀 **Middle Wildcard** (`App*Name`):
   - `Fire*fox` matches:
     - `Firefox`
     - `Firebirdfox`
   - Uses regex pattern matching

4. ✅ **Exact Match** (no wildcard):
   - `Microsoft.Office` matches only `Microsoft.Office`
   - Case-sensitive matching

### 🔍 Finding App IDs

To find the correct App ID for blacklisting:

1. 📋 **Using Winget CLI**:
   ```powershell
   winget list
   ```
   Look for the "Id" column

2. 🔎 **Search for App**:
   ```powershell
   winget search "AppName"
   ```
   The "Id" column shows the package identifier

3. 📦 **Check Installed Apps**:
   ```powershell
   winget list | Select-String "AppName"
   ```

## 📝 Logging

### 📂 Log Locations

Logs are written to:
- 🔍 **Detection**: `%ProgramData%\IntuneLogs\Scripts\Winget-AppUpdate-Blacklist\detection.log`
- 🔧 **Remediation**: `%ProgramData%\IntuneLogs\Scripts\Winget-AppUpdate-Blacklist\remediation.log`

### 🏷️ Log Tags

The scripts use structured logging with the following tags:

- 🚀 **Start/End**: Script lifecycle events
- 🔍 **Get**: Information retrieval operations (Winget queries, parsing)
- ⚙️ **Run**: Execution operations (update commands)
- ℹ️ **Info**: General information (progress, configuration)
- ✅ **Success**: Successful operations (updates completed)
- ❌ **Error**: Errors and failures
- 🐛 **Debug**: Detailed debugging information (disabled by default in production)

### ⚙️ Logging Configuration

Control logging behavior in the **Logging Setup** section:

```powershell
$log           = $true      # Master logging switch
$logDebug      = $false     # Debug logging (set to $true for troubleshooting)
$logGet        = $true      # [Get] tag logging
$logRun        = $true      # [Run] tag logging
$enableLogFile = $true      # File logging enabled
```

### Example Log Entry

```
2026-01-28 21:18:20 [  Start   ] ======== Script Started ========
2026-01-28 21:18:20 [  Info    ] ComputerName: DESKTOP-P611MM0 | User: DESKTOP-P611MM0$ | Script: Winget-AppUpdate-Blacklist
2026-01-28 21:18:20 [  Get     ] Checking Winget availability
2026-01-28 21:18:20 [  Success ] Winget version: 1.12.460
2026-01-28 21:18:20 [  Get     ] Checking for available updates
2026-01-28 21:18:21 [  Get     ] Found 4 apps with available updates
2026-01-28 21:18:21 [  Info    ] Using blacklist with 15 entries
2026-01-28 21:18:21 [  Get     ] Filtered to 4 apps requiring updates
2026-01-28 21:18:21 [  Info    ] Apps requiring updates:
2026-01-28 21:18:21 [  Info    ]   - Microsoft.DotNet.DesktopRuntime.8: 8.0.22 -> 8.0.23
2026-01-28 21:18:21 [  Success ] Detection complete: 4 app(s) need updating
2026-01-28 21:18:21 [  Info    ] Script execution time: 00:00:00.83
2026-01-28 21:18:21 [  Info    ] Exit Code: 1
2026-01-28 21:18:21 [  End     ] ======== Script Completed ========
```

## 🔧 How It Works (Technical Details)

### 🔍 Update Detection

1. 💻 **Winget Command**: `winget upgrade --source winget`
   - Lists all installed packages with available updates
   - Output is tabular format with columns: Name, Id, Version, Available

2. 📊 **Output Parsing**:
   - Uses WAU's exact parsing logic
   - Finds header separator line (`-----`)
   - Calculates column positions based on header
   - Handles non-Latin characters (CJK) with declination calculation
   - Replaces ellipsis characters for proper alignment
   - Extracts data using substring operations with TrimEnd()

3. 🚫 **Filtering**:
   - Skips apps with "Unknown" versions
   - Skips apps where current version equals available version
   - Applies blacklist filtering with wildcard support

### ⬆️ Update Execution

1. 💻 **Winget Command**: `winget upgrade --id <AppId> -e --accept-package-agreements --accept-source-agreements -s winget -h`
   - `--id <AppId>`: Specific app to update
   - `-e`: Exact match (prevents multiple matches)
   - `--accept-package-agreements`: Auto-accept package agreements
   - `--accept-source-agreements`: Auto-accept source agreements
   - `-s winget`: Specify source
   - `-h`: Silent mode (no user interaction)

2. 🔍 **Output Filtering**:
   - Removes lines starting with space (progress indicators)
   - Extracts meaningful error messages
   - Checks exit code for success/failure

3. ⚠️ **Error Handling**:
   - Logs specific error messages
   - Continues with next app if one fails
   - Returns appropriate exit code based on overall success

## 🔄 Differences from WAU

- ⏰ **No scheduled tasks**: Intune handles scheduling
- 🔔 **No notifications**: Intune provides remediation status
- 👤 **No user context**: Currently system context only
- 🔌 **No mods support**: Simplified version without mod scripts
- 🔄 **No self-update**: Intune manages script updates
- 📝 **Blacklist in script**: Configured directly in script Config section (no external files)
- 🌐 **UTF-8 encoding fix**: Properly handles Unicode characters to prevent parsing issues

## 🔧 Troubleshooting

### ❌ Winget Not Found

**Symptoms**: Script fails with "Winget is not available or not working properly"

**Solutions**:
- ✅ Ensure Microsoft.DesktopAppInstaller is installed
- 📂 Check `%ProgramFiles%\WindowsApps\Microsoft.DesktopAppInstaller_*`
- 🔍 Verify Winget is accessible: Run `winget --version` manually
- 🔐 Check if running in system context (required)

### 🔍 Updates Not Detected

**Symptoms**: Detection script returns exit code 0 but updates are available

**Solutions**:
- ✅ Verify app IDs match Winget package IDs exactly (use `winget list`)
- 📝 Check blacklist array in script Config section for typos
- 📋 Review detection script logs for parsing errors
- 📦 Ensure apps are installed via Winget (not MSI/EXE installers)
- ⚠️ Check if apps have "Unknown" versions (these are skipped)

### ⬆️ Updates Fail During Remediation

**Symptoms**: Remediation script returns exit code 1

**Solutions**:
- 📋 Check remediation script logs for specific errors
- ✅ Verify apps are installed via Winget
- 👤 Some apps may require user interaction (not supported in silent mode)
- 🔄 Check for pending reboots (may prevent updates)
- ✅ Verify app IDs are correct (use exact match with `-e` flag)
- 💾 Check disk space availability

### 🌐 Encoding Issues (Garbled Characters)

**Symptoms**: App names or IDs show garbled characters like `ΓÇª` or `ª`

**Solutions**:
- ✅ Scripts now handle UTF-8 encoding automatically
- 🔄 Ensure you're using the latest version of the scripts
- 🔧 The encoding fix sets `[Console]::OutputEncoding` to UTF-8 before calling Winget

### 🚫 Blacklist Not Working

**Symptoms**: Apps in blacklist are still being updated

**Solutions**:
- ✅ Verify blacklist is identical in both scripts
- 🔍 Check for typos in app IDs
- 🧪 Test wildcard patterns manually
- 📋 Review logs for "App excluded (in blacklist)" messages
- ✅ Ensure wildcard syntax is correct (`AppName*`, not `AppName.*`)

## 💡 Best Practices

1. 🧪 **Test Before Production**:
   - Deploy to a test group first
   - Monitor logs for a few days
   - Verify blacklist is working correctly

2. 🔄 **Keep Blacklists Synchronized**:
   - Always update both scripts with the same blacklist
   - Document why apps are blacklisted

3. 📊 **Monitor Logs Regularly**:
   - Check logs weekly for errors
   - Review update success rates
   - Identify apps that frequently fail updates

4. 🛡️ **Start Conservative**:
   - Begin with a larger blacklist
   - Gradually reduce blacklist as confidence grows
   - Monitor for unwanted updates

5. ⏰ **Schedule Appropriately**:
   - Daily detection is recommended
   - Avoid peak usage hours
   - Consider network bandwidth

## 📄 License

This project follows the same MIT license as Winget-AutoUpdate.

## 🙏 Credits

Based on the excellent work by [Romanitho/Winget-AutoUpdate](https://github.com/Romanitho/Winget-AutoUpdate)
