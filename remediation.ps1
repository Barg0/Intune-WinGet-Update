# ---------------------------[ Config ]---------------------------
# ListMode: 'Blacklist' = update all except listed apps; 'Whitelist' = update only listed apps
$listMode = 'Blacklist'

# Blacklist: apps EXCLUDED from updates (used when listMode = 'Blacklist')
$blacklistApps = @(
    'Microsoft.Edge*',
    'Microsoft.Teams*',
    'Microsoft.Office',
    'Microsoft.OneDrive',
    'Microsoft.AppInstaller',
    'Microsoft.RemoteDesktopClient',
    'Microsoft.GlobalSecureAccessClient',
    'Microsoft.VCLibs.*',
    'Fortinet.FortiClientVPN',
    'Mozilla.Firefox*',
    'Opera.Opera*',
    'TeamViewer.TeamViewer*',
    'Brave.Brave*',
    'KeePassXCTeam.KeePassXC',
    'Microsoft.WindowsTerminal',
    'Adobe.Acrobat.Pro',
    'Adobe.CreativeCloud',
    'Adobe.Acrobat.Reader.32-bit',
    'Adobe.Acrobat.Reader.64-bit',
    'Microsoft.PowerShell',
    'Lenovo.SUHelper'
)

# Whitelist: apps INCLUDED in updates (used when listMode = 'Whitelist'). Wildcards supported.
$whitelistApps = @(
    '7zip.7zip',
    'Google.Chrome',
    'Microsoft.DotNet*',
    'Microsoft.VCRedist*',
    'Notepad++.Notepad++'
)

# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Script Name ]---------------------------
$scriptName  = 'Winget-AppUpdate'
$logFileName = "remediation.log"

# ---------------------------[ Logging Setup ]---------------------------
$log           = $true
$logDebug      = $true
$logGet        = $true
$logRun        = $true
$enableLogFile = $true

$logFileDirectory = "$env:ProgramData\IntuneLogs\Scripts\$scriptName"
$logFile          = "$logFileDirectory\$logFileName"

# ---------------------------[ Winget Exit Code Map ]---------------------------
# Only codes that can realistically fire during winget upgrade / install.
# Sources:
#   https://kb.filewave.com/books/microsoft-windows-package-manager-winget/page/troubleshooting-errors-with-winget
#   https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md
#
# Categories drive the retry engine:
#   Success    - Desired state reached. No action needed.
#   RetryScope - No applicable installer for scope; retry with --scope user (if user-scoped app).
#   RetryLater - Transient (app in use, disk full, reboot, network). Defer to next Intune cycle.
#   Fail       - Unrecoverable in automation (policy block, hash mismatch, unsupported).
function Get-WingetExitCodeInfo {
    [CmdletBinding()]
    param([int]$ExitCode)
    $codeMap = @{
        # ── Success ──
        0              = @{ Category = 'Success';    Description = 'Success' }
        3010           = @{ Category = 'Success';    Description = 'Success (reboot required to complete)' }              # MSI ERROR_SUCCESS_REBOOT_REQUIRED
        -1978335135    = @{ Category = 'Success';    Description = 'Package already installed' }                           # 0x8A150061
        -1978334963    = @{ Category = 'Success';    Description = 'Another version already installed' }                   # 0x8A15010D
        -1978334962    = @{ Category = 'Success';    Description = 'Higher version already installed (downgrade)' }        # 0x8A15010E
        -1978334965    = @{ Category = 'Success';    Description = 'Reboot initiated to finish installation' }             # 0x8A15010B
        -1978335189    = @{ Category = 'Success';    Description = 'No applicable update found' }                          # 0x8A15002B

        # ── RetryScope: retry with --scope user (for user-detected apps) ──
        -1978335216    = @{ Category = 'RetryScope'; Description = 'No applicable installer for current scope' }           # 0x8A150010

        # ── RetryLater: transient – defer to next Intune cycle ──
        -1978334975    = @{ Category = 'RetryLater'; Description = 'Application is currently running' }                    # 0x8A150101
        -1978334974    = @{ Category = 'RetryLater'; Description = 'Another installation in progress' }                    # 0x8A150102
        -1978334973    = @{ Category = 'RetryLater'; Description = 'One or more file is in use' }                          # 0x8A150103
        -1978334971    = @{ Category = 'RetryLater'; Description = 'Disk full' }                                           # 0x8A150105
        -1978334970    = @{ Category = 'RetryLater'; Description = 'Insufficient memory' }                                 # 0x8A150106
        -1978334969    = @{ Category = 'RetryLater'; Description = 'No network connectivity' }                             # 0x8A150107
        -1978334967    = @{ Category = 'RetryLater'; Description = 'Reboot required to finish installation' }              # 0x8A150109
        -1978334966    = @{ Category = 'RetryLater'; Description = 'Reboot required then try again' }                      # 0x8A15010A
        -1978334959    = @{ Category = 'RetryLater'; Description = 'Application in use by another application' }           # 0x8A150111
        -1978335123    = @{ Category = 'RetryLater'; Description = 'Service busy or unavailable' }                         # 0x8A15006D
        -1978335224    = @{ Category = 'RetryLater'; Description = 'Download failed' }                                     # 0x8A150008
        -1978335186    = @{ Category = 'RetryLater'; Description = 'Download size mismatch' }                              # 0x8A15002E
        -1978335126    = @{ Category = 'RetryLater'; Description = 'Application shutdown signal received' }                # 0x8A15006A
        -1978335125    = @{ Category = 'RetryLater'; Description = 'Failed to download dependencies' }                     # 0x8A15006B

        # ── Fail: unrecoverable without human intervention ──
        -1978335231    = @{ Category = 'Fail'; Description = 'Internal error' }                                            # 0x8A150001
        -1978335230    = @{ Category = 'Fail'; Description = 'Invalid command line arguments' }                            # 0x8A150002
        -1978335229    = @{ Category = 'Fail'; Description = 'Command failed' }                                            # 0x8A150003
        -1978335228    = @{ Category = 'Fail'; Description = 'Opening manifest failed' }                                   # 0x8A150004
        -1978335226    = @{ Category = 'Fail'; Description = 'ShellExecute install failed' }                               # 0x8A150006
        -1978335225    = @{ Category = 'Fail'; Description = 'Manifest version higher than supported; update winget' }     # 0x8A150007
        -1978335221    = @{ Category = 'Fail'; Description = 'Configured source information is corrupt' }                  # 0x8A15000B
        -1978335217    = @{ Category = 'Fail'; Description = 'Source data missing' }                                       # 0x8A15000F
        -1978335215    = @{ Category = 'Fail'; Description = 'Installer hash does not match manifest' }                    # 0x8A150011
        -1978335212    = @{ Category = 'Fail'; Description = 'No packages found' }                                         # 0x8A150014
        -1978335210    = @{ Category = 'Fail'; Description = 'Multiple packages found matching criteria' }                 # 0x8A150016
        -1978335209    = @{ Category = 'Fail'; Description = 'No manifest found matching criteria' }                       # 0x8A150017
        -1978335207    = @{ Category = 'Fail'; Description = 'Command requires administrator privileges' }                 # 0x8A150019
        -1978335205    = @{ Category = 'Fail'; Description = 'Microsoft Store client blocked by policy' }                  # 0x8A15001B
        -1978335204    = @{ Category = 'Fail'; Description = 'Microsoft Store app blocked by policy' }                     # 0x8A15001C
        -1978335187    = @{ Category = 'Fail'; Description = 'Installer failed security check' }                           # 0x8A15002D
        -1978335174    = @{ Category = 'Fail'; Description = 'Blocked by Group Policy' }                                   # 0x8A15003A
        -1978335169    = @{ Category = 'Fail'; Description = 'Source data corrupted or tampered' }                         # 0x8A15003F
        -1978335163    = @{ Category = 'Fail'; Description = 'Failed to open source' }                                     # 0x8A150045
        -1978335157    = @{ Category = 'Fail'; Description = 'Failed to open one or more sources' }                        # 0x8A15004B
        -1978335159    = @{ Category = 'Fail'; Description = 'MSI install failed' }                                        # 0x8A150049
        -1978335153    = @{ Category = 'Fail'; Description = 'Upgrade version not newer than installed' }                   # 0x8A15004F
        -1978335146    = @{ Category = 'Fail'; Description = 'Installer prohibits elevation' }                              # 0x8A150056
        -1978335138    = @{ Category = 'Fail'; Description = 'Pinned certificate mismatch' }                               # 0x8A15005E
        -1978335128    = @{ Category = 'Fail'; Description = 'Package has a pin that prevents upgrade' }                    # 0x8A150068
        -1978335122    = @{ Category = 'Fail'; Description = 'Package is a stub; full package needed' }                     # 0x8A150069
        -1978334972    = @{ Category = 'Fail'; Description = 'Missing dependency on system' }                               # 0x8A150104
        -1978334968    = @{ Category = 'Fail'; Description = 'Installation error; contact support' }                        # 0x8A150108
        -1978334961    = @{ Category = 'Fail'; Description = 'Blocked by organization policy' }                             # 0x8A15010F
        -1978334960    = @{ Category = 'Fail'; Description = 'Failed to install package dependencies' }                     # 0x8A150110
        -1978334958    = @{ Category = 'Fail'; Description = 'Invalid parameter' }                                          # 0x8A150112
        -1978334957    = @{ Category = 'Fail'; Description = 'Package not supported by system' }                            # 0x8A150113
        -1978334956    = @{ Category = 'Fail'; Description = 'Installer does not support upgrading existing package' }      # 0x8A150114
    }
    if ($codeMap.ContainsKey($ExitCode)) { return $codeMap[$ExitCode] }
    return @{ Category = 'Unknown'; Description = "Unmapped exit code $ExitCode (hex: 0x$( '{0:X8}' -f [uint32]$ExitCode ))" }
}

if ($enableLogFile -and -not (Test-Path -Path $logFileDirectory)) {
    New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null
}

# ---------------------------[ Logging Function ]---------------------------
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [string]$Tag = "Info"
    )

    if (-not $log) { return }

    if (($Tag -eq "Debug") -and (-not $logDebug)) { return }
    if (($Tag -eq "Get")   -and (-not $logGet))   { return }
    if (($Tag -eq "Run")   -and (-not $logRun))   { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList   = @("Start","Get","Run","Info","Success","Error","Debug","End")
    $rawTag    = $Tag.Trim()

    if ($tagList -contains $rawTag) {
        $rawTag = $rawTag.PadRight(7)
    }
    else {
        $rawTag = "Error  "
    }

    $color = switch ($rawTag.Trim()) {
        "Start"   { "Cyan" }
        "Get"     { "Blue" }
        "Run"     { "Magenta" }
        "Info"    { "Yellow" }
        "Success" { "Green" }
        "Error"   { "Red" }
        "Debug"   { "DarkYellow" }
        "End"     { "Cyan" }
        default   { "White" }
    }

    $logMessage = "$timestamp [  $rawTag ] $Message"

    if ($enableLogFile) {
        try {
            Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
        }
        catch { }
    }

    Write-Host "$timestamp " -NoNewline
    Write-Host "[  " -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host " ] " -NoNewline -ForegroundColor White
    Write-Host "$Message"
}

# ---------------------------[ Exit Function ]---------------------------
function Complete-Script {
    param([int]$ExitCode)

    $scriptEndTime = Get-Date
    $duration      = $scriptEndTime - $scriptStartTime

    Write-Log "Script execution time: $($duration.ToString('hh\:mm\:ss\.ff'))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Script Completed ========" -Tag "End"

    exit $ExitCode
}

# ---------------------------[ Script Start ]---------------------------
Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"

# ---------------------------[ Winget Path Resolver ]---------------------------
function Get-WingetPath {
    [CmdletBinding()]
    param()

    $wingetBase = "$env:ProgramW6432\WindowsApps"
    $patterns   = @(
        'Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe'
        'Microsoft.DesktopAppInstaller_*_arm64__8wekyb3d8bbwe'
    )

    try {
        foreach ($pattern in $patterns) {
            $wingetFolders = Get-ChildItem -Path $wingetBase -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like $pattern }

            if (-not $wingetFolders) { continue }

            $candidates = foreach ($folder in $wingetFolders) {
                $exePath = Join-Path $folder.FullName 'winget.exe'
                if (-not (Test-Path -LiteralPath $exePath)) { continue }
                try {
                    $ver = (Get-Item -LiteralPath $exePath -ErrorAction Stop).VersionInfo.FileVersionRaw
                } catch {
                    $ver = $null
                }
                [PSCustomObject]@{
                    Path         = $exePath
                    FileVersion  = $ver
                    CreationTime = $folder.CreationTime
                }
            }

            if (-not $candidates) { continue }

            $latest = $candidates |
                Sort-Object { $_.FileVersion }, CreationTime -Descending |
                Select-Object -First 1

            if ($latest.Path) {
                return $latest.Path
            }
        }

        $userPath = "$env:LocalAppData\Microsoft\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe"
        if (Test-Path -LiteralPath $userPath) {
            return $userPath
        }

        Write-Log "Failed to detect Winget installation: no x64/arm64 folder or winget.exe found." -Tag "Error"
        throw "Winget not found in system or user context"
    }
    catch {
        if ($_.Exception.Message -notlike 'Winget not found*') {
            Write-Log "Failed to detect Winget installation: $_" -Tag "Error"
        }
        throw "Winget not found in system or user context"
    }
}

# ---------------------------[ Test Pending Reboot ]---------------------------
function Test-PendingReboot {
    try {
        $paths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        )
        foreach ($p in $paths) { if (Test-Path $p) { return $true } }
        $pn = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
        if ($pn -and $pn.PendingFileRenameOperations) { return $true }
        return $false
    }
    catch { return $false }
}

# ---------------------------[ Test Winget Function ]---------------------------
function Test-Winget {
    [CmdletBinding()]
    param()

    Write-Log "Checking Winget availability" -Tag "Get"

    try {
        $wingetPath = Get-WingetPath
        $rawOutput = & $wingetPath -v 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            $versionLine = $rawOutput | Where-Object { $_ -and ($_ -match '\d+\.\d+') } | Select-Object -First 1
            if ($versionLine -and $versionLine -match '(\d+\.\d+(?:\.\d+)?(?:\.\d+)?)') {
                Write-Log "Winget version: $($matches[1])" -Tag "Success"
            }
            else {
                Write-Log "Winget is available (execution successful)" -Tag "Success"
            }
            return $true
        }
        else {
            Write-Log "Winget execution failed with exit code: $exitCode" -Tag "Error"
            $errorOutput = $rawOutput | Where-Object { $_ -and $_ -notmatch '^\s*$' } | Select-Object -First 3
            if ($errorOutput) {
                Write-Log "Error details: $($errorOutput -join '; ')" -Tag "Debug"
            }
            return $false
        }
    }
    catch {
        Write-Log "Error testing Winget: $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Test-AppMatch ]---------------------------
function Test-AppMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [string[]]$PatternList
    )

    foreach ($pattern in $PatternList) {
        if ($AppId -like $pattern) {
            return $true
        }
    }

    return $false
}

# ---------------------------[ Parse Winget Upgrade Output ]---------------------------
function Parse-WingetUpgradeOutput {
    [CmdletBinding()]
    param(
        [string]$RawOutput,
        [string]$Scope
    )

    $updates = @()
    $unknownCount = 0

    if (-not ($RawOutput -match "-----")) {
        return $updates
    }

    $lines = $RawOutput.Split([Environment]::NewLine) | Where-Object { $_ }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lines[$i] = $lines[$i] -replace "[\u2026]", " "
    }

    $fl = 0
    while ($fl -lt $lines.Count -and -not $lines[$fl].StartsWith("-----")) { $fl++ }
    if ($fl -ge $lines.Count) { return $updates }
    $fl = $fl - 1
    if ($fl -lt 0) { return $updates }

    $index = $lines[$fl] -split '(?<=\s)(?!\s)'
    if ($index.Count -lt 3) { return $updates }

    $idStart = $($index[0] -replace '[\u4e00-\u9fa5]', '**').Length
    $versionStart = $idStart + $($index[1] -replace '[\u4e00-\u9fa5]', '**').Length
    $availableStart = $versionStart + $($index[2] -replace '[\u4e00-\u9fa5]', '**').Length

    for ($i = $fl + 2; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line.StartsWith("-----")) {
            $fl = $i - 1
            $index = $lines[$fl] -split '(?<=\s)(?!\s)'
            $idStart = $($index[0] -replace '[\u4e00-\u9fa5]', '**').Length
            $versionStart = $idStart + $($index[1] -replace '[\u4e00-\u9fa5]', '**').Length
            $availableStart = $versionStart + $($index[2] -replace '[\u4e00-\u9fa5]', '**').Length
            continue
        }
        if ($line -match "\w\.\w") {
            $nameDeclination = $($line.Substring(0, $idStart) -replace '[\u4e00-\u9fa5]', '**').Length - $line.Substring(0, $idStart).Length
            $appName = $line.Substring(0, $idStart - $nameDeclination).TrimEnd()
            $appId = $line.Substring($idStart - $nameDeclination, $versionStart - $idStart).TrimEnd()
            $currentVersion = $line.Substring($versionStart - $nameDeclination, $availableStart - $versionStart).TrimEnd()
            $availableVersion = $line.Substring($availableStart - $nameDeclination).TrimEnd()
            if ($currentVersion -eq "Unknown" -or $availableVersion -eq "Unknown") {
                $unknownCount++
                continue
            }
            if ($currentVersion -ne $availableVersion) {
                $updates += @{
                    AppId            = $appId
                    AppName          = $appName
                    CurrentVersion   = $currentVersion
                    AvailableVersion = $availableVersion
                    Scope            = $Scope
                }
            }
        }
    }
    return $updates
}

# ---------------------------[ Get Available Updates ]---------------------------
# Two calls: without --scope (machine apps by default) + with --scope user (additional user apps)
function Get-AvailableUpdates {
    [CmdletBinding()]
    param()

    Write-Log "Checking for available updates (default + user scope)" -Tag "Get"
    $wingetPath = Get-WingetPath

    try {
        $previousOutputEncoding = [Console]::OutputEncoding
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

        $allUpdates = @()

        # Call 1: no --scope flag (returns machine-scoped apps by default)
        try {
            $upgradeResult = & $wingetPath upgrade --source winget |
                Where-Object { $_ -notlike " *" } |
                Out-String
            $parsed = Parse-WingetUpgradeOutput -RawOutput $upgradeResult -Scope $null
            foreach ($u in $parsed) { $allUpdates += $u }
            Write-Log "Default scope: found $($parsed.Count) app(s) with updates" -Tag "Debug"
        }
        catch {
            Write-Log "Error getting updates (default scope): $_" -Tag "Debug"
        }

        # Call 2: --scope user (returns additional user-scoped apps)
        try {
            $upgradeResult = & $wingetPath upgrade --source winget --scope user |
                Where-Object { $_ -notlike " *" } |
                Out-String
            $parsed = Parse-WingetUpgradeOutput -RawOutput $upgradeResult -Scope 'user'
            foreach ($u in $parsed) { $allUpdates += $u }
            Write-Log "User scope: found $($parsed.Count) app(s) with updates" -Tag "Debug"
        }
        catch {
            Write-Log "Error getting updates (user scope): $_" -Tag "Debug"
        }

        # Deduplicate by AppId (prefer non-user scope if same app appears in both)
        $seen = @{}
        $updates = @()
        foreach ($u in $allUpdates) {
            if (-not $seen.ContainsKey($u.AppId)) {
                $seen[$u.AppId] = $true
                $updates += $u
            }
        }

        Write-Log "Found $($updates.Count) unique apps with available updates" -Tag "Get"
        return $updates
    }
    catch {
        Write-Log "Error getting available updates: $_" -Tag "Error"
        Write-Log $_.ScriptStackTrace -Tag "Debug"
        return @()
    }
    finally {
        [Console]::OutputEncoding = $previousOutputEncoding
    }
}

# ---------------------------[ Filter Updates by Blacklist or Whitelist ]---------------------------
function Select-FilteredUpdates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Updates,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Blacklist', 'Whitelist')]
        [string]$ListMode,

        [Parameter(Mandatory = $false)]
        [string[]]$Blacklist = @(),

        [Parameter(Mandatory = $false)]
        [string[]]$Whitelist = @()
    )

    if ($null -eq $Updates -or $Updates.Count -eq 0) {
        return @()
    }

    $filteredUpdates = @()

    foreach ($update in $Updates) {
        if (-not $update -or -not $update.AppId) {
            Write-Log "Invalid update object encountered, skipping" -Tag "Debug"
            continue
        }

        $appId = $update.AppId

        if ($ListMode -eq 'Blacklist') {
            if ($null -ne $Blacklist -and $Blacklist.Count -gt 0) {
                if (Test-AppMatch -AppId $appId -PatternList $Blacklist) {
                    Write-Log "App excluded (blacklist): $appId" -Tag "Debug"
                    continue
                }
            }
        }
        elseif ($ListMode -eq 'Whitelist') {
            if ($null -eq $Whitelist -or $Whitelist.Count -eq 0) {
                Write-Log "Whitelist mode with empty whitelist; no apps included" -Tag "Info"
                return @()
            }
            if (-not (Test-AppMatch -AppId $appId -PatternList $Whitelist)) {
                Write-Log "App excluded (not in whitelist): $appId" -Tag "Debug"
                continue
            }
        }

        $filteredUpdates += $update
    }

    Write-Log "Filtered to $($filteredUpdates.Count) apps requiring updates" -Tag "Get"
    return $filteredUpdates
}

# ---------------------------[ Update Application ]---------------------------
# Upgrade flow:
#   1. Try winget upgrade without --scope (let winget auto-detect)
#   2. In-progress wait loop: if another install is running, wait 30s and retry (up to 5x)
#   3. RetryScope: if "no applicable installer" and app was detected as user-scoped,
#      retry once with --scope user
#   4. RetryLater: transient errors (disk full, no network, app running) return $null
#      so Intune retries next cycle instead of flagging a failure
#   5. Fail: unrecoverable errors return $false
function Update-Application {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [string]$WingetPath,

        [Parameter(Mandatory = $false)]
        [string]$Scope
    )

    $scopeLabel = if ($Scope -eq 'user') { 'user' } else { 'default' }
    Write-Log "Updating: $AppId (detected scope: $scopeLabel)" -Tag "Run"

    $maxInProgressRetries   = 5
    $inProgressDelaySeconds = 30

    function Invoke-Upgrade {
        param([bool]$WithScopeUser)
        $wingetArgs = @('upgrade', '--id', $AppId, '-e', '--accept-package-agreements', '--accept-source-agreements', '-h', '--source', 'winget')
        if ($WithScopeUser) { $wingetArgs += '--scope', 'user' }
        Write-Log "Invoking: winget $($wingetArgs -join ' ')" -Tag "Debug"
        & $WingetPath @wingetArgs 2>&1 | Where-Object { $_ -notlike " *" }
    }

    try {
        # ── Attempt 1: upgrade without --scope ──
        $inProgressCount = 0
        do {
            if ($inProgressCount -gt 0) {
                Write-Log "Another installation in progress. Waiting ${inProgressDelaySeconds}s ($inProgressCount/$maxInProgressRetries)..." -Tag "Info"
                Start-Sleep -Seconds $inProgressDelaySeconds
            }

            $upgradeOutput = Invoke-Upgrade -WithScopeUser $false
            $exitCode = $LASTEXITCODE
            if ($null -eq $upgradeOutput) { $upgradeOutput = @() }

            if ($exitCode -ne -1978334974) { break }
            $inProgressCount++
        } while ($inProgressCount -le $maxInProgressRetries)

        if ($exitCode -eq -1978334974) {
            Write-Log "Still blocked after $maxInProgressRetries in-progress retries for $AppId" -Tag "Info"
            return $null
        }

        $exitInfo = Get-WingetExitCodeInfo -ExitCode $exitCode

        # ── Success ──
        if ($exitInfo.Category -eq 'Success') {
            Write-Log "Successfully updated: $AppId" -Tag "Success"
            return $true
        }

        # ── RetryScope: if "no applicable installer" and the app was detected as user-scoped ──
        if ($exitInfo.Category -eq 'RetryScope' -and $Scope -eq 'user') {
            Write-Log "No applicable installer without scope; retrying $AppId with --scope user." -Tag "Info"

            $upgradeOutput = Invoke-Upgrade -WithScopeUser $true
            $exitCode = $LASTEXITCODE
            $exitInfo = Get-WingetExitCodeInfo -ExitCode $exitCode

            if ($exitInfo.Category -eq 'Success') {
                Write-Log "Successfully updated: $AppId (with --scope user)" -Tag "Success"
                return $true
            }

            Write-Log "Retry with --scope user failed for $AppId - $exitCode ($($exitInfo.Description))" -Tag "Debug"
        }

        # ── RetryLater: transient ──
        if ($exitInfo.Category -eq 'RetryLater') {
            Write-Log "Transient error for $AppId ($($exitInfo.Description)); will retry next Intune cycle." -Tag "Info"
            return $null
        }

        # ── Fail / Unknown ──
        $errorMessages = $upgradeOutput | Where-Object {
            $_ -match 'error|failed|exception|unable|cannot|could not' -or
            ($_ -match '^[A-Z]' -and $_ -notmatch '^Loading|^Found|^Verifying|^Successfully')
        }
        if ($errorMessages) {
            Write-Log "Winget output for $AppId : $($errorMessages -join '; ')" -Tag "Debug"
        }
        Write-Log "Failed to update $AppId - $($exitInfo.Description) ($($exitInfo.Category))" -Tag "Error"
        return $false
    }
    catch {
        Write-Log "Error updating $AppId : $_" -Tag "Error"
        Write-Log $_.ScriptStackTrace -Tag "Debug"
        return $false
    }
}

# ---------------------------[ Main Remediation Logic ]---------------------------
try {
    if (-not (Test-Winget)) {
        Write-Log "Winget is not available or not working properly." -Tag "Error"
        Complete-Script -ExitCode 1
    }

    # One-time source refresh before we start
    $wingetPath = Get-WingetPath
    Write-Log "Refreshing winget sources..." -Tag "Run"
    & $wingetPath source update 2>&1 | Out-Null
    Write-Log "Source refresh complete." -Tag "Debug"

    if (Test-PendingReboot) {
        Write-Log "Pending reboot detected on system." -Tag "Info"
    }

    # Get available updates
    $availableUpdates = Get-AvailableUpdates

    if ($availableUpdates.Count -eq 0) {
        Write-Log "No updates available - all apps are up to date" -Tag "Success"
        Complete-Script -ExitCode 0
    }

    if ($listMode -eq 'Blacklist') {
        $listCount = if ($null -ne $blacklistApps) { $blacklistApps.Count } else { 0 }
        Write-Log "Using blacklist mode with $listCount entries" -Tag "Info"
    }
    else {
        $listCount = if ($null -ne $whitelistApps) { $whitelistApps.Count } else { 0 }
        Write-Log "Using whitelist mode with $listCount entries" -Tag "Info"
    }

    $filteredUpdates = Select-FilteredUpdates -Updates $availableUpdates -ListMode $listMode -Blacklist $blacklistApps -Whitelist $whitelistApps

    if ($filteredUpdates.Count -eq 0) {
        Write-Log "No updates needed after filtering - all managed apps are up to date" -Tag "Success"
        Complete-Script -ExitCode 0
    }

    # Perform updates
    Write-Log "Starting update process for $($filteredUpdates.Count) application(s)" -Tag "Run"

    $successCount   = 0
    $failureCount   = 0
    $deferredCount  = 0
    $failedApps     = @()
    $deferredApps   = @()

    $updateIndex = 0
    foreach ($update in $filteredUpdates) {
        $updateIndex++

        if (-not $update -or -not $update.AppId) {
            Write-Log "[$updateIndex/$($filteredUpdates.Count)] Skipping invalid update object" -Tag "Error"
            $failureCount++
            continue
        }

        Write-Log "[$updateIndex/$($filteredUpdates.Count)] Processing: $($update.AppId) ($($update.CurrentVersion) -> $($update.AvailableVersion))" -Tag "Info"

        $result = Update-Application -AppId $update.AppId -WingetPath $wingetPath -Scope $update.Scope

        if ($result -eq $true) {
            $successCount++
        }
        elseif ($null -eq $result) {
            $deferredCount++
            $deferredApps += $update.AppId
        }
        else {
            $failureCount++
            $failedApps += $update.AppId
        }

        Start-Sleep -Seconds 3
    }

    # Summary
    Write-Log "Successfully updated: $successCount application(s)" -Tag "Success"

    if ($deferredCount -gt 0) {
        Write-Log "Deferred (transient): $deferredCount application(s) - $($deferredApps -join ', ')" -Tag "Info"
    }

    if ($failureCount -gt 0) {
        Write-Log "Failed to update: $failureCount application(s) - $($failedApps -join ', ')" -Tag "Error"
        Complete-Script -ExitCode 1
    }
    else {
        Write-Log "All updates completed or deferred successfully" -Tag "Success"
        Complete-Script -ExitCode 0
    }
}
catch {
    Write-Log "Unexpected error in remediation script: $_" -Tag "Error"
    Write-Log $_.ScriptStackTrace -Tag "Debug"
    Complete-Script -ExitCode 1
}
