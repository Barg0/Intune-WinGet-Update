# ---------------------------[ Config ]---------------------------
# ListMode: 'Blacklist' = update all except listed apps; 'Whitelist' = update only listed apps
$listMode = 'Blacklist'

# Blacklist: apps EXCLUDED from updates (used when listMode = 'Blacklist')
$blacklistApps = @(
    '3CX.PhoneSystem',
    '3CX.Softphone',
    'Adobe.Acrobat.Pro',
    'Adobe.Acrobat.Reader.32-bit',
    'Adobe.Acrobat.Reader.64-bit',
    'Adobe.CreativeCloud',
    'Brave.Brave*',
    'dotPDN.PaintDotNet',
    'Fortinet.FortiClientVPN',
    'LANCOM.TrustedAccessClient',
    'Lenovo.SUHelper',
    'Microsoft.AdministrativeTemplates',
    'Microsoft.AppInstaller',
    'Microsoft.Edge*',
    'Microsoft.GlobalSecureAccessClient',
    'Microsoft.Office',
    'Microsoft.OneDrive',
    'Microsoft.PowerShell',
    'Microsoft.RemoteDesktopClient',
    'Microsoft.SurfaceApp',
    'Microsoft.Teams*',
    'Microsoft.VCLibs.*',
    'Microsoft.WindowsTerminal',
    'Mozilla.Firefox*',
    'Opera.Opera*',
    'TeamViewer.TeamViewer*',
    'TrackerSoftware.PDF-XChange*'
)

# Whitelist: apps INCLUDED in updates (used when listMode = 'Whitelist'). Wildcards supported.
$whitelistApps = @(
    '7zip.7zip',
    'Google.Chrome',
    'Microsoft.DotNet*',
    'Microsoft.VCRedist*',
    'Notepad++.Notepad++'
)

# ---------------------------[ Locale Workaround ]---------------------------
# If the upgrade ladder ends with 0x8A150014 (-1978335212, "No packages found") — e.g. after `--scope user` on a machine-scoped package —
# run the same ladder once more with `winget upgrade ... --locale <this>` (manifest/installer locale). Empty string disables.
$wingetLocaleWorkaround = 'en-US'

# ---------------------------[ WinGet installer busy ]---------------------------
# When WinGet returns INSTALL_IN_PROGRESS (0x8A150102 / -1978334974), wait and re-run the same upgrade attempt.
# Defaults match https://github.com/Barg0/Intune-WinGet/blob/main/templates/install.ps1 ($maxInProgressRetries / $inProgressDelaySeconds).
$wingetInProgressMaxRetries      = 15
$wingetInProgressWaitSeconds     = 120

# ---------------------------[ WinGet download retry ]---------------------------
# When WinGet returns a transient download error (0x8A150008, 0x8A15002E, 0x8A150086), wait this long then retry once.
$wingetDownloadRetryWaitSeconds  = 30

# ---------------------------[ Install fallback allowlist ]---------------------------
# When all upgrade attempts fail with 0x8A150014 ("No packages found"), fall back to `winget install --version --force`.
# The install command resolves against source manifests instead of ARP entries, bypassing the matching bug.
# Only AppIds matching a pattern here may use this workaround. Wildcards supported (same as blacklist). Empty @() = never.
# @('*') = allow for every package.
$wingetInstallFallbackAllowlist = @( 'Nextcloud.Talk' )

# ---------------------------[ Uninstall-previous allowlist ]---------------------------
# LAST RESORT: `winget upgrade --uninstall-previous`. Uninstalls the existing version before installing the new one.
# Destructive if uninstall succeeds but install fails. Only AppIds matching a pattern here may use this workaround.
# Empty @() = never. @('*') = allow for every package.
$wingetUninstallPreviousAllowlist = @(
    # 'Contoso.ExampleApp'
)

# ---------------------------[ Scope ladder ]---------------------------
# Ordered WinGet `--scope` attempts: upgrades, install fallback, and uninstall-previous.
# Use 'Machine' (--scope machine), 'Default' (omit --scope), 'User' (--scope user).
# Aliases (case-insensitive): 'system' -> Machine, 'none' -> Default. No duplicate scopes.
$wingetScopeLadderOrder = @('machine', 'none', 'user')

# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Script Name ]---------------------------
$scriptName  = 'WinGet-Update'
$logFileName = "remediation.log"

# ---------------------------[ Logging Setup ]---------------------------
$log           = $true
$logDebug      = $false
$logGet        = $true
$logRun        = $true
$enableLogFile = $true

$logFileDirectory = "$env:ProgramData\IntuneLogs\Scripts\$scriptName"
$logFile          = "$logFileDirectory\$logFileName"

# ---------------------------[ Winget Exit Code Map ]---------------------------
# Only codes that can realistically fire during winget upgrade / install.
# Sources:
#   https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md
#   https://kb.filewave.com/books/microsoft-windows-package-manager-winget/page/troubleshooting-errors-with-winget
#
# Categories drive the retry engine:
#   Success            – Desired state reached. No action needed.
#   RetryScope         – No applicable installer / no packages for scope; advance upgrade ladder (machine → default → user).
#   RetrySourceRepair  – Source index is corrupt/missing. Run `winget source reset --force` + `source update`, then retry ladder.
#   RetryHashRefresh   – Installer hash mismatch. Run `winget source update` (pull fresh manifests), then retry same scope.
#   RetryDownload      – Transient CDN / download failure. Wait, then retry same scope once.
#   RetryLater         – Transient (app in use, disk full, reboot, network). Defer to next Intune cycle.
#   Fail               – Unrecoverable in automation (policy block, unsupported, structural).
#
# WinGet often returns signed 32-bit HRESULTs (negative). Do not cast those directly to
# [uint32] for hex display — it throws. Reinterpret bits via BitConverter instead.
function Format-WingetExitCodeHex {
    param([int]$Code)
    $u = [System.BitConverter]::ToUInt32([System.BitConverter]::GetBytes([int32]$Code), 0)
    return ('0x{0:X8}' -f $u)
}

function Get-WingetExitCodeInfo {
    [CmdletBinding()]
    param([int]$ExitCode)
    $codeMap = @{
        # ── Success: desired state reached ──
        0              = @{ Category = 'Success';            Description = 'Success' }
        3010           = @{ Category = 'Success';            Description = 'Success (reboot required to complete)' }               # MSI ERROR_SUCCESS_REBOOT_REQUIRED
        -1978335135    = @{ Category = 'Success';            Description = 'Package already installed' }                            # 0x8A150061
        -1978334965    = @{ Category = 'Success';            Description = 'Reboot initiated to finish installation' }              # 0x8A15010B
        -1978334963    = @{ Category = 'Success';            Description = 'Another version already installed' }                    # 0x8A15010D
        -1978334962    = @{ Category = 'Success';            Description = 'Higher version already installed (downgrade)' }         # 0x8A15010E

        # ── RetryScope: advance upgrade ladder (machine → default → user) ──
        # Aligned with Barg0/Intune-WinGet install.ps1.
        -1978335216    = @{ Category = 'RetryScope';         Description = 'No applicable installer for current scope' }            # 0x8A150010
        -1978335212    = @{ Category = 'RetryScope';         Description = 'No packages found' }                                    # 0x8A150014
        # ShellExecute failed to *launch* the installer — often a scope/elevation mismatch in SYSTEM context.
        -1978335226    = @{ Category = 'RetryScope';         Description = 'ShellExecute install failed (try other scopes)' }       # 0x8A150006

        # ── RetrySourceRepair: source corrupt/missing → winget source reset --force + source update, then retry ladder ──
        -1978335222    = @{ Category = 'RetrySourceRepair';  Description = 'Index is corrupt' }                                     # 0x8A15000A
        -1978335221    = @{ Category = 'RetrySourceRepair';  Description = 'Configured source information is corrupt' }             # 0x8A15000B
        -1978335217    = @{ Category = 'RetrySourceRepair';  Description = 'Source data missing' }                                  # 0x8A15000F
        -1978335169    = @{ Category = 'RetrySourceRepair';  Description = 'Source data corrupted or tampered' }                    # 0x8A15003F
        -1978335163    = @{ Category = 'RetrySourceRepair';  Description = 'Failed to open source' }                                # 0x8A150045
        -1978335157    = @{ Category = 'RetrySourceRepair';  Description = 'Failed to open one or more sources' }                   # 0x8A15004B

        # ── RetryHashRefresh: installer hash mismatch → winget source update (pull fresh manifests), then retry ──
        # --force does NOT override hash checks when running as admin/SYSTEM (winget-cli #1812).
        # A source update often fixes it because the publisher updated the binary before the manifest was refreshed.
        -1978335215    = @{ Category = 'RetryHashRefresh';   Description = 'Installer hash does not match manifest' }               # 0x8A150011

        # ── RetryDownload: transient CDN / download error → wait, then retry once ──
        -1978335224    = @{ Category = 'RetryDownload';      Description = 'Download failed' }                                      # 0x8A150008
        -1978335186    = @{ Category = 'RetryDownload';      Description = 'Download size mismatch' }                               # 0x8A15002E
        -1978335098    = @{ Category = 'RetryDownload';      Description = 'Downloaded zero-byte installer' }                       # 0x8A150086

        # ── RetryLater: transient – defer to next Intune cycle ──
        -1978335227    = @{ Category = 'RetryLater';         Description = 'Cancellation signal received' }                         # 0x8A150005
        -1978335126    = @{ Category = 'RetryLater';         Description = 'Application shutdown signal received' }                 # 0x8A15006A
        -1978335125    = @{ Category = 'RetryLater';         Description = 'Failed to download dependencies' }                      # 0x8A15006B
        -1978335123    = @{ Category = 'RetryLater';         Description = 'Service busy or unavailable' }                          # 0x8A15006D
        -1978334975    = @{ Category = 'RetryLater';         Description = 'Application is currently running' }                     # 0x8A150101
        -1978334974    = @{ Category = 'RetryLater';         Description = 'Another installation in progress' }                     # 0x8A150102
        -1978334973    = @{ Category = 'RetryLater';         Description = 'One or more file is in use' }                           # 0x8A150103
        -1978334971    = @{ Category = 'RetryLater';         Description = 'Disk full' }                                            # 0x8A150105
        -1978334970    = @{ Category = 'RetryLater';         Description = 'Insufficient memory' }                                  # 0x8A150106
        -1978334969    = @{ Category = 'RetryLater';         Description = 'No network connectivity' }                              # 0x8A150107
        -1978334967    = @{ Category = 'RetryLater';         Description = 'Reboot required to finish installation' }               # 0x8A150109
        -1978334966    = @{ Category = 'RetryLater';         Description = 'Reboot required then try again' }                       # 0x8A15010A
        -1978334959    = @{ Category = 'RetryLater';         Description = 'Application in use by another application' }            # 0x8A150111

        # ── Fail: unrecoverable without human intervention ──
        -1978335231    = @{ Category = 'Fail';               Description = 'Internal error' }                                       # 0x8A150001
        -1978335230    = @{ Category = 'Fail';               Description = 'Invalid command line arguments' }                       # 0x8A150002
        -1978335229    = @{ Category = 'Fail';               Description = 'Command failed' }                                       # 0x8A150003
        -1978335228    = @{ Category = 'Fail';               Description = 'Opening manifest failed' }                              # 0x8A150004
        -1978335225    = @{ Category = 'Fail';               Description = 'Manifest version higher than supported; update winget' } # 0x8A150007
        -1978335210    = @{ Category = 'Fail';               Description = 'Multiple packages found matching criteria' }            # 0x8A150016
        -1978335209    = @{ Category = 'Fail';               Description = 'No manifest found matching criteria' }                  # 0x8A150017
        -1978335207    = @{ Category = 'Fail';               Description = 'Command requires administrator privileges' }            # 0x8A150019
        -1978335205    = @{ Category = 'Fail';               Description = 'Microsoft Store client blocked by policy' }             # 0x8A15001B
        -1978335204    = @{ Category = 'Fail';               Description = 'Microsoft Store app blocked by policy' }                # 0x8A15001C
        -1978335189    = @{ Category = 'RetryScope';          Description = 'No applicable upgrade (does not apply to system or scope)' } # 0x8A15002B
        -1978335188    = @{ Category = 'Fail';               Description = 'upgrade --all completed with failures' }                # 0x8A15002C
        -1978335187    = @{ Category = 'Fail';               Description = 'Installer failed security check' }                      # 0x8A15002D
        -1978335174    = @{ Category = 'Fail';               Description = 'Blocked by Group Policy' }                              # 0x8A15003A
        -1978335159    = @{ Category = 'Fail';               Description = 'MSI install failed' }                                   # 0x8A150049
        -1978335153    = @{ Category = 'Fail';               Description = 'Upgrade version not newer than installed' }              # 0x8A15004F
        -1978335152    = @{ Category = 'Fail';               Description = 'Upgrade version unknown; override not specified' }       # 0x8A150050
        -1978335146    = @{ Category = 'Fail';               Description = 'Installer prohibits elevation' }                         # 0x8A150056
        -1978335138    = @{ Category = 'Fail';               Description = 'Pinned certificate mismatch' }                           # 0x8A15005E
        -1978335128    = @{ Category = 'Fail';               Description = 'Package has a pin that prevents upgrade' }               # 0x8A150068
        -1978335127    = @{ Category = 'Fail';               Description = 'Package is a stub; full package needed' }                # 0x8A150069
        -1978335090    = @{ Category = 'Fail';               Description = 'Install technology mismatch (different installer type)' } # 0x8A15008E
        -1978334972    = @{ Category = 'Fail';               Description = 'Missing dependency on system' }                          # 0x8A150104
        -1978334968    = @{ Category = 'Fail';               Description = 'Installation error; contact support' }                   # 0x8A150108
        -1978334964    = @{ Category = 'Fail';               Description = 'Installation cancelled by user' }                        # 0x8A15010C
        -1978334961    = @{ Category = 'Fail';               Description = 'Blocked by organization policy' }                        # 0x8A15010F
        -1978334960    = @{ Category = 'Fail';               Description = 'Failed to install package dependencies' }                # 0x8A150110
        -1978334958    = @{ Category = 'Fail';               Description = 'Invalid parameter' }                                     # 0x8A150112
        -1978334957    = @{ Category = 'Fail';               Description = 'Package not supported by system' }                       # 0x8A150113
        -1978334956    = @{ Category = 'Fail';               Description = 'Installer does not support upgrading existing package' }  # 0x8A150114
        -1978334955    = @{ Category = 'Fail';               Description = 'Installer custom error' }                                # 0x8A150115

        # ── Win32 / WinINet codes – frequently seen as "Unmapped exit code" in Intune runs ──
        # These originate from the installer process (EXE/MSI) or from WinGet's own HTTP stack, not from
        # the AppInstaller HRESULT space (0x8A15xxxx), so WinGet logs them verbatim as "Unmapped".
        # Reference: winerror.h / wininet.h; confirmed against winget-cli issues #1546, #5232, #4283.
        -2147023673    = @{ Category = 'RetryLater';         Description = 'Operation cancelled - ERROR_CANCELLED (0x800704C7)' }
        -2147012894    = @{ Category = 'RetryLater';         Description = 'Connection timed out - ERROR_INTERNET_TIMEOUT (0x80072EE2)' }
        -2147012889    = @{ Category = 'RetryLater';         Description = 'DNS name not resolved - ERROR_INTERNET_NAME_NOT_RESOLVED (0x80072EE7)' }
        -2147012867    = @{ Category = 'RetryLater';         Description = 'Cannot connect to server - ERROR_INTERNET_CANNOT_CONNECT (0x80072EFD)' }
        -2147012866    = @{ Category = 'RetryLater';         Description = 'Connection aborted - ERROR_INTERNET_CONNECTION_ABORTED (0x80072EFE)' }
        -2147012465    = @{ Category = 'RetryLater';         Description = 'TLS/SSL error - ERROR_INTERNET_DECRYPTION_FAILED (0x80072F8F)' }
        -2147221003    = @{ Category = 'Fail';               Description = 'Application/uninstaller not found - orphaned ARP entry (0x800401F5)' }
        -2147024891    = @{ Category = 'Fail';               Description = 'Access denied - ERROR_ACCESS_DENIED (0x80070005)' }
        -2147023293    = @{ Category = 'Fail';               Description = 'MSI fatal error - ERROR_INSTALL_FAILURE (0x80070643 / 1603)' }
        -2147023286    = @{ Category = 'RetryLater';         Description = 'Windows Installer busy - ERROR_INSTALL_ALREADY_RUNNING (0x8007064A / 1610)' }
    }
    if ($codeMap.ContainsKey($ExitCode)) { return $codeMap[$ExitCode] }
    $hex = Format-WingetExitCodeHex -Code $ExitCode
    return @{ Category = 'Unknown'; Description = "Unmapped exit code $ExitCode ($hex)" }
}

if ($enableLogFile -and -not (Test-Path -Path $logFileDirectory)) {
    try {
        $null = New-Item -ItemType Directory -Path $logFileDirectory -Force -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to create log directory '$logFileDirectory': $($_.Exception.Message)"
    }
}

# Logging aligned with https://github.com/Barg0/Intune-Win32-Scripts (compact line, I/O warnings).
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$message,
        [string]$tag = "Info"
    )

    if (-not $log) { return }

    if (($tag -eq "Debug") -and (-not $logDebug)) { return }
    if (($tag -eq "Get") -and (-not $logGet)) { return }
    if (($tag -eq "Run") -and (-not $logRun)) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList   = @("Start", "Get", "Run", "Info", "Success", "Error", "Debug", "End")
    $rawTag    = $tag.Trim()
    if ($tagList -contains $rawTag) { $rawTag = $rawTag.PadRight(7) }
    else { $rawTag = "Error " }

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

    $logMessage = "$timestamp [ $rawTag ] $message"
    if ($enableLogFile) {
        try {
            Add-Content -Path $logFile -Value $logMessage -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }

    Write-Host "$timestamp " -NoNewline
    Write-Host "[ " -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host " ] " -NoNewline -ForegroundColor White
    Write-Host "$message"
}

# One-line summary for Intune portal (last console line). Not written to the log file.
function Build-RemediationPortalSummaryLine {
    param(
        [string[]]$Succeeded = @(),
        [string[]]$Failed = @(),
        [string[]]$Deferred = @()
    )
    $okList  = if (@($Succeeded).Count -gt 0) { @($Succeeded) -join ', ' } else { '(none)' }
    $badList = if (@($Failed).Count -gt 0) { @($Failed) -join ', ' } else { '(none)' }
    $line    = "Updated: $okList | Failed: $badList"
    if (@($Deferred).Count -gt 0) {
        $line += " | Deferred: $(@($Deferred) -join ', ')"
    }
    return $line
}

# ---------------------------[ Exit Function ]---------------------------
function Complete-Script {
    param(
        [int]$ExitCode,
        [string]$PortalSummaryLine = $null
    )

    $scriptEndTime = Get-Date
    $duration      = $scriptEndTime - $scriptStartTime

    Write-Log "Runtime $($duration.ToString('hh\:mm\:ss\.ff'))" -Tag "Info"
    Write-Log "Exit $ExitCode" -Tag "Info"
    Write-Log "==================== End ====================" -Tag "End"

    if (-not [string]::IsNullOrEmpty($PortalSummaryLine)) {
        Write-Host $PortalSummaryLine
    }

    exit $ExitCode
}

# ---------------------------[ Script Start ]---------------------------
Write-Log "==================== Start ====================" -Tag "Start"
Write-Log "Host $env:COMPUTERNAME | $env:USERNAME | $scriptName" -Tag "Info"

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

        Write-Log "Winget: no DesktopAppInstaller / winget.exe" -Tag "Error"
        throw "Winget not found in system or user context"
    }
    catch {
        if ($_.Exception.Message -notlike 'Winget not found*') {
            Write-Log "Winget resolve: $_" -Tag "Error"
        }
        throw "Winget not found in system or user context"
    }
}

# ---------------------------[ Test Pending Reboot ]---------------------------
# Same checks as https://github.com/Barg0/Intune-Win32-Scripts/blob/main/installExe.ps1 (Test-PendingReboot).
# Log-only here; script does not exit on pending reboot.
function Test-PendingReboot {
    [CmdletBinding()]
    param()

    $rebootKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts'
    )
    foreach ($key in $rebootKeys) {
        if (Test-Path -Path $key) { return $true }
    }

    $sessionManagerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    if (Test-Path -Path $sessionManagerPath) {
        $pfro = Get-ItemProperty -Path $sessionManagerPath -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
        if ($pfro -and $pfro.PendingFileRenameOperations) { return $true }
        $pfro2 = Get-ItemProperty -Path $sessionManagerPath -Name 'PendingFileRenameOperations2' -ErrorAction SilentlyContinue
        if ($pfro2 -and $pfro2.PendingFileRenameOperations2) { return $true }
    }

    $updatePaths = @(
        'HKLM:\SOFTWARE\Microsoft\Updates',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Updates'
    )
    foreach ($updatePath in $updatePaths) {
        if (Test-Path -Path $updatePath) {
            $volatile = Get-ItemProperty -Path $updatePath -Name 'UpdateExeVolatile' -ErrorAction SilentlyContinue
            if ($volatile -and $volatile.UpdateExeVolatile -ne 0) { return $true }
        }
    }

    $computerNamePath = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName'
    $activeNamePath = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName'
    if ((Test-Path -Path $computerNamePath) -and (Test-Path -Path $activeNamePath)) {
        $computerName = (Get-ItemProperty -Path $computerNamePath -Name 'ComputerName' -ErrorAction SilentlyContinue).ComputerName
        $activeName = (Get-ItemProperty -Path $activeNamePath -Name 'ComputerName' -ErrorAction SilentlyContinue).ComputerName
        if ($computerName -and $activeName -and $computerName -ne $activeName) { return $true }
    }

    return $false
}

# ---------------------------[ Test Winget Function ]---------------------------
function Test-Winget {
    [CmdletBinding()]
    param()

    Write-Log "WinGet check" -Tag "Debug"

    try {
        $wingetPath = Get-WingetPath
        $rawOutput = & $wingetPath -v 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            $versionLine = $rawOutput | Where-Object { $_ -and ($_ -match '\d+\.\d+') } | Select-Object -First 1
            if ($versionLine -and $versionLine -match '(\d+\.\d+(?:\.\d+)?(?:\.\d+)?)') {
                Write-Log "WinGet: v$($matches[1])" -Tag "Success"
            }
            else {
                Write-Log "WinGet" -Tag "Success"
            }
            return $true
        }
        else {
            Write-Log "Winget failed: exit $exitCode" -Tag "Error"
            $errorOutput = $rawOutput | Where-Object { $_ -and $_ -notmatch '^\s*$' } | Select-Object -First 3
            if ($errorOutput) {
                Write-Log "Details: $($errorOutput -join '; ')" -Tag "Debug"
            }
            return $false
        }
    }
    catch {
        Write-Log "Winget test: $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Test-AppMatch ]---------------------------
function Test-AppMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,

        # Empty @() is valid: no patterns → no match. Required for empty blacklist/whitelist/allowlist configs.
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$PatternList
    )

    foreach ($pattern in $PatternList) {
        if ($AppId -like $pattern) {
            return $true
        }
    }

    return $false
}

# ---------------------------[ Convert Winget Upgrade Output ]---------------------------
# Returns use `return , $updates` so a single row stays Object[]; otherwise PowerShell unwraps
# one hashtable and .Count becomes the number of keys (5), not the number of apps.
function ConvertFrom-WingetUpgradeOutput {
    [CmdletBinding()]
    param(
        [string]$RawOutput,
        [string]$Scope
    )

    $updates = @()
    $unknownCount = 0

    if (-not ($RawOutput -match "-----")) {
        return , $updates
    }

    $lines = $RawOutput.Split([Environment]::NewLine) | Where-Object { $_ }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lines[$i] = $lines[$i] -replace "[\u2026]", " "
    }

    $fl = 0
    while ($fl -lt $lines.Count -and -not $lines[$fl].StartsWith("-----")) { $fl++ }
    if ($fl -ge $lines.Count) { return , $updates }
    $fl = $fl - 1
    if ($fl -lt 0) { return , $updates }

    $index = $lines[$fl] -split '(?<=\s)(?!\s)'
    if ($index.Count -lt 3) { return , $updates }

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
    return , $updates
}

# ---------------------------[ Get Available Updates ]---------------------------
# Three calls: no --scope, then --scope user, then --scope machine. Union by AppId (first-seen wins
# for metadata: default list, then user-only additions, then machine-only additions).
function Get-AvailableUpdates {
    [CmdletBinding()]
    param()

    Write-Log "Upgrades: list (default, user, machine)" -Tag "Debug"
    $wingetPath = Get-WingetPath

    try {
        $previousOutputEncoding = [Console]::OutputEncoding
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

        $allUpdates = @()

        # Call 1: no --scope (winget default)
        try {
            $upgradeResult = & $wingetPath upgrade --source winget |
                Where-Object { $_ -notlike " *" } |
                Out-String
            $parsed = ConvertFrom-WingetUpgradeOutput -RawOutput $upgradeResult -Scope $null
            foreach ($u in $parsed) { $allUpdates += $u }
            Write-Log "Upgrades default: $($parsed.Count)" -Tag "Debug"
        }
        catch {
            Write-Log "List error (default): $_" -Tag "Debug"
        }

        # Call 2: --scope user
        try {
            $upgradeResult = & $wingetPath upgrade --source winget --scope user |
                Where-Object { $_ -notlike " *" } |
                Out-String
            $parsed = ConvertFrom-WingetUpgradeOutput -RawOutput $upgradeResult -Scope 'user'
            foreach ($u in $parsed) { $allUpdates += $u }
            Write-Log "Upgrades user: $($parsed.Count)" -Tag "Debug"
        }
        catch {
            Write-Log "List error (user): $_" -Tag "Debug"
        }

        # Call 3: --scope machine
        try {
            $upgradeResult = & $wingetPath upgrade --source winget --scope machine |
                Where-Object { $_ -notlike " *" } |
                Out-String
            $parsed = ConvertFrom-WingetUpgradeOutput -RawOutput $upgradeResult -Scope 'machine'
            foreach ($u in $parsed) { $allUpdates += $u }
            Write-Log "Upgrades machine: $($parsed.Count)" -Tag "Debug"
        }
        catch {
            Write-Log "List error (machine): $_" -Tag "Debug"
        }

        $seen = @{}
        $updates = @()
        foreach ($u in $allUpdates) {
            if (-not $seen.ContainsKey($u.AppId)) {
                $seen[$u.AppId] = $true
                $updates += $u
            }
        }

        Write-Log "Upgrades: $($updates.Count)" -Tag "Get"
        return , $updates
    }
    catch {
        Write-Log "Get updates: $_" -Tag "Error"
        Write-Log "$($_.ScriptStackTrace)" -Tag "Debug"
        return , @()
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

    if ($null -eq $Updates) {
        return , @()
    }
    $Updates = @($Updates)
    if ($Updates.Count -eq 0) {
        return , @()
    }

    $filteredUpdates = @()

    foreach ($update in $Updates) {
        if (-not $update -or -not $update.AppId) {
            Write-Log "Invalid row; skip" -Tag "Debug"
            continue
        }

        $appId = $update.AppId

        if ($ListMode -eq 'Blacklist') {
            if ($null -ne $Blacklist -and $Blacklist.Count -gt 0) {
                if (Test-AppMatch -AppId $appId -PatternList $Blacklist) {
                    Write-Log "Blacklist: $appId" -Tag "Info"
                    continue
                }
            }
        }
        elseif ($ListMode -eq 'Whitelist') {
            if ($null -eq $Whitelist -or $Whitelist.Count -eq 0) {
                Write-Log "Whitelist empty; none" -Tag "Info"
                return , @()
            }
            if (-not (Test-AppMatch -AppId $appId -PatternList $Whitelist)) {
                Write-Log "Whitelist: skip $appId" -Tag "Debug"
                continue
            }
        }

        $filteredUpdates += $update
    }

    Write-Log "Filtered: $($filteredUpdates.Count)" -Tag "Get"
    return , $filteredUpdates
}

function Get-WingetScopeLadderOrderNormalized {
    [CmdletBinding()]
    param()
    $rawList = @($wingetScopeLadderOrder)
    if ($rawList.Count -eq 0) {
        throw 'wingetScopeLadderOrder must be a non-empty array.'
    }
    # Expand each array element: commas and/or spaces become separate tokens so
    # @('Default Machine User'), @('Machine,Default,User'), or a single mistaken string do not pass one invalid ScopeMode.
    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($raw in $rawList) {
        $t = [string]$raw
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        foreach ($commaPart in ($t -split ',')) {
            $cp = $commaPart.Trim()
            if ([string]::IsNullOrWhiteSpace($cp)) { continue }
            foreach ($word in ($cp -split '\s+')) {
                $w = $word.Trim()
                if (-not [string]::IsNullOrWhiteSpace($w)) {
                    [void]$tokens.Add($w)
                }
            }
        }
    }
    if ($tokens.Count -eq 0) {
        throw 'wingetScopeLadderOrder resolved to no tokens (only blank entries?).'
    }

    $normalized = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    foreach ($key in $tokens) {
        $mode = switch -Regex ($key.ToLowerInvariant()) {
            '^(machine|system)$' { 'Machine'; break }
            '^(default|none)$' { 'Default'; break }
            '^(user)$' { 'User'; break }
            default { throw "Invalid wingetScopeLadderOrder token: '$key'. Use Machine, Default, User (or system, none)." }
        }
        if ($seen.ContainsKey($mode)) {
            throw "Duplicate scope in wingetScopeLadderOrder: $mode"
        }
        $seen[$mode] = $true
        [void]$normalized.Add($mode)
    }
    # Do not use `return , $array` — the unary comma wraps the array so callers get one nested object;
    # coercing that to [string] for -ScopeMode becomes "Default Machine User" (joined with spaces).
    return $normalized.ToArray()
}

function Get-WingetScopeSuccessSuffix {
    param(
        [ValidateSet('Machine', 'Default', 'User')]
        [string]$ScopeMode
    )
    switch ($ScopeMode) {
        'Machine' { return 'machine' }
        'Default' { return 'default scope' }
        'User' { return 'user' }
    }
}

function Get-WingetScopeUpgradeRetryLog {
    param(
        [ValidateSet('Machine', 'Default', 'User')]
        [string]$ScopeMode
    )
    switch ($ScopeMode) {
        'Machine' { return 'Retry: --scope machine' }
        'Default' { return 'Retry: no --scope' }
        'User' { return 'Retry: --scope user' }
    }
}

function Get-WingetScopeShortName {
    param(
        [ValidateSet('Machine', 'Default', 'User')]
        [string]$ScopeMode
    )
    switch ($ScopeMode) {
        'Machine' { return 'machine' }
        'Default' { return 'default' }
        'User' { return 'user' }
    }
}

# WinGet sometimes prints "No applicable upgrade found" (exit 0 or 0x8A15002B) while `winget upgrade` list still shows a newer version.
function Test-WingetUpgradeOutputClaimsNoApplicable {
    param(
        [AllowNull()]
        $Lines
    )
    if ($null -eq $Lines) { return $false }
    $t = ($Lines | Out-String)
    return $t -match '(?i)No applicable upgrade|does not apply to your system or requirements'
}

# ---------------------------[ Update Application ]---------------------------
# Upgrade ladder: order from $wingetScopeLadderOrder (normalized in main to $script:WingetScopeLadderNormalized).
#   - First scope always runs; each further scope runs only if prior attempt did not defer and
#     (Test-ShouldAdvanceScopeLadder) or exit 0x8A150014 (-1978335212).
#   - Success only if exit category Success AND output does not claim no applicable upgrade (guards false exit 0)
#   - RetryLater / hash / download -> $null
#   - Locale pass + source repair unchanged; Get-AvailableUpdates listing order is separate from this ladder.
function Update-Application {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [string]$WingetPath,

        [string]$AvailableVersion = ''
    )

    function Invoke-Upgrade {
        param(
            [ValidateSet('Machine', 'Default', 'User')]
            [string]$ScopeMode,

            [string]$Locale
        )
        $wingetArgs = @('upgrade', '--id', $AppId, '-e', '--force', '--accept-package-agreements', '--accept-source-agreements', '--silent', '--disable-interactivity', '--skip-dependencies', '--source', 'winget')
        if (-not [string]::IsNullOrWhiteSpace($Locale)) { $wingetArgs += '--locale', $Locale.Trim() }
        if ($ScopeMode -eq 'Machine') { $wingetArgs += '--scope', 'machine' }
        elseif ($ScopeMode -eq 'User') { $wingetArgs += '--scope', 'user' }
        Write-Log "winget $($wingetArgs -join ' ')" -Tag "Debug"
        & $WingetPath @wingetArgs 2>&1 | Where-Object { $_ -notlike " *" }
    }

    function Invoke-UpgradeWithInProgressWait {
        param(
            [ValidateSet('Machine', 'Default', 'User')]
            [string]$ScopeMode,

            [string]$Locale
        )
        $inProgressCount = 0
        $upgradeOutput = @()
        $exitCode = 0
        do {
            if ($inProgressCount -gt 0) {
                Write-Log "Install busy; wait ${wingetInProgressWaitSeconds}s ($inProgressCount/$wingetInProgressMaxRetries)" -Tag "Info"
                Start-Sleep -Seconds $wingetInProgressWaitSeconds
            }
            $upgradeOutput = Invoke-Upgrade -ScopeMode $ScopeMode -Locale $Locale
            $exitCode = $LASTEXITCODE
            if ($null -eq $upgradeOutput) { $upgradeOutput = @() }
            if ($exitCode -ne -1978334974) { break }
            $inProgressCount++
        } while ($inProgressCount -le $wingetInProgressMaxRetries)
        return @{ Output = $upgradeOutput; ExitCode = $exitCode }
    }

    function Invoke-UpgradeAttempt {
        param(
            [ValidateSet('Machine', 'Default', 'User')]
            [string]$ScopeMode,

            [string]$Locale
        )
        $result = Invoke-UpgradeWithInProgressWait -ScopeMode $ScopeMode -Locale $Locale
        $exitInfo = Get-WingetExitCodeInfo -ExitCode $result.ExitCode

        if ($exitInfo.Category -eq 'RetryHashRefresh') {
            Write-Log "Hash mismatch ${AppId}: refreshing source index" -Tag "Info"
            & $WingetPath source update --name winget 2>&1 | Out-Null
            $result = Invoke-UpgradeWithInProgressWait -ScopeMode $ScopeMode -Locale $Locale
        }
        elseif ($exitInfo.Category -eq 'RetryDownload') {
            Write-Log "Download failed ${AppId}: retry in ${wingetDownloadRetryWaitSeconds}s" -Tag "Info"
            Start-Sleep -Seconds $wingetDownloadRetryWaitSeconds
            $result = Invoke-UpgradeWithInProgressWait -ScopeMode $ScopeMode -Locale $Locale
        }

        return $result
    }

    function Test-ShouldAdvanceScopeLadder {
        param(
            $ExitInfo,
            [int]$ExitCode,
            [bool]$OutputClaimsNoApplicable
        )
        return ($ExitInfo.Category -eq 'RetryScope') -or $OutputClaimsNoApplicable
    }

    try {
        $scopeOrder = $script:WingetScopeLadderNormalized
        if ($null -eq $scopeOrder -or @($scopeOrder).Count -eq 0) {
            throw 'WingetScopeLadderNormalized is not set. Remediation main must call Get-WingetScopeLadderOrderNormalized first.'
        }
        $scopeOrder = @($scopeOrder)

        $localePassMax = if ([string]::IsNullOrWhiteSpace($wingetLocaleWorkaround)) { 1 } else { 2 }
        $sourceRepairDone = $false

        for ($localePass = 0; $localePass -lt $localePassMax; $localePass++) {
            $localeArg = ''
            if ($localePass -eq 1) {
                $localeArg = $wingetLocaleWorkaround.Trim()
                Write-Log "Retry: --locale $localeArg" -Tag "Info"
            }

            $deferCategories = @('RetryLater', 'RetryHashRefresh', 'RetryDownload')

            # Scope ladder — runs twice at most: once normally, once after source repair
            for ($ladderRun = 0; $ladderRun -lt 2; $ladderRun++) {
                if ($ladderRun -eq 1) {
                    Write-Log "Source repair: winget source reset + update" -Tag "Info"
                    & $WingetPath source reset --force 2>&1 | Out-Null
                    & $WingetPath source update 2>&1 | Out-Null
                    Write-Log "Source repaired; retrying upgrade" -Tag "Info"
                }

                $runNotes = @()
                if ($localePass -eq 1) { $runNotes += 'locale' }
                if ($ladderRun -eq 1)  { $runNotes += 'source repair' }
                $successNote = if ($runNotes.Count -gt 0) { " ($($runNotes -join ', '))" } else { '' }

                for ($si = 0; $si -lt $scopeOrder.Count; $si++) {
                    $scopeMode = $scopeOrder[$si]

                    if ($si -gt 0) {
                        $tryNextScope = (Test-ShouldAdvanceScopeLadder -ExitInfo $exitInfo -ExitCode $exitCode -OutputClaimsNoApplicable $outputClaimsNoApplicable) -or ($exitCode -eq -1978335212)
                        if (-not $tryNextScope) { break }
                        Write-Log (Get-WingetScopeUpgradeRetryLog -ScopeMode $scopeMode) -Tag "Info"
                    }

                    $attempt = Invoke-UpgradeAttempt -ScopeMode $scopeMode -Locale $localeArg
                    $upgradeOutput = $attempt.Output
                    $exitCode = $attempt.ExitCode

                    if ($exitCode -eq -1978334974) {
                        Write-Log "Defer ${AppId}: install busy (max waits)" -Tag "Info"
                        return $null
                    }

                    $exitInfo = Get-WingetExitCodeInfo -ExitCode $exitCode
                    $outputClaimsNoApplicable = Test-WingetUpgradeOutputClaimsNoApplicable -Lines $upgradeOutput
                    $treatAsSuccess = ($exitInfo.Category -eq 'Success') -and -not $outputClaimsNoApplicable

                    if ($treatAsSuccess) {
                        $suffix = Get-WingetScopeSuccessSuffix -ScopeMode $scopeMode
                        Write-Log "$AppId ($suffix)$successNote" -Tag "Success"
                        return $true
                    }

                    if ($exitInfo.Category -in $deferCategories) {
                        Write-Log "Defer ${AppId}: $($exitInfo.Description)" -Tag "Info"
                        return $null
                    }

                    if ($si -eq ($scopeOrder.Count - 1)) {
                        $short = Get-WingetScopeShortName -ScopeMode $scopeMode
                        Write-Log "$short scope: exit $exitCode $($exitInfo.Description)" -Tag "Debug"
                    }
                }

                # Source repair: if source-related error and not yet repaired, re-run ladder after repair
                if ($ladderRun -eq 0 -and -not $sourceRepairDone -and $exitInfo.Category -eq 'RetrySourceRepair') {
                    $sourceRepairDone = $true
                    continue
                }
                break
            }

            if ($localePass -eq 0 -and $exitCode -eq -1978335212 -and -not [string]::IsNullOrWhiteSpace($wingetLocaleWorkaround)) {
                continue
            }

            # Install fallback: if all upgrade paths failed with 0x8A150014, try winget install --version.
            # The install command resolves packages against the source manifest rather than ARP entries,
            # bypassing the matching bug that causes upgrade to return "No packages found" (winget-cli #5249, #6095).
            if ((Test-AppMatch -AppId $AppId -PatternList @($wingetInstallFallbackAllowlist)) -and $exitCode -eq -1978335212 -and -not [string]::IsNullOrWhiteSpace($AvailableVersion)) {
                Write-Log "Retry: install fallback (winget install --version $AvailableVersion)" -Tag "Info"
                $installScopeIx = 0
                foreach ($scopeMode in $scopeOrder) {
                    if ($installScopeIx -gt 0) {
                        $sn = Get-WingetScopeShortName -ScopeMode $scopeMode
                        Write-Log "Retry: ${sn} scope (install)" -Tag "Info"
                    }
                    $installScopeIx++

                    $installArgs = @('install', '--id', $AppId, '-e', '--version', $AvailableVersion, '--force',
                        '--accept-package-agreements', '--accept-source-agreements', '--silent', '--disable-interactivity', '--skip-dependencies', '--source', 'winget')
                    if ($scopeMode -eq 'Machine') { $installArgs += '--scope', 'machine' }
                    elseif ($scopeMode -eq 'User') { $installArgs += '--scope', 'user' }
                    Write-Log "winget $($installArgs -join ' ')" -Tag "Debug"
                    $upgradeOutput = & $WingetPath @installArgs 2>&1 | Where-Object { $_ -notlike " *" }
                    $exitCode = $LASTEXITCODE
                    $exitInfo = Get-WingetExitCodeInfo -ExitCode $exitCode
                    $outputClaimsNoApplicable = Test-WingetUpgradeOutputClaimsNoApplicable -Lines $upgradeOutput

                    if (($exitInfo.Category -eq 'Success') -and -not $outputClaimsNoApplicable) {
                        $isn = Get-WingetScopeShortName -ScopeMode $scopeMode
                        Write-Log "$AppId ($isn, install fallback)" -Tag "Success"
                        return $true
                    }

                    if ($exitInfo.Category -in $deferCategories) {
                        Write-Log "Defer ${AppId}: $($exitInfo.Description)" -Tag "Info"
                        return $null
                    }

                    if (-not (Test-ShouldAdvanceScopeLadder -ExitInfo $exitInfo -ExitCode $exitCode -OutputClaimsNoApplicable $outputClaimsNoApplicable) -and $exitCode -ne -1978335212) {
                        break
                    }
                }
            }

            # Last resort: --uninstall-previous removes the old version first, then installs the new one.
            # Risky: if uninstall succeeds but install fails, the app is gone. Gated by $wingetUninstallPreviousAllowlist.
            if ((Test-AppMatch -AppId $AppId -PatternList @($wingetUninstallPreviousAllowlist)) -and -not [string]::IsNullOrWhiteSpace($AvailableVersion)) {
                Write-Log "Retry: uninstall-previous" -Tag "Info"
                $uninstScopeIx = 0
                foreach ($scopeMode in $scopeOrder) {
                    if ($uninstScopeIx -gt 0) {
                        $usn = Get-WingetScopeShortName -ScopeMode $scopeMode
                        Write-Log "Retry: ${usn} scope (uninstall-previous)" -Tag "Info"
                    }
                    $uninstScopeIx++

                    $uninstPrevArgs = @('upgrade', '--id', $AppId, '-e', '--version', $AvailableVersion, '--force',
                        '--uninstall-previous', '--accept-package-agreements', '--accept-source-agreements', '--silent', '--disable-interactivity', '--skip-dependencies', '--source', 'winget')
                    if ($scopeMode -eq 'Machine') { $uninstPrevArgs += '--scope', 'machine' }
                    elseif ($scopeMode -eq 'User') { $uninstPrevArgs += '--scope', 'user' }
                    Write-Log "winget $($uninstPrevArgs -join ' ')" -Tag "Debug"
                    $upgradeOutput = & $WingetPath @uninstPrevArgs 2>&1 | Where-Object { $_ -notlike " *" }
                    $exitCode = $LASTEXITCODE
                    $exitInfo = Get-WingetExitCodeInfo -ExitCode $exitCode
                    $outputClaimsNoApplicable = Test-WingetUpgradeOutputClaimsNoApplicable -Lines $upgradeOutput

                    if (($exitInfo.Category -eq 'Success') -and -not $outputClaimsNoApplicable) {
                        $usn2 = Get-WingetScopeShortName -ScopeMode $scopeMode
                        Write-Log "$AppId ($usn2, uninstall-previous)" -Tag "Success"
                        return $true
                    }

                    if ($exitInfo.Category -in $deferCategories) {
                        Write-Log "Defer ${AppId}: $($exitInfo.Description)" -Tag "Info"
                        return $null
                    }

                    if (-not (Test-ShouldAdvanceScopeLadder -ExitInfo $exitInfo -ExitCode $exitCode -OutputClaimsNoApplicable $outputClaimsNoApplicable)) {
                        break
                    }
                }
            }

            $errorMessages = $upgradeOutput | Where-Object {
                $_ -match 'error|failed|exception|unable|cannot|could not' -or
                ($_ -match '^[A-Z]' -and $_ -notmatch '^Loading|^Found|^Verifying|^Successfully')
            }
            if ($errorMessages) {
                Write-Log "Output ${AppId}: $($errorMessages -join '; ')" -Tag "Debug"
            }
            Write-Log "Fail ${AppId}: $($exitInfo.Description) ($($exitInfo.Category))" -Tag "Error"
            return $false
        }
    }
    catch {
        Write-Log "Upgrade $AppId error: $_" -Tag "Error"
        Write-Log "$($_.ScriptStackTrace)" -Tag "Debug"
        return $false
    }
}

# ---------------------------[ Main Remediation Logic ]---------------------------
try {
    if (-not (Test-Winget)) {
        Write-Log "Winget unavailable." -Tag "Error"
        Complete-Script -ExitCode 1 -PortalSummaryLine 'Updated: (none) | Failed: (none) | Winget unavailable'
    }

    try {
        $nl = Get-WingetScopeLadderOrderNormalized
        $scopeFlat = [System.Collections.Generic.List[string]]::new()
        foreach ($o in @($nl)) {
            if ($null -eq $o) { continue }
            if ($o -is [string]) {
                [void]$scopeFlat.Add($o)
            }
            elseif ($o -is [System.Array]) {
                foreach ($x in $o) {
                    if ($null -ne $x) { [void]$scopeFlat.Add([string]$x) }
                }
            }
            else {
                [void]$scopeFlat.Add([string]$o)
            }
        }
        $script:WingetScopeLadderNormalized = $scopeFlat.ToArray()
    }
    catch {
        Write-Log "Invalid wingetScopeLadderOrder: $_" -Tag "Error"
        Complete-Script -ExitCode 1 -PortalSummaryLine 'Updated: (none) | Failed: (none) | Invalid scope ladder config'
    }

    # One-time source refresh before we start
    $wingetPath = Get-WingetPath
    Write-Log "WinGet source update" -Tag "Run"
    & $wingetPath source update 2>&1 | Out-Null
    Write-Log "Sources OK" -Tag "Debug"

    if (Test-PendingReboot) {
        Write-Log "Reboot pending." -Tag "Info"
    }

    # Get available updates
    $availableUpdates = Get-AvailableUpdates

    if ($availableUpdates.Count -eq 0) {
        Write-Log "No upgrades" -Tag "Success"
        Complete-Script -ExitCode 0 -PortalSummaryLine (Build-RemediationPortalSummaryLine)
    }

    if ($listMode -eq 'Blacklist') {
        $listCount = if ($null -ne $blacklistApps) { $blacklistApps.Count } else { 0 }
        Write-Log "Mode: blacklist ($listCount)" -Tag "Info"
    }
    else {
        $listCount = if ($null -ne $whitelistApps) { $whitelistApps.Count } else { 0 }
        Write-Log "Mode: whitelist ($listCount)" -Tag "Info"
    }

    $filteredUpdates = Select-FilteredUpdates -Updates $availableUpdates -ListMode $listMode -Blacklist $blacklistApps -Whitelist $whitelistApps

    if ($filteredUpdates.Count -eq 0) {
        Write-Log "No upgrades after filter" -Tag "Success"
        Complete-Script -ExitCode 0 -PortalSummaryLine (Build-RemediationPortalSummaryLine)
    }

    # Perform updates
    $successCount   = 0
    $failureCount   = 0
    $deferredCount  = 0
    $succeededApps  = @()
    $failedApps     = @()
    $deferredApps   = @()

    $updateIndex = 0
    foreach ($update in $filteredUpdates) {
        $updateIndex++

        if (-not $update -or -not $update.AppId) {
            Write-Log "[$updateIndex/$($filteredUpdates.Count)] Invalid row" -Tag "Error"
            $failureCount++
            continue
        }

        Write-Log "[$updateIndex/$($filteredUpdates.Count)] $($update.AppId) $($update.CurrentVersion) -> $($update.AvailableVersion)" -Tag "Run"

        $result = Update-Application -AppId $update.AppId -WingetPath $wingetPath -AvailableVersion $update.AvailableVersion

        if ($result -eq $true) {
            $successCount++
            $succeededApps += $update.AppId
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
    Write-Log "OK count: $successCount" -Tag "Success"

    if ($deferredCount -gt 0) {
        Write-Log "Deferred ($deferredCount): $($deferredApps -join ', ')" -Tag "Info"
    }

    $portalLine = Build-RemediationPortalSummaryLine -Succeeded $succeededApps -Failed $failedApps -Deferred $deferredApps

    if ($failureCount -gt 0) {
        Write-Log "Failed ($failureCount): $($failedApps -join ', ')" -Tag "Error"
        Complete-Script -ExitCode 1 -PortalSummaryLine $portalLine
    }
    else {
        Write-Log "Done (no hard failures)" -Tag "Success"
        Complete-Script -ExitCode 0 -PortalSummaryLine $portalLine
    }
}
catch {
    Write-Log "Unhandled: $_" -Tag "Error"
    Write-Log "$($_.ScriptStackTrace)" -Tag "Debug"
    Complete-Script -ExitCode 1 -PortalSummaryLine 'Updated: (unknown) | Failed: (unknown) | see remediation.log'
}
