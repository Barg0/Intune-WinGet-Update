# ---------------------------[ Config ]---------------------------
# ListMode: 'Blacklist' = update all except listed; 'Whitelist' = update only listed (must match remediation.ps1)
$listMode = 'Blacklist'

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
    'Microsoft.WindowsTerminal',
    'Adobe.Acrobat.Pro',
    'Adobe.Acrobat.Reader.32-bit',
    'Adobe.Acrobat.Reader.64-bit',
    'Microsoft.PowerShell',
    'Lenovo.SUHelper'
)

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
$logFileName = "detection.log"

# ---------------------------[ Logging Setup ]---------------------------
$log           = $true
$logDebug      = $true
$logGet        = $true
$logRun        = $true
$enableLogFile = $true

$logFileDirectory = "$env:ProgramData\IntuneLogs\Scripts\$scriptName"
$logFile          = "$logFileDirectory\$logFileName"

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

# ---------------------------[ Main Detection Logic ]---------------------------
try {
    if (-not (Test-Winget)) {
        Write-Log "Winget is not available; detection cannot proceed." -Tag "Error"
        Complete-Script -ExitCode 1
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

    Write-Log "Apps requiring updates:" -Tag "Info"
    foreach ($update in $filteredUpdates) {
        $scopeTag = if ($update.Scope -eq 'user') { ' [user]' } else { '' }
        Write-Log "  - $($update.AppId): $($update.CurrentVersion) -> $($update.AvailableVersion)$scopeTag" -Tag "Info"
    }

    Write-Log "Detection complete: $($filteredUpdates.Count) app(s) need updating" -Tag "Success"
    Complete-Script -ExitCode 1
}
catch {
    Write-Log "Unexpected error in detection script: $_" -Tag "Error"
    Write-Log $_.ScriptStackTrace -Tag "Debug"
    Complete-Script -ExitCode 1
}
