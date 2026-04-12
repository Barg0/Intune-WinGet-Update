# ---------------------------[ Config ]---------------------------
# ListMode: 'Blacklist' = update all except listed; 'Whitelist' = update only listed (must match remediation.ps1)
$listMode = 'Blacklist'

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
$scriptName  = 'WinGet-Update'
$logFileName = "detection.log"

# ---------------------------[ Logging Setup ]---------------------------
$log           = $true
$logDebug      = $false
$logGet        = $true
$logRun        = $true
$enableLogFile = $true

$logFileDirectory = "$env:ProgramData\IntuneLogs\Scripts\$scriptName"
$logFile          = "$logFileDirectory\$logFileName"

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

function Format-AvailableUpdateSummaryLine {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Update
    )
    return "$($Update.AppId) $($Update.CurrentVersion) -> $($Update.AvailableVersion)"
}

# Multiline summary for Intune portal (final console output). Not written to the log file.
function Build-DetectionPortalSummaryLine {
    param(
        [object[]]$Updates = @(),
        [string]$Note = ''
    )
    $arr = @($Updates)
    if ($arr.Count -gt 0) {
        $lines = @('Available:') + ($arr | ForEach-Object { Format-AvailableUpdateSummaryLine -Update $_ })
        $result = $lines -join [Environment]::NewLine
        if (-not [string]::IsNullOrWhiteSpace($Note)) {
            $result += [Environment]::NewLine + '| ' + $Note
        }
        return $result
    }

    $line = 'Available: (none)'
    if (-not [string]::IsNullOrWhiteSpace($Note)) {
        $line += " | $Note"
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
# Three calls: no --scope, then --scope user, then --scope machine. Union by AppId (first-seen wins).
function Get-AvailableUpdates {
    [CmdletBinding()]
    param()

    Write-Log "Upgrades: list (default, user, machine)" -Tag "Debug"
    $wingetPath = Get-WingetPath

    try {
        $previousOutputEncoding = [Console]::OutputEncoding
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

        $allUpdates = @()

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

# ---------------------------[ Main Detection Logic ]---------------------------
try {
    if (-not (Test-Winget)) {
        Write-Log "Winget unavailable." -Tag "Error"
        Complete-Script -ExitCode 0 -PortalSummaryLine (Build-DetectionPortalSummaryLine -Note 'Winget unavailable')
    }

    # Get available updates
    $availableUpdates = Get-AvailableUpdates

    if ($availableUpdates.Count -eq 0) {
        Write-Log "No upgrades" -Tag "Success"
        Complete-Script -ExitCode 0 -PortalSummaryLine (Build-DetectionPortalSummaryLine)
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
        Complete-Script -ExitCode 0 -PortalSummaryLine (Build-DetectionPortalSummaryLine)
    }

    Write-Log "Available:" -Tag "Info"
    foreach ($update in $filteredUpdates) {
        Write-Log "  $(Format-AvailableUpdateSummaryLine -Update $update)" -Tag "Info"
    }

    Write-Log "Detect: $($filteredUpdates.Count) non-compliant" -Tag "Success"
    Complete-Script -ExitCode 1 -PortalSummaryLine (Build-DetectionPortalSummaryLine -Updates $filteredUpdates)
}
catch {
    Write-Log "Unhandled: $_" -Tag "Error"
    Write-Log "$($_.ScriptStackTrace)" -Tag "Debug"
    Complete-Script -ExitCode 1 -PortalSummaryLine 'Available: (unknown) | see detection.log'
}
