# ---------------------------[ UTF-8 / Encoding Normalization (WAU-style) ]---------------------------
try {
    $null = & "$env:WINDIR\System32\cmd.exe" /c ""      # prime console like WAU
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8BOM'
    $PSDefaultParameterValues['*:Encoding']        = 'utf8'
    $ProgressPreference = 'SilentlyContinue'
} catch {}

# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Script name ]---------------------------
$scriptName   = "Winget - App Update"
$logFileName  = "detection.log"

# ---------------------------[ Config ]---------------------------
$WingetSource         = 'winget'   # WAU also uses a source
$SkipUnknownInstalled = $true      # WAU-like: skip Installed='Unknown'

# Blacklist (Winget IDs; wildcards OK)
$ExcludeIds = @(
    'Microsoft.Edge*',
    'Microsoft.Teams*',    
    'Microsoft.Office',
    'Microsoft.OneDrive',
    'Microsoft.AppInstaller',
    'Microsoft.RemoteDesktopClient',
    'Microsoft.VCLibs.*',
    'Fortinet.FortiClientVPN',
    'Mozilla.Firefox*',
    'Opera.Opera*',
    'TeamViewer.TeamViewer*',
    'Brave.Brave*',
    'Adobe.Acrobat.Reader.32-bit',
    'Adobe.Acrobat.Reader.64-bit',
    'Microsoft.WindowsTerminal',
    'Microsoft.PowerShell'
)

# ---------------------------[ Logging Setup (append; consistent encoding) ]---------------------------
$log           = $true
$enableLogFile = $true
$logFileDirectory = "$env:ProgramData\IntuneLogs\Scripts\$scriptName"
$logFile          = "$logFileDirectory\$logFileName"

function Get-LogEncoding {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        if ($PSVersionTable.PSVersion.Major -ge 7) { return 'utf8BOM' } else { return 'utf8' }
    }
    try {
        $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 3 -ErrorAction Stop
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return 'utf8' }   # treat as UTF-8
        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return 'unicode' }                         # UTF-16 LE
        return 'utf8'
    } catch { return 'utf8' }
}

if ($enableLogFile -and -not (Test-Path $logFileDirectory)) { New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null }
$Script:LogEncoding = if ($enableLogFile) { Get-LogEncoding -Path $logFile } else { 'utf8' }

# Function to write structured logs to file and console  (your style)
function Write-Log {
    param ([string]$Message, [string]$Tag = "Info")

    if (-not $log) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList = @("Start", "Check", "Info", "Success", "Error", "Debug", "End")
    $rawTag = $Tag.Trim()

    if ($tagList -contains $rawTag) { $rawTag = $rawTag.PadRight(7) } else { $rawTag = "Error  " }

    $color = switch ($rawTag.Trim()) {
        "Start"{"Cyan"};"Check"{"Blue"};"Info"{"Yellow"};"Success"{"Green"};"Error"{"Red"};"Debug"{"DarkYellow"};"End"{"Cyan"} default{"White"}
    }

    $logMessage = "$timestamp [  $rawTag ] $Message"

    if ($enableLogFile) { Add-Content -Path $logFile -Value $logMessage -Encoding $Script:LogEncoding }

    Write-Host "$timestamp " -NoNewline
    Write-Host "[  " -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host " ] " -NoNewline -ForegroundColor White
    Write-Host "$Message"
}

# ---------------------------[ Exit Function ]---------------------------
function Complete-Script {
    param([int]$ExitCode)
    $duration = (Get-Date) - $scriptStartTime
    Write-Log "Script execution time: $($duration.ToString('hh\:mm\:ss\.ff'))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Detection Script Completed ========" -Tag "End"
    exit $ExitCode
}

# ---------------------------[ Winget health / path ]---------------------------
function Get-WingetPath {
    $wingetBase = "$env:ProgramW6432\WindowsApps"
    try {
        # Try x64 first
        $wingetFolders = Get-ChildItem -Path $wingetBase -Directory -ErrorAction Stop |
            Where-Object { $_.Name -like 'Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe' }

        # If x64 not found, try arm64
        if (-not $wingetFolders) {
            $wingetFolders = Get-ChildItem -Path $wingetBase -Directory -ErrorAction Stop |
                Where-Object { $_.Name -like 'Microsoft.DesktopAppInstaller_*_arm64__8wekyb3d8bbwe' }
        }

        if (-not $wingetFolders) {
            throw "No matching Winget installation folders found (x64 or arm64)."
        }

        $latestWingetFolder = $wingetFolders |
            Sort-Object CreationTime -Descending |
            Select-Object -First 1

        $wingetPath = Join-Path $latestWingetFolder.FullName 'winget.exe'

        if (-not (Test-Path $wingetPath)) {
            throw "winget.exe not found at expected location."
        }

        return $wingetPath
    }
    catch {
        Write-Log "Failed to detect Winget installation: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Register-WingetDependencyPaths {
    Write-Log "Registering Winget dependency DLL directories into SYSTEM PATH..." -Tag "Info"
    try {
        $windowsApps = "$env:ProgramW6432\WindowsApps"
        if (-not (Test-Path $windowsApps)) { Write-Log "WindowsApps folder not found: $windowsApps" -Tag "Error"; return }
        $dllPackages = @('Microsoft.VCLibs.140.00.UWPDesktop','Microsoft.UI.Xaml.2.8')

        $paths = New-Object System.Collections.Generic.List[string]
        foreach ($pkg in $dllPackages) {
            $folder = Get-ChildItem -Path $windowsApps -Directory |
                Where-Object { $_.Name -like "$pkg*_x64__*" } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1
            if ($folder) { [void]$paths.Add($folder.FullName); Write-Log "Found $pkg at: $($folder.FullName)" -Tag "Info" }
            else { Write-Log "Could not find folder for $pkg" -Tag "Error" }
        }
        $uniquePaths = $paths | Select-Object -Unique

        $currentPath = [Environment]::GetEnvironmentVariable("Path","Machine")
        $updated = $false
        foreach ($p in $uniquePaths) {
            if ($currentPath -notlike "*$p*") { $currentPath += ";$p"; $updated = $true; Write-Log "Adding path to SYSTEM PATH: $p" -Tag "Info" }
            else { Write-Log "Path already present: $p" -Tag "Debug" }
        }
        if ($updated) { [Environment]::SetEnvironmentVariable("Path",$currentPath,"Machine"); Write-Log "SYSTEM PATH updated." -Tag "Success" }
        else { Write-Log "No changes made to SYSTEM PATH." -Tag "Info" }
    } catch { Write-Log "Dependency PATH registration failed: $_" -Tag "Error" }
}

function Invoke-WingetRepair {
    Write-Log "Attempting Winget repair..." -Tag "Info"
    Register-WingetDependencyPaths
    Write-Log "Repair complete (restart may be required)." -Tag "Info"
    Complete-Script -ExitCode 0
}

function Test-Winget {
    $wg = Get-WingetPath
    $v = & $wg -v 2>$null
    if ($LASTEXITCODE -eq 0 -and $v) { Write-Log "Winget version: $v" -Tag "Info"; return $true }
    Write-Log "Winget not healthy (exit $LASTEXITCODE)." -Tag "Error"
    return $false
}

# ---------------------------[ WAU-style parsing function ]---------------------------
function Get-WingetOutdatedApps {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Source
    )

    class Software { [string]$Name; [string]$Id; [string]$Version; [string]$AvailableVersion }

    $wg = Get-WingetPath

    # Get list of available upgrades; drop indented/progress lines (WAU behavior)
    try {
        $raw = & $wg upgrade --source $Source | Where-Object { $_ -notlike "   *" } | Out-String
    } catch {
        Write-Log "Error while receiving winget upgrade list: $_" -Tag "Error"
        $raw = $null
    }

    # "No update found" handling like WAU
    if (-not ($raw -match "-----")) {
        return "No update found. 'winget upgrade' output:`n$raw"
    }

    # Split to lines
    $lines = $raw.Split([Environment]::NewLine) | Where-Object { $_ }

    # Find the dashed separator and header line
    $fl = 0
    while (-not $lines[$fl].StartsWith("-----")) { $fl++ }
    $fl = $fl - 1

    # Compute column boundaries with CJK compensation (WAU)
    $index = $lines[$fl] -split '(?<=\s)(?!\s)'
    $idStart        = ($index[0] -replace '[\u4e00-\u9fa5]', '**').Length
    $versionStart   = $idStart        + (($index[1] -replace '[\u4e00-\u9fa5]', '**').Length)
    $availableStart = $versionStart   + (($index[2] -replace '[\u4e00-\u9fa5]', '**').Length)

    # Parse rows
    $upgradeList = @()
    for ($i = $fl + 2; $i -lt $lines.Length; $i++) {
        $line = $lines[$i] -replace "[\u2026]", " "  # normalize ellipsis
        if ($line.StartsWith("-----")) {
            # header changed -> recompute columns
            $fl = $i - 1
            $index = $lines[$fl] -split '(?<=\s)(?!\s)'
            $idStart        = ($index[0] -replace '[\u4e00-\u9fa5]', '**').Length
            $versionStart   = $idStart        + (($index[1] -replace '[\u4e00-\u9fa5]', '**').Length)
            $availableStart = $versionStart   + (($index[2] -replace '[\u4e00-\u9fa5]', '**').Length)
            continue
        }
        # apps-only heuristic (WAU)
        if ($line -match "\w\.\w") {
            $nameDecl = (($line.Substring(0, $idStart) -replace '[\u4e00-\u9fa5]', '**').Length) - ($line.Substring(0, $idStart).Length)
            $soft = [Software]::new()
            $soft.Name             = $line.Substring(0, $idStart - $nameDecl).TrimEnd()
            $soft.Id               = $line.Substring($idStart - $nameDecl, $versionStart - $idStart).TrimEnd()
            $soft.Version          = $line.Substring($versionStart - $nameDecl, $availableStart - $versionStart).TrimEnd()
            $soft.AvailableVersion = $line.Substring($availableStart - $nameDecl).TrimEnd()   # may include " winget" and that's fine
            $upgradeList += $soft
        }
    }

    return $upgradeList | Sort-Object { Get-Random }
}

function Test-IdInList { param([string]$Id,[string[]]$List) foreach ($pat in $List) { if ([string]::IsNullOrWhiteSpace($pat)) { continue } if ($Id -like $pat) { return $true } } return $false }

# ---------------------------[ Script Start ]---------------------------
Write-Log "======== Detection Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"

if (-not (Test-Winget)) {
    Write-Log "Winget unhealthy; attempting repair." -Tag "Info"
    Invoke-WingetRepair
    if (-not (Test-Winget)) { Write-Log "Winget still unhealthy after repair." -Tag "Error"; Complete-Script -ExitCode 1 }
}

$outdated = Get-WingetOutdatedApps -Source $WingetSource

# WAU's "No update found..." string path
if ($outdated -is [string]) {
    Write-Log $outdated -Tag "Success"
    Complete-Script -ExitCode 0
}

# Filter Unknown + blacklist
$eligible = foreach ($app in $outdated) {
    if ($SkipUnknownInstalled -and $app.Version -eq 'Unknown') { continue }
    if (Test-IdInList -Id $app.Id -List $ExcludeIds) { continue }
    $app
}

if ($eligible.Count -gt 0) {
    foreach ($a in $eligible) {
        Write-Log "Update available: $($a.Id) $($a.Version) -> $($a.AvailableVersion)" -Tag "Info"
    }
    Complete-Script -ExitCode 1
} else {
    Write-Log "Updates exist but none eligible after filters (Unknown/Blacklist)." -Tag "Success"
    Complete-Script -ExitCode 0
}







