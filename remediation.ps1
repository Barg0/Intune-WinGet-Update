# ---------------------------[ Config ]---------------------------
$blacklistApps = @(
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
    'Microsoft.WindowsTerminal',
    'Adobe.Acrobat.Reader.32-bit',
    'Adobe.Acrobat.Reader.64-bit',
    'Microsoft.PowerShell'
)

# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Script Name ]---------------------------
$scriptName  = "Winget-AppUpdate-Blacklist"
$logFileName = "remediation.log"

# ---------------------------[ Logging Setup ]---------------------------
# Logging configuration
$log           = $true
$logDebug      = $true    # Set to $true for verbose DEBUG logging
$logGet        = $true    # enable/disable all [Get] logs
$logRun        = $true    # enable/disable all [Run] logs
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

    # Per-tag switches
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
        catch {
            # Logging must never block script execution
        }
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
    
    # WAU approach: Use Get-Item with wildcard and sort by FileVersionRaw
    $systemPath = "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_8wekyb3d8bbwe\winget.exe"
    
    try {
        # Try system context first (newest version) - matches WAU Get-WingetCmd
        $WingetInfo = (Get-Item $systemPath -ErrorAction Stop).VersionInfo |
            Sort-Object FileVersionRaw -Descending |
            Select-Object -First 1

        if ($WingetInfo.FileName) {
            return $WingetInfo.FileName
        }
    }
    catch {
        # System context not found, try user context (WAU fallback)
        $userPath = "$env:LocalAppData\Microsoft\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe"
        if (Test-Path $userPath) {
            return $userPath
        }
        
        Write-Log "Failed to detect Winget installation: $_" -Tag "Error"
        throw "Winget not found in system or user context"
    }
    
    throw "Winget not found"
}

# ---------------------------[ Test Winget Function ]---------------------------
function Test-Winget {
    [CmdletBinding()]
    param()
    
    Write-Log "Checking Winget availability" -Tag "Get"
    
    try {
        $wingetPath = Get-WingetPath
        # Get version - capture all output first
        $rawOutput = & $wingetPath -v 2>&1
        $exitCode = $LASTEXITCODE

        # Check exit code first - if 0, winget executed successfully
        if ($exitCode -eq 0) {
            # Extract version from output (handle various formats: "1.2.3", "v1.2.3", "winget version 1.2.3", etc.)
            $versionLine = $rawOutput | Where-Object { $_ -and ($_ -match '\d+\.\d+') } | Select-Object -First 1
            if ($versionLine) {
                # Extract version number (first match of pattern)
                if ($versionLine -match '(\d+\.\d+(?:\.\d+)?(?:\.\d+)?)') {
                    $version = $matches[1]
                    Write-Log "Winget version: $version" -Tag "Success"
                    return $true
                }
                else {
                    # If we can't parse version but exit code was 0, still consider it success
                    Write-Log "Winget is available (version output: $($versionLine.Trim()))" -Tag "Success"
                    return $true
                }
            }
            else {
                # Exit code 0 but no version found - still consider success (winget works)
                Write-Log "Winget is available (execution successful)" -Tag "Success"
                return $true
            }
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
        if ($pattern -match '^\*') {
            # Wildcard at start: *Firefox
            $suffix = $pattern.TrimStart('*')
            if ($AppId -like "*$suffix") {
                return $true
            }
        }
        elseif ($pattern -match '\*$') {
            # Wildcard at end: Firefox*
            $prefix = $pattern.TrimEnd('*')
            if ($AppId -like "$prefix*") {
                return $true
            }
        }
        elseif ($pattern -match '\*') {
            # Wildcard in middle: Fire*fox
            $regexPattern = $pattern -replace '\*', '.*'
            if ($AppId -match "^$regexPattern$") {
                return $true
            }
        }
        else {
            # Exact match
            if ($AppId -eq $pattern) {
                return $true
            }
        }
    }
    
    return $false
}

# ---------------------------[ Get Available Updates ]---------------------------
function Get-AvailableUpdates {
    [CmdletBinding()]
    param()
    
    Write-Log "Checking for available updates" -Tag "Get"
    $wingetPath = Get-WingetPath
    
    try {
        # Set UTF-8 encoding to properly handle Unicode characters (like ellipsis \u2026)
        # This prevents encoding issues where characters appear as garbled text (e.g., ΓÇª)
        $previousOutputEncoding = [Console]::OutputEncoding
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
        
        try {
            # WAU uses: winget upgrade --source winget (NO accept flags for listing)
            # Filter out lines starting with space (progress indicators) - matches WAU exactly
            # WAU does NOT redirect stderr (no 2>&1)
            $upgradeResult = & $wingetPath upgrade --source winget | 
                Where-Object { $_ -notlike " *" } | 
                Out-String
        }
        finally {
            # Restore previous encoding
            [Console]::OutputEncoding = $previousOutputEncoding
        }
        
        # WAU checks for separator line - doesn't check exit code!
        # Check if output contains valid data (header separator line with dashes)
        if (-not ($upgradeResult -match "-----")) {
            Write-Log "No update found. Winget upgrade output does not contain table separator." -Tag "Info"
            return @()
        }
        
        # Split output into lines, removing empty lines
        $lines = $upgradeResult.Split([Environment]::NewLine) | Where-Object { $_ }
        
        # Replace ellipsis characters (\u2026) - WAU does this to handle long names
        # With proper UTF-8 encoding, this should now work correctly
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $lines[$i] = $lines[$i] -replace "[\u2026]", " "
        }
        
        # Find the separator line (starts with "-----")
        $fl = 0
        while ($fl -lt $lines.Count -and -not $lines[$fl].StartsWith("-----")) {
            $fl++
        }
        
        if ($fl -ge $lines.Count) {
            Write-Log "Could not find table separator in winget output" -Tag "Error"
            return @()
        }
        
        # Get header line (one line before separator)
        $fl = $fl - 1
        if ($fl -lt 0) {
            Write-Log "Could not find header line in winget output" -Tag "Error"
            return @()
        }
        
        # Split header into columns (preserving trailing spaces for positioning)
        # WAU uses this exact regex to split on space boundaries
        $index = $lines[$fl] -split '(?<=\s)(?!\s)'
        
        # Ensure we have at least 3 columns (Name, Id, Version, Available)
        if ($index.Count -lt 3) {
            Write-Log "Invalid header format - expected at least 3 columns" -Tag "Error"
            return @()
        }
        
        # Calculate column positions (handle non-Latin characters by replacing with **)
        # WAU calculates these positions exactly like this
        $idStart = $($index[0] -replace '[\u4e00-\u9fa5]', '**').Length
        $versionStart = $idStart + $($index[1] -replace '[\u4e00-\u9fa5]', '**').Length
        $availableStart = $versionStart + $($index[2] -replace '[\u4e00-\u9fa5]', '**').Length
        
        # Parse each data line
        $updates = @()
        $unknownCount = 0
        
        For ($i = $fl + 2; $i -lt $lines.Count; $i++) {
            # Ellipsis already replaced earlier, just get the line
            $line = $lines[$i]
            
            # Handle multiple tables (new header encountered)
            if ($line.StartsWith("-----")) {
                $fl = $i - 1
                $index = $lines[$fl] -split '(?<=\s)(?!\s)'
                $idStart = $($index[0] -replace '[\u4e00-\u9fa5]', '**').Length
                $versionStart = $idStart + $($index[1] -replace '[\u4e00-\u9fa5]', '**').Length
                $availableStart = $versionStart + $($index[2] -replace '[\u4e00-\u9fa5]', '**').Length
                continue
            }
            
            # Check if line contains an application entry (has format word.word)
            if ($line -match "\w\.\w") {
                # Calculate name declination for non-Latin character handling
                # WAU uses this exact calculation
                $nameDeclination = $($line.Substring(0, $idStart) -replace '[\u4e00-\u9fa5]', '**').Length - $line.Substring(0, $idStart).Length
                
                # Extract values using WAU's exact substring logic - ONLY TrimEnd(), no additional cleaning
                # WAU does NOT do any additional Trim(), character removal, or validation
                $appName = $line.Substring(0, $idStart - $nameDeclination).TrimEnd()
                $appId = $line.Substring($idStart - $nameDeclination, $versionStart - $idStart).TrimEnd()
                $currentVersion = $line.Substring($versionStart - $nameDeclination, $availableStart - $versionStart).TrimEnd()
                $availableVersion = $line.Substring($availableStart - $nameDeclination).TrimEnd()
                
                # Skip apps with "Unknown" version (WAU behavior)
                if ($currentVersion -eq "Unknown" -or $availableVersion -eq "Unknown") {
                    $unknownCount++
                    Write-Log "Skipping app with Unknown version: $appId" -Tag "Debug"
                    continue
                }
                
                # Skip if versions are the same
                if ($currentVersion -ne $availableVersion) {
                    $updates += @{
                        AppId            = $appId
                        AppName          = $appName
                        CurrentVersion   = $currentVersion
                        AvailableVersion = $availableVersion
                    }
                }
            }
        }
        
        if ($unknownCount -gt 0) {
            Write-Log "Skipped $unknownCount app(s) with Unknown version" -Tag "Debug"
        }
        
        Write-Log "Found $($updates.Count) apps with available updates" -Tag "Get"
        return $updates
    }
    catch {
        Write-Log "Error getting available updates: $_" -Tag "Error"
        Write-Log $_.ScriptStackTrace -Tag "Debug"
        return @()
    }
}

# ---------------------------[ Filter Updates by Blacklist ]---------------------------
function Select-FilteredUpdates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Updates,
        
        [Parameter(Mandatory = $true)]
        [string[]]$Blacklist
    )
    
    if ($null -eq $Updates -or $Updates.Count -eq 0) {
        return @()
    }
    
    $filteredUpdates = @()
    
    foreach ($update in $Updates) {
        # Validate update object structure
        if (-not $update -or -not $update.AppId) {
            Write-Log "Invalid update object encountered, skipping" -Tag "Debug"
            continue
        }
        
        $appId = $update.AppId
        
        # Exclude apps in blacklist
        if ($null -ne $Blacklist -and $Blacklist.Count -gt 0) {
            if (Test-AppMatch -AppId $appId -PatternList $Blacklist) {
                Write-Log "App excluded (in blacklist): $appId" -Tag "Debug"
                continue
            }
        }
        
        $filteredUpdates += $update
    }
    
    Write-Log "Filtered to $($filteredUpdates.Count) apps requiring updates" -Tag "Get"
    return $filteredUpdates
}

# ---------------------------[ Update Application ]---------------------------
function Update-Application {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppId,
        
        [Parameter(Mandatory = $true)]
        [string]$WingetPath
    )
    
    Write-Log "Updating application: $AppId" -Tag "Debug"
    
    try {
        # Execute winget upgrade - WAU uses: winget upgrade --id <AppId> -e --accept-package-agreements --accept-source-agreements -s winget -h
        # -e = exact match, -h = silent mode, -s winget = source
        # Filter out lines starting with space (progress indicators) - matches WAU behavior
        $upgradeOutput = & $WingetPath upgrade --id $AppId `
            -e `
            --accept-package-agreements `
            --accept-source-agreements `
            -s winget `
            -h 2>&1 | 
            Where-Object { $_ -notlike " *" }
        
        $exitCode = $LASTEXITCODE
        
        # Ensure we have an array even if filtering returns null
        if ($null -eq $upgradeOutput) {
            $upgradeOutput = @()
        }
        
        # WAU checks installation confirmation - we'll check exit code and verify app is updated
        if ($exitCode -eq 0) {
            Write-Log "Successfully updated: $AppId" -Tag "Success"
            return $true
        }
        else {
            Write-Log "Failed to update $AppId - Exit code: $exitCode" -Tag "Error"
            # Only log actual error messages, filter out noise
            $errorMessages = $upgradeOutput | Where-Object { 
                $_ -match 'error|failed|exception|unable|cannot|could not' -or 
                ($_ -match '^[A-Z]' -and $_ -notmatch '^Loading|^Found|^Verifying')
            }
            if ($errorMessages) {
                Write-Log "Error: $($errorMessages -join '; ')" -Tag "Debug"
            }
            return $false
        }
    }
    catch {
        Write-Log "Error updating $AppId : $_" -Tag "Error"
        Write-Log $_.ScriptStackTrace -Tag "Debug"
        return $false
    }
}

# ---------------------------[ Main Remediation Logic ]---------------------------
try {
    # Test Winget availability
    if (-not (Test-Winget)) {
        Write-Log "Winget is not available or not working properly" -Tag "Error"
        Complete-Script -ExitCode 1
        return
    }
    
    # Get available updates
    $availableUpdates = Get-AvailableUpdates
    
    if ($availableUpdates.Count -eq 0) {
        Write-Log "No updates available - all apps are up to date" -Tag "Success"
        Complete-Script -ExitCode 0
    }
    
    # Filter updates based on blacklist
    if ($null -ne $blacklistApps -and $blacklistApps.Count -gt 0) {
        Write-Log "Using blacklist with $($blacklistApps.Count) entries" -Tag "Info"
    }
    else {
        Write-Log "No blacklist configured - all apps will be updated" -Tag "Info"
    }
    
    $filteredUpdates = Select-FilteredUpdates -Updates $availableUpdates -Blacklist $blacklistApps
    
    if ($filteredUpdates.Count -eq 0) {
        Write-Log "No updates needed after filtering - all managed apps are up to date" -Tag "Success"
        Complete-Script -ExitCode 0
    }
    
    # Get Winget path for updates (reuse to avoid multiple calls)
    $wingetPath = Get-WingetPath
    
    # Perform updates
    Write-Log "Starting update process for $($filteredUpdates.Count) application(s)" -Tag "Run"
    
    $successCount = 0
    $failureCount = 0
    $failedApps = @()
    
    $updateIndex = 0
    foreach ($update in $filteredUpdates) {
        $updateIndex++
        
        # Validate update object before processing
        if (-not $update -or -not $update.AppId) {
            Write-Log "[$updateIndex/$($filteredUpdates.Count)] Skipping invalid update object" -Tag "Error"
            $failureCount++
            continue
        }
        
        Write-Log "[$updateIndex/$($filteredUpdates.Count)] Processing: $($update.AppId) ($($update.CurrentVersion) -> $($update.AvailableVersion))" -Tag "Info"
        
        if (Update-Application -AppId $update.AppId -WingetPath $wingetPath) {
            $successCount++
        }
        else {
            $failureCount++
            $failedApps += $update.AppId
        }
        
        # Small delay between updates to avoid overwhelming the system
        Start-Sleep -Seconds 2
    }
    
    # Summary
    Write-Log "Update process completed" -Tag "Debug"
    Write-Log "Successfully updated: $successCount application(s)" -Tag "Success"
    
    if ($failureCount -gt 0) {
        Write-Log "Failed to update: $failureCount application(s)" -Tag "Error"
        Write-Log "Failed apps: $($failedApps -join ', ')" -Tag "Error"
        Complete-Script -ExitCode 1
    }
    else {
        Write-Log "All updates completed successfully" -Tag "Success"
        Complete-Script -ExitCode 0
    }
}
catch {
    Write-Log "Unexpected error in remediation script: $_" -Tag "Error"
    Write-Log $_.ScriptStackTrace -Tag "Debug"
    Complete-Script -ExitCode 1
}
