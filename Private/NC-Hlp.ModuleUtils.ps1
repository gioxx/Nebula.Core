#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Module helpers =============================================================================================================

function Format-OutputString {
    <#
    .SYNOPSIS
        Truncates a string to a maximum length.
    .DESCRIPTION
        Returns the original value when shorter than the specified length; otherwise appends ellipsis.
    .PARAMETER Value
        String to trim.
    .PARAMETER MaxLength
        Maximum allowed length (defaults to module configuration).
    #>
    [CmdletBinding()]
    param(
        [string]$Value,
        [ValidateRange(3, 260)]
        [int]$MaxLength = $NCVars.MaxFieldLength
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return $Value
    }

    if ($Value.Length -le $MaxLength) {
        return $Value
    }

    $length = [Math]::Max(3, $MaxLength)
    return $Value.Substring(0, $length - 3) + '...'
}

function Format-NCDateTime {
    <#
    .SYNOPSIS
        Formats a date value using Nebula.Core conventions.
    .DESCRIPTION
        Converts a date-like value to a string using the configured Nebula date format.
        Returns the original value when it cannot be parsed as a date.
    .PARAMETER Value
        Date value to format.
    .PARAMETER Format
        Target string format. Defaults to the configured full date/time format.
    .PARAMETER AsLocalTime
        Convert the value to local time before formatting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Value,
        [string]$Format = $NCVars.DateTimeString_Full,
        [switch]$AsLocalTime
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $dateTimeOffset = $null
    $parsed = $false

    if ($Value -is [datetimeoffset]) {
        $dateTimeOffset = [datetimeoffset]$Value
        $parsed = $true
    }
    elseif ($Value -is [datetime]) {
        $dateTimeOffset = [datetimeoffset]::new([datetime]$Value)
        $parsed = $true
    }
    else {
        $text = [string]$Value
        $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces -bor [System.Globalization.DateTimeStyles]::RoundtripKind
        $formats = @(
            $Format,
            'o',
            'O',
            's',
            'yyyy-MM-ddTHH:mm:ssK',
            'yyyy-MM-ddTHH:mm:ss.FFFFFFFK',
            'yyyy-MM-dd HH:mm:ssK'
        )

        foreach ($fmt in $formats) {
            if ([datetimeoffset]::TryParseExact($text, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$dateTimeOffset)) {
                $parsed = $true
                break
            }
        }

        if (-not $parsed) {
            if ([datetimeoffset]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$dateTimeOffset)) {
                $parsed = $true
            }
            elseif ([datetimeoffset]::TryParse($text, [System.Globalization.CultureInfo]::CurrentCulture, $styles, [ref]$dateTimeOffset)) {
                $parsed = $true
            }
        }

        if (-not $parsed) {
            return [string]$Value
        }
    }

    $targetTimeZoneId = $NCVars.DateTimeTimeZone
    if (-not [string]::IsNullOrWhiteSpace([string]$targetTimeZoneId)) {
        $timeZoneInfo = Get-NCDateTimeZoneInfo -TimeZoneId $targetTimeZoneId
        if ($timeZoneInfo) {
            $dateTimeOffset = [System.TimeZoneInfo]::ConvertTime($dateTimeOffset, $timeZoneInfo)
        }
    }
    elseif ($AsLocalTime) {
        $dateTimeOffset = $dateTimeOffset.ToLocalTime()
    }

    return $dateTimeOffset.ToString($Format)
}

function Get-NCDateTimeZoneInfo {
    <#
    .SYNOPSIS
        Resolves a time zone ID for Nebula.Core date formatting.
    .DESCRIPTION
        Tries the provided ID first, then a small alias map for common cross-platform zone names.
    .PARAMETER TimeZoneId
        Time zone identifier from configuration.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TimeZoneId
    )

    if ([string]::IsNullOrWhiteSpace($TimeZoneId)) {
        return $null
    }

    if ($script:NCTimeZoneCache -and $script:NCTimeZoneCache.TimeZoneId -eq $TimeZoneId) {
        return $script:NCTimeZoneCache.TimeZoneInfo
    }

    $candidateIds = [System.Collections.Generic.List[string]]::new()
    $candidateIds.Add($TimeZoneId.Trim())

    $aliasMap = @{
        'Europe/Rome' = 'W. Europe Standard Time'
    }

    if ($aliasMap.ContainsKey($TimeZoneId.Trim())) {
        $candidateIds.Add($aliasMap[$TimeZoneId.Trim()])
    }

    foreach ($candidateId in $candidateIds) {
        try {
            $timeZoneInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($candidateId)
            $script:NCTimeZoneCache = [pscustomobject]@{
                TimeZoneId   = $TimeZoneId
                TimeZoneInfo = $timeZoneInfo
            }
            return $timeZoneInfo
        }
        catch {
            continue
        }
    }

    Write-NCMessage "Unable to resolve time zone '$TimeZoneId'. Falling back to local time." -Level WARNING
    $script:NCTimeZoneCache = [pscustomobject]@{
        TimeZoneId   = $TimeZoneId
        TimeZoneInfo = $null
    }
    return $null
}

function Get-NormalizedText {
    <#
    .SYNOPSIS
        Returns a lower-cased, trimmed string from a value or object.
    .DESCRIPTION
        Accepts plain strings, deserialized Exchange objects, and other objects with useful identity
        properties. Falls back to the object's string representation and returns $null for blank values.
    .PARAMETER Value
        Value to normalize.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Value
    )

    $text = $null

    if ($Value -is [string]) {
        $text = $Value
    }
    elseif ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = foreach ($item in $Value) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item)) {
                [string]$item
            }
        }

        if ($items) {
            $text = $items -join ', '
        }
    }
    else {
        foreach ($propertyName in @(
            'PrimarySmtpAddress',
            'UserPrincipalName',
            'WindowsEmailAddress',
            'SmtpAddress',
            'EmailAddress',
            'Value',
            'Name',
            'DisplayName',
            'Identity',
            'RawIdentity'
        )) {
            if ($Value.PSObject.Properties.Match($propertyName).Count -gt 0) {
                $candidate = $Value.$propertyName
                if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
                    $text = [string]$candidate
                    break
                }
            }
        }
    }

    if (-not $text) {
        $text = [string]$Value
    }

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return $text.Trim().ToLowerInvariant()
}

function Invoke-NCRetry {
    <#
    .SYNOPSIS
        Executes a scriptblock with retry logic.
    .DESCRIPTION
        Runs the provided block up to MaxAttempts, invoking OnError between retries.
        Throws the last error once all attempts are exhausted.
    .PARAMETER Action
        Script block to execute.
    .PARAMETER MaxAttempts
        Maximum number of attempts before throwing (default 3).
    .PARAMETER DelaySeconds
        Pause between attempts (default 5 seconds).
    .PARAMETER OperationDescription
        Friendly description used in log messages.
    .PARAMETER OnError
        Optional callback invoked after each failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxAttempts = 3,
        [ValidateRange(0, [int]::MaxValue)]
        [int]$DelaySeconds = 5,
        [string]$OperationDescription = 'operation',
        [scriptblock]$OnError
    )

    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            return & $Action
        }
        catch {
            if ($OnError) {
                $safeAttempt = if ($attempt -gt 0) { $attempt } else { 1 }
                $safeMax = if ($MaxAttempts -gt 0) { $MaxAttempts } else { 1 }
                & $OnError $safeAttempt $safeMax $_
            }
            else {
                Write-NCMessage "Operation '$OperationDescription' failed (attempt $attempt of $MaxAttempts). $($_.Exception.Message)" -Level ERROR
            }

            if ($attempt -ge $MaxAttempts) {
                throw
            }

            if ($DelaySeconds -gt 0) {
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
}

function New-File {
    <#
    .SYNOPSIS
        Generates a non-colliding file path.
    .DESCRIPTION
        Given a desired path, appends _N before the extension until an unused name is found.
    .PARAMETER Path
        Desired output file path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)
    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if (-not $directory) {
        $directory = (Get-Location).ProviderPath
    }

    $candidate = Join-Path -Path $directory -ChildPath ($baseName + $extension)
    $count = 1
    while (Test-Path -LiteralPath $candidate) {
        $fileName = "{0}_{1}{2}" -f $baseName, $count, $extension
        $candidate = Join-Path -Path $directory -ChildPath $fileName
        $count++
    }

    return $candidate
}

function Restore-ProgressAndInfoPreferences {
    <#
    .SYNOPSIS
        Restores Information/Progress preference variables.
    .DESCRIPTION
        Reverts preference variables previously captured by Set-ProgressAndInfoPreferences.
        No-ops if nothing was captured.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:PreferencesCaptured) {
        return
    }

    if ($null -ne $script:PreviousInformationPreference) {
        Set-Variable -Name InformationPreference -Value $script:PreviousInformationPreference -Scope Global
    }

    if ($null -ne $script:PreviousProgressPreference) {
        Set-Variable -Name ProgressPreference -Value $script:PreviousProgressPreference -Scope Global
    }

    $script:PreferencesCaptured = $false
    $script:PreviousInformationPreference = $null
    $script:PreviousProgressPreference = $null
}

function Set-ProgressAndInfoPreferences {
    <#
    .SYNOPSIS
        Forces Information/Progress preference variables to Continue.
    .DESCRIPTION
        Saves current preference values (once per session) and sets global
        InformationPreference and ProgressPreference to Continue for verbose output.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:PreferencesCaptured) {
        $script:PreviousInformationPreference = $InformationPreference
        $script:PreviousProgressPreference = $ProgressPreference
        $script:PreferencesCaptured = $true
    }

    Set-Variable -Name InformationPreference -Value Continue -Scope Global
    Set-Variable -Name ProgressPreference -Value Continue -Scope Global
}

function Get-NCProgressPercent {
    <#
    .SYNOPSIS
        Calculates a percentage for progress reporting.
    .DESCRIPTION
        Returns a rounded percentage for the current work item against a total.
        The total is clamped to at least 1 to avoid divide-by-zero errors.
    .PARAMETER Current
        Current work item count.
    .PARAMETER Total
        Total number of work items.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Current,
        [Parameter(Mandatory)]
        [double]$Total
    )

    $safeTotal = [Math]::Max($Total, 1)
    return [Math]::Round(($Current / $safeTotal) * 100, 2)
}

function Show-Table {
    <#
    .SYNOPSIS
        Outputs a table of rows.
    .DESCRIPTION
        Outputs a table of rows with a title.
    .PARAMETER Rows
        Table rows.
    .PARAMETER AsTable
        Output as a table.
    #>
    [CmdletBinding()]
    param(
        [array]$Rows,
        [switch]$AsTable
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-NCMessage "(none)" -Level INFO
        return
    }

    if ($AsTable) {
        $Rows | Format-Table -AutoSize
    }
    else {
        $Rows | Format-List
    }
}

function Test-Folder {
    <#
    .SYNOPSIS
        Normalizes and validates a folder path.
    .DESCRIPTION
        Returns the current directory when input is blank, trims trailing separators,
        resolves relative paths against the current location, and throws if the path is invalid.
    .PARAMETER Path
        Folder path to validate (optional).
    #>
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return (Get-Location).ProviderPath
    }

    $normalized = $Path.TrimEnd('\')

    # Resolve existing paths directly
    if (Test-Path -LiteralPath $normalized) {
        return (Resolve-Path -LiteralPath $normalized).ProviderPath
    }

    # Build full path for non-existing targets (supports relative paths)
    $basePath = if ([IO.Path]::IsPathRooted($normalized)) {
        ''
    }
    else {
        (Get-Location).ProviderPath
    }

    $candidate = if ($basePath) { Join-Path -Path $basePath -ChildPath $normalized } else { $normalized }
    try {
        return [System.IO.Path]::GetFullPath($candidate)
    }
    catch {
        throw "Invalid folder path '$Path'. $($_.Exception.Message)"
    }
}
