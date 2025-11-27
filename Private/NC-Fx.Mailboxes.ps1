#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Mailboxes's utilities ======================================================================================================

function Resolve-MbxQuotaValue {
    <#
    .SYNOPSIS
        Converts Exchange quota strings to numeric GB values.
    .DESCRIPTION
        Removes the trailing " GB (...)" portion returned by Exchange cmdlets and
        optionally rounds up the resulting number.
    .PARAMETER RawValue
        Original quota string (e.g. "100 GB (107,374,182,400 bytes)").
    .PARAMETER Round
        When present, round the numeric value up to the next integer.
    #>
    [CmdletBinding()]
    param(
        [string]$RawValue,
        [switch]$Round
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $numericPart = ($RawValue -replace ' GB.*').Trim()
    if (-not $numericPart) {
        return $null
    }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $parsed = 0.0
    if (-not [double]::TryParse($numericPart, [System.Globalization.NumberStyles]::Float, $culture, [ref]$parsed)) {
        return $numericPart
    }

    if ($Round) {
        return [Math]::Ceiling($parsed)
    }

    return $parsed
}

function Convert-MbxSizeToGB {
    <#
    .SYNOPSIS
        Converts mailbox size objects (UnlimitedByteQuantifiedSize) to GB.
    .DESCRIPTION
        Attempts to parse the byte count from TotalItemSize.ToString() output and rounds to two decimals.
    .PARAMETER SizeObject
        The TotalItemSize value returned by Get-MailboxStatistics.
    #>
    [CmdletBinding()]
    param([object]$SizeObject)

    if ($null -eq $SizeObject) {
        return 0
    }

    $text = $SizeObject.ToString()
    $match = [regex]::Match($text, '\(([0-9,\.]+)\s')
    if (-not $match.Success) {
        return 0
    }

    $bytesString = $match.Groups[1].Value.Replace(',', '')
    $bytes = 0.0
    if (-not [double]::TryParse($bytesString, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$bytes)) {
        return 0
    }

    return [Math]::Round(($bytes / 1GB), 2)
}

function Get-MailboxStatisticsSafe {
    <#
    .SYNOPSIS
        Wraps Get-MailboxStatistics with retry logic and friendly messages.
    .PARAMETER Identity
        Mailbox identity.
    .PARAMETER Archive
        When present, retrieves archive statistics.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,
        [switch]$Archive
    )

    $operation = if ($Archive) {
        "retrieve archive statistics for $Identity"
    }
    else {
        "retrieve mailbox statistics for $Identity"
    }

    try {
        return Invoke-NCRetry -Action {
            if ($Archive) {
                Get-MailboxStatistics -Identity $Identity -Archive -ErrorAction Stop
            }
            else {
                Get-MailboxStatistics -Identity $Identity -ErrorAction Stop
            }
        } -MaxAttempts 3 -DelaySeconds 5 -OperationDescription $operation -OnError {
            param($attempt, $max, $err)
            Write-NCMessage "Unable to $operation (attempt $attempt of $max)." -Level ERROR
        }
    }
    catch {
        Write-NCMessage "Failed to $operation after multiple attempts. $($_.Exception.Message)" -Level ERROR
        return $null
    }
}
