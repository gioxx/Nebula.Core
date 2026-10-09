#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Quarantine helpers =========================================================================================================

function ConvertTo-QuarantineMessageId {
    <#
    .SYNOPSIS
        Normalizes a quarantine MessageId.
    .DESCRIPTION
        Adds angle brackets to a MessageId when missing, ensuring it can be used with Get-QuarantineMessage.
    .PARAMETER MessageId
        MessageId to normalize.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MessageId
    )

    $normalized = $MessageId.Trim()
    if (-not $normalized.StartsWith('<')) {
        $normalized = "<$normalized"
    }
    if (-not $normalized.EndsWith('>')) {
        $normalized = "$normalized>"
    }
    return $normalized
}

function Get-NCQuarantineMessageAllPages {
    <#
    .SYNOPSIS
        Returns every quarantined message matching a Get-QuarantineMessage query, across all pages.
    .DESCRIPTION
        Get-QuarantineMessage returns one page (100 messages by default). This helper requests pages of
        1000 (the cmdlet maximum) until a partial page comes back. Errors from Get-QuarantineMessage are
        not handled here, so callers keep their own error reporting.
    .PARAMETER Parameters
        Get-QuarantineMessage parameters to apply to every page (e.g. RecipientAddress, SenderAddress,
        StartReceivedDate, EndReceivedDate).
    .PARAMETER PageSize
        Messages per page (1-1000).
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Parameters = @{},
        [ValidateRange(1, 1000)]
        [int]$PageSize = 1000
    )

    $page = 1
    do {
        $pageItems = @(Get-QuarantineMessage @Parameters -PageSize $PageSize -Page $page -ErrorAction Stop)
        $pageItems
        $page++
    } while ($pageItems.Count -eq $PageSize)
}
