#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Quarantine's utilities =====================================================================================================

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
