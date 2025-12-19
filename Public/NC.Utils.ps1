#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Utilities ============================================================================================================================

function Format-MessageIDsFromClipboard {
    <#
    .SYNOPSIS
        Formats quarantine identities from the clipboard.
    .DESCRIPTION
        Reads quarantine identities (one per line) from the clipboard, deduplicates them, copies the
        formatted list back to the clipboard, and optionally releases the messages immediately.
    .PARAMETER NoRelease
        Skip the automatic release of the identity entries.
    .PARAMETER PassThru
        Emit the formatted, comma-separated string to the pipeline in addition to copying it to
        the clipboard.
    .EXAMPLE
        Format-MessageIDsFromClipboard -PassThru
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [switch]$NoRelease,
        [switch]$PassThru
    )

    $clipboard = Get-Clipboard
    if ([string]::IsNullOrWhiteSpace($clipboard)) {
        Write-NCMessage "Clipboard is empty. Copy quarantine identities first." -Level WARNING
        return
    }

    $ids = $clipboard -split "`r?`n" | ForEach-Object { $_.Trim().Trim('"').Trim("'") } | Where-Object { $_ }
    if (-not $ids -or $ids.Count -eq 0) {
        Write-NCMessage "No quarantine identities found in clipboard." -Level WARNING
        return
    }

    $normalized = $ids | Select-Object -Unique
    $quoted = $normalized | ForEach-Object { "`"$($_)`"" }
    $output = $quoted -join ", "

    $output | Set-Clipboard
    Write-NCMessage ("Copied {0} quarantine identity value(s) to clipboard:`n{1}" -f $normalized.Count, ($normalized -join "`n")) -Level INFO
    Add-EmptyLine
    Write-NCMessage "Please wait for messages to be released." -Level INFO

    if (-not $NoRelease.IsPresent -and $PSCmdlet.ShouldProcess("quarantine", "Release messages by Identity")) {
        Unlock-QuarantineMessageId -Identity $normalized
    }

    if ($PassThru.IsPresent) { $output }
}

Set-Alias -Name mids -Value Format-MessageIDsFromClipboard -Description "Format MessageId values from clipboard"

function Format-SortedEmailsFromClipboard {
    <#
    .SYNOPSIS
        Extracts, deduplicates, and sorts e-mail addresses from clipboard text.
    .DESCRIPTION
        Parses e-mail addresses from the clipboard, removes duplicates, sorts them, and copies
        a quoted, comma-separated list back to the clipboard.
    .PARAMETER PassThru
        Emit the formatted string to the pipeline in addition to copying it to the clipboard.
    .EXAMPLE
        Format-SortedEmailsFromClipboard -PassThru
    #>
    [CmdletBinding()]
    param(
        [switch]$PassThru
    )

    $clipboard = Get-Clipboard
    if ([string]::IsNullOrWhiteSpace($clipboard)) {
        Write-NCMessage "Clipboard is empty. Copy text containing e-mail addresses first." -Level WARNING
        return
    }

    $emailPattern = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    $emailList = [regex]::Matches($clipboard, $emailPattern) | ForEach-Object { $_.Value }

    if (-not $emailList -or $emailList.Count -eq 0) {
        Write-NCMessage "No e-mail addresses found in clipboard content." -Level WARNING
        return
    }

    $uniqueSortedEmails = $emailList | Sort-Object -Unique
    $quoted = $uniqueSortedEmails | ForEach-Object { "`"$($_)`"" }
    $output = $quoted -join ", "

    $output | Set-Clipboard
    Write-NCMessage ("Copied {0} unique e-mail address(es) to clipboard." -f $uniqueSortedEmails.Count) -Level INFO

    if ($PassThru.IsPresent) { $output }
}

Set-Alias -Name fse -Value Format-SortedEmailsFromClipboard -Description "Format sorted e-mails from clipboard"
