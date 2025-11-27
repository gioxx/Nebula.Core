#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Statistics ===========================================================================================================================

function Export-MboxStatistics {
    <#
    .SYNOPSIS
        Exports mailbox (and archive) size/quota statistics.
    .DESCRIPTION
        Ensures an Exchange Online session, retrieves either all mailboxes or a single identity,
        calculates usage/quota information (optionally rounding quotas), and writes to CSV or
        returns objects to the pipeline.
    .PARAMETER UserPrincipalName
        Optional single mailbox identity. When omitted, exports all mailboxes to CSV.
    .PARAMETER CsvFolder
        Destination folder for the CSV file (defaults to current directory when exporting all mailboxes).
    .PARAMETER Round
        Round quota values up to the nearest integer GB.
    .PARAMETER BatchSize
        Number of processed mailboxes before flushing partial CSV output (defaults to 25).
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'Identity')]
        [string]$UserPrincipalName,
        [string]$CsvFolder,
        [switch]$Round,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25
    )

    Set-ProgressAndInfoPreferences
    try {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        $exportAll = [string]::IsNullOrWhiteSpace($UserPrincipalName)
        $mailboxes = @()

        try {
            if ($exportAll) {
                $mailboxes = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue
            }
            else {
                $mailboxes = @(Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop)
            }
        }
        catch {
            Write-NCMessage "Failed to retrieve mailbox information: $($_.Exception.Message)" -Level ERROR
            return
        }

        if (-not $mailboxes -or $mailboxes.Count -eq 0) {
            Write-NCMessage "No mailboxes found matching the provided criteria." -Level WARNING
            return
        }

        $folder = if ($CsvFolder) { Test-Folder $CsvFolder } else { Test-Folder $null }
        $statsBuffer = New-Object System.Collections.Generic.List[object]
        $processedCount = 0
        $totalMailboxes = $mailboxes.Count
        $writeToCsv = $exportAll
        $csvPath = $null
        $csvInitialized = $false

        if ($writeToCsv) {
            $csvPath = New-File("$($folder)\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-MailboxStatistics.csv")
            Write-NCMessage "Saving report to $csvPath" -Level DEBUG
        }

        foreach ($mailbox in $mailboxes) {
            $processedCount++
            $percentComplete = (($processedCount / $totalMailboxes) * 100)
            Write-Progress -Activity "Processing $($mailbox.DisplayName)" -Status "$processedCount of $totalMailboxes ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            $stats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName
            $mailboxSizeGb = if ($stats) { Convert-MbxSizeToGB -SizeObject $stats.TotalItemSize } else { "Error" }

            $archiveSize = $null
            if ($mailbox.ArchiveDatabase) {
                $archiveStats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName -Archive
                $archiveSize = if ($archiveStats) { Convert-MbxSizeToGB -SizeObject $archiveStats.TotalItemSize } else { "Error" }
            }

            $record = [pscustomobject][ordered]@{
                UserName                     = $mailbox.DisplayName
                ServerName                   = $mailbox.ServerName
                Database                     = $mailbox.Database
                RecipientTypeDetails         = $mailbox.RecipientTypeDetails
                PrimarySmtpAddress           = $mailbox.PrimarySmtpAddress
                "Mailbox Size (GB)"          = $mailboxSizeGb
                "Issue Warning Quota (GB)"   = Resolve-MbxQuotaValue -RawValue $mailbox.IssueWarningQuota -Round:$Round
                "Prohibit Send Quota (GB)"   = Resolve-MbxQuotaValue -RawValue $mailbox.ProhibitSendQuota -Round:$Round
                "Archive Database"           = if ($mailbox.ArchiveDatabase) { $mailbox.ArchiveDatabase } else { $null }
                "Archive Name"               = if ($mailbox.ArchiveDatabase) { $mailbox.ArchiveName } else { $null }
                "Archive State"              = if ($mailbox.ArchiveDatabase) { $mailbox.ArchiveState } else { $null }
                "Archive Mailbox Size (GB)"  = $archiveSize
                "Archive Warning Quota (GB)" = if ($mailbox.ArchiveDatabase) {
                    Resolve-MbxQuotaValue -RawValue $mailbox.ArchiveWarningQuota -Round:$Round
                }
                else { $null }
                "Archive Quota (GB)"         = if ($mailbox.ArchiveDatabase) {
                    Resolve-MbxQuotaValue -RawValue $mailbox.ArchiveQuota -Round:$Round
                }
                else { $null }
                AutoExpandingArchiveEnabled  = $mailbox.AutoExpandingArchiveEnabled
            }

            $statsBuffer.Add($record) | Out-Null

            if ($writeToCsv -and (($processedCount % $BatchSize) -eq 0)) {
                if ($csvInitialized) {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -DefaultLimiter $($NCVars.CSV_DefaultLimiter) -Append
                }
                else {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -DefaultLimiter $($NCVars.CSV_DefaultLimiter)
                    $csvInitialized = $true
                }
                Write-NCMessage "Processed $processedCount / $totalMailboxes mailboxes, flushed batch to CSV." -Level VERBOSE
                $statsBuffer.Clear()
            }
        }

        if ($writeToCsv) {
            if ($statsBuffer.Count -gt 0) {
                if ($csvInitialized) {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -DefaultLimiter $($NCVars.CSV_DefaultLimiter) -Append
                }
                else {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -DefaultLimiter $($NCVars.CSV_DefaultLimiter)
                }
            }

            Write-NCMessage "Mailbox statistics exported to $csvPath." -Level SUCCESS
        }
        else {
            $statsBuffer
        }

        Write-Progress -Activity "Export complete" -Completed
    }
    finally {
        Restore-ProgressAndInfoPreferences
    }
}
