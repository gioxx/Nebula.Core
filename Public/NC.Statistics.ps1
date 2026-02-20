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
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
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

            $hasArchive = ($mailbox.ArchiveStatus -eq 'Active') -or ($mailbox.ArchiveGuid -and $mailbox.ArchiveGuid -ne [guid]::Empty)
            $archiveSize = $null
            if ($hasArchive) {
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
                "Archive Name"               = if ($hasArchive) { $mailbox.ArchiveName } else { $null }
                "Archive State"              = if ($hasArchive) { $mailbox.ArchiveState } else { $null }
                "Archive Mailbox Size (GB)"  = $archiveSize
                "Archive Warning Quota (GB)" = if ($hasArchive) {
                    Resolve-MbxQuotaValue -RawValue $mailbox.ArchiveWarningQuota -Round:$Round
                }
                else { $null }
                "Archive Quota (GB)"         = if ($hasArchive) {
                    Resolve-MbxQuotaValue -RawValue $mailbox.ArchiveQuota -Round:$Round
                }
                else { $null }
                AutoExpandingArchiveEnabled  = $mailbox.AutoExpandingArchiveEnabled
            }

            $statsBuffer.Add($record) | Out-Null

            if ($writeToCsv -and (($processedCount % $BatchSize) -eq 0)) {
                if ($csvInitialized) {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter) -Append
                }
                else {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                    $csvInitialized = $true
                }
                Write-NCMessage "Processed $processedCount / $totalMailboxes mailboxes, flushed batch to CSV." -Level VERBOSE
                $statsBuffer.Clear()
            }
        }

        if ($writeToCsv) {
            if ($statsBuffer.Count -gt 0) {
                if ($csvInitialized) {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter) -Append
                }
                else {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
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

function Get-MboxStatistics {
    <#
    .SYNOPSIS
        Returns simplified mailbox statistics.
    .DESCRIPTION
        Ensures an Exchange Online session, retrieves mailbox statistics and returns
        a concise set of key fields (size, quotas, basic usage info, latest message trace,
        and oldest mailbox item metadata).
    .PARAMETER UserPrincipalName
        Optional single mailbox identity. When omitted, returns all mailboxes.
    .PARAMETER IncludeArchive
        When present, includes archive size and archive usage percentage (if available).
    .PARAMETER IncludeMessageActivity
        When present, includes latest message trace info and oldest mailbox item metadata
        (LastReceived, LastSent, OldestItemReceivedDate, OldestItemFolderPath).
    .PARAMETER Round
        Round quota values up to the nearest integer GB (default: $true).
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'Identity')]
        [string]$UserPrincipalName,
        [switch]$IncludeArchive,
        [switch]$IncludeMessageActivity,
        [bool]$Round = $true
    )

    Set-ProgressAndInfoPreferences
    try {
        if (-not (Test-EOLConnection)) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
            return
        }

        $mailboxes = @()
        try {
            if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
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

        $processedCount = 0
        $totalMailboxes = $mailboxes.Count

        foreach ($mailbox in $mailboxes) {
            $processedCount++
            $percentComplete = (($processedCount / $totalMailboxes) * 100)
            Write-Progress -Activity "Processing $($mailbox.DisplayName)" -Status "$processedCount of $totalMailboxes ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            $stats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName
            if (-not $stats) {
                continue
            }

            $mailboxSizeGb = Convert-MbxSizeToGB -SizeObject $stats.TotalItemSize
            $prohibitSendQuota = Resolve-MbxQuotaValue -RawValue $mailbox.ProhibitSendQuota -Round:$Round
            $warningQuota = Resolve-MbxQuotaValue -RawValue $mailbox.IssueWarningQuota -Round:$Round
            $oldestItemReceivedDate = $null
            $oldestItemFolderPath = $null
            $lastTrace = $null
            if ($IncludeMessageActivity) {
                $lastTrace = Get-MboxLastMessageTrace -SourceMailbox $mailbox.UserPrincipalName

                try {
                    $oldestItem = Get-MailboxFolderStatistics -Identity $mailbox.UserPrincipalName -IncludeOldestAndNewestItems -ErrorAction Stop |
                        Where-Object { $_.OldestItemReceivedDate -ne $null } |
                        Sort-Object -Property OldestItemReceivedDate |
                        Select-Object -First 1

                    if ($oldestItem) {
                        $oldestItemReceivedDate = $oldestItem.OldestItemReceivedDate
                        $oldestItemFolderPath = $oldestItem.FolderPath
                    }
                }
                catch {
                    Write-NCMessage ("Unable to retrieve oldest mailbox item details for '{0}'. {1}" -f $mailbox.PrimarySmtpAddress, $_.Exception.Message) -Level WARNING
                }
            }

            $percentUsed = $null
            if ($prohibitSendQuota -is [double] -and $prohibitSendQuota -gt 0) {
                $percentUsed = [Math]::Round(($mailboxSizeGb / $prohibitSendQuota) * 100, 2)
            }

            $archiveSize = $null
            $archivePercentUsed = $null
            $hasArchive = ($mailbox.ArchiveStatus -eq 'Active') -or ($mailbox.ArchiveGuid -and $mailbox.ArchiveGuid -ne [guid]::Empty)
            if ($IncludeArchive -and $hasArchive) {
                $archiveStats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName -Archive
                if ($archiveStats) {
                    $archiveSize = Convert-MbxSizeToGB -SizeObject $archiveStats.TotalItemSize
                    $archiveQuota = Resolve-MbxQuotaValue -RawValue $mailbox.ArchiveQuota -Round:$Round
                    if ($archiveQuota -is [double] -and $archiveQuota -gt 0) {
                        $archivePercentUsed = [Math]::Round(($archiveSize / $archiveQuota) * 100, 2)
                    }
                }
            }

            $mailboxTypeDetail = if ($stats.PSObject.Properties.Match('MailboxTypeDetail').Count -gt 0) {
                $stats.MailboxTypeDetail
            }
            elseif ($stats.PSObject.Properties.Match('RecipientTypeDetails').Count -gt 0) {
                $stats.RecipientTypeDetails
            }
            else {
                $mailbox.RecipientTypeDetails
            }

            $mailboxCreated = $null
            if ($mailbox.PSObject.Properties.Match('WhenCreatedUTC').Count -gt 0 -and $mailbox.WhenCreatedUTC) {
                $mailboxCreated = $mailbox.WhenCreatedUTC
            }
            elseif ($mailbox.PSObject.Properties.Match('WhenCreated').Count -gt 0 -and $mailbox.WhenCreated) {
                $mailboxCreated = $mailbox.WhenCreated
            }
            elseif ($stats.PSObject.Properties.Match('WhenMailboxCreated').Count -gt 0 -and $stats.WhenMailboxCreated) {
                $mailboxCreated = $stats.WhenMailboxCreated
            }
            elseif ($stats.PSObject.Properties.Match('DateCreated').Count -gt 0 -and $stats.DateCreated) {
                $mailboxCreated = $stats.DateCreated
            }
            elseif ($stats.PSObject.Properties.Match('Created').Count -gt 0 -and $stats.Created) {
                $mailboxCreated = $stats.Created
            }

            $record = [ordered]@{
                DisplayName          = $mailbox.DisplayName
                UserPrincipalName    = $mailbox.UserPrincipalName
                PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                MailboxTypeDetail    = $mailboxTypeDetail
                MailboxSizeGB        = $mailboxSizeGb
                ItemCount            = $stats.ItemCount
                MailboxCreated       = $mailboxCreated
                LastLogonTime        = $stats.LastLogonTime
                WarningQuotaGB       = $warningQuota
                ProhibitSendQuotaGB  = $prohibitSendQuota
                PercentUsed          = $percentUsed
            }

            if ($IncludeMessageActivity) {
                $record.LastReceived = if ($lastTrace) { $lastTrace.LastReceived } else { $null }
                $record.LastSent = if ($lastTrace) { $lastTrace.LastSent } else { $null }
                $record.OldestItemReceivedDate = $oldestItemReceivedDate
                $record.OldestItemFolderPath = $oldestItemFolderPath
            }

            if ($IncludeArchive) {
                $record.ArchiveSizeGB = $archiveSize
                $record.ArchivePercentUsed = $archivePercentUsed
            }

            [pscustomobject]$record
        }

        Write-Progress -Activity "Export complete" -Completed
    }
    finally {
        Restore-ProgressAndInfoPreferences
    }
}
