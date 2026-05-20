#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Statistics helpers ===================================================================================================================

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
    .PARAMETER Resume
        Resume from the most recent existing mailbox statistics CSV in the target folder.
    .PARAMETER CsvPath
        Optional CSV file to resume. When omitted, the most recent matching CSV in the target folder is used.
    .PARAMETER MaxConsecutiveErrors
        Stop the export after this many mailbox-level failures in a row.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'Identity')]
        [string]$UserPrincipalName,
        [string]$CsvFolder,
        [string]$CsvPath,
        [switch]$Round,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25,
        [switch]$Resume,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5
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
            Write-NCMessage "No mailboxes matched the provided criteria." -Level WARNING
            return
        }

        $folder = if ($CsvFolder) { Test-Folder $CsvFolder } else { Test-Folder $null }
        $statsBuffer = New-Object System.Collections.Generic.List[object]
        $processedCount = 0
        $writeToCsv = $exportAll
        $csvPath = $null
        $csvInitialized = $false
        $failedInARow = 0
        $aborted = $false
        $normalizeIdentity = {
            param([object]$Value)

            if ($null -eq $Value) {
                return $null
            }

            $text = [string]$Value
            if ([string]::IsNullOrWhiteSpace($text)) {
                return $null
            }

            return $text.Trim().ToLowerInvariant()
        }
        $processedIdentities = [System.Collections.Generic.HashSet[string]]::new()
        $pendingMailboxes = [System.Collections.Generic.List[object]]::new()

        if ($writeToCsv) {
            $defaultCsvPath = New-File("$($folder)\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-MailboxStatistics.csv")
            if ($Resume) {
                $resumePath = $null
                if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                    $resumePath = $CsvPath
                }
                else {
                    $existingCsv = Get-ChildItem -LiteralPath $folder -File -Filter "*_M365-MailboxStatistics.csv" |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1

                    if ($existingCsv) {
                        $resumePath = $existingCsv.FullName
                    }
                }

                if ($resumePath) {
                    $csvPath = $resumePath
                    if (Test-Path -LiteralPath $csvPath) {
                        $csvInitialized = ((Get-Item -LiteralPath $csvPath).Length -gt 0)

                        try {
                            foreach ($row in (Import-CSV -LiteralPath $csvPath -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                                $identity = & $normalizeIdentity $row.UserPrincipalName
                                if (-not $identity) {
                                    $identity = & $normalizeIdentity $row.PrimarySmtpAddress
                                }
                                if (-not $identity) {
                                    $identity = & $normalizeIdentity $row.UserName
                                }

                                if ($identity) {
                                    $null = $processedIdentities.Add($identity)
                                }
                            }
                        }
                        catch {
                            Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csvPath, $_.Exception.Message) -Level WARNING
                            $processedIdentities.Clear()
                            $csvPath = $defaultCsvPath
                            $csvInitialized = $false
                        }

                        if ($csvPath -ne $defaultCsvPath) {
                            Write-NCMessage ("Resuming mailbox statistics from {0}; {1} mailbox(es) already recorded." -f $csvPath, $processedIdentities.Count) -Level INFO
                        }
                    }
                    else {
                        Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csvPath) -Level INFO
                    }
                }
                else {
                    $csvPath = $defaultCsvPath
                    Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $csvPath) -Level INFO
                }
            }
            else {
                $csvPath = $defaultCsvPath
            }

            Write-NCMessage ("Mailbox statistics export will flush every {0} mailbox(es). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
            Write-NCMessage "Saving report to $csvPath" -Level DEBUG

            foreach ($mailbox in $mailboxes) {
                $identity = & $normalizeIdentity $mailbox.UserPrincipalName
                if (-not $identity) {
                    $identity = & $normalizeIdentity $mailbox.PrimarySmtpAddress
                }

                if ($Resume -and $identity -and $processedIdentities.Contains($identity)) {
                    continue
                }

                $null = $pendingMailboxes.Add($mailbox)
            }
        }
        else {
            foreach ($mailbox in $mailboxes) {
                $null = $pendingMailboxes.Add($mailbox)
            }
        }

        $totalMailboxes = $pendingMailboxes.Count

        if ($writeToCsv -and $Resume -and $csvPath -and $processedIdentities.Count -gt 0 -and $totalMailboxes -eq 0) {
            Write-NCMessage "All matching mailboxes are already present in the CSV. Nothing to do." -Level WARNING
            return
        }

        foreach ($mailbox in $pendingMailboxes) {
            $processedCount++
            $Percentage = Get-NCProgressPercent -Current $processedCount -Total $totalMailboxes
            Write-Progress -Activity "Processing $($mailbox.DisplayName)" -Status "$processedCount of $totalMailboxes - $Percentage%" -PercentComplete $Percentage

            $mailboxHadError = $false
            $stats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName
            if ($null -eq $stats) {
                $mailboxHadError = $true
            }
            $mailboxSizeGb = if ($stats) { Convert-MbxSizeToGB -SizeObject $stats.TotalItemSize } else { "Error" }

            $hasArchive = ($mailbox.ArchiveStatus -eq 'Active') -or ($mailbox.ArchiveGuid -and $mailbox.ArchiveGuid -ne [guid]::Empty)
            $archiveSize = $null
            if ($hasArchive) {
                $archiveStats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName -Archive
                if ($null -eq $archiveStats) {
                    $mailboxHadError = $true
                }
                $archiveSize = if ($archiveStats) { Convert-MbxSizeToGB -SizeObject $archiveStats.TotalItemSize } else { "Error" }
            }

            $record = [pscustomobject][ordered]@{
                UserPrincipalName           = $mailbox.UserPrincipalName
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

            if ($mailboxHadError) {
                $failedInARow++
            }
            else {
                $failedInARow = 0
            }

            if ($MaxConsecutiveErrors -gt 0 -and $failedInARow -ge $MaxConsecutiveErrors) {
                if ($writeToCsv -and $statsBuffer.Count -gt 0) {
                    if ($csvInitialized) {
                        $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter) -Append
                    }
                    else {
                        $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                        $csvInitialized = $true
                    }
                    $statsBuffer.Clear()
                }

                Write-NCMessage ("Stopping export after {0} consecutive mailbox error(s). Partial report kept at {1}." -f $failedInARow, $csvPath) -Level ERROR
                $aborted = $true
                break
            }

            if ($writeToCsv -and (($processedCount % $BatchSize) -eq 0)) {
                if ($csvInitialized) {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter) -Append
                }
                else {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                    $csvInitialized = $true
                }
                Write-Verbose "Processed $processedCount / $totalMailboxes mailboxes, flushed batch to CSV."
                $statsBuffer.Clear()
            }
        }

        if ($writeToCsv -and -not $aborted) {
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
        elseif ($aborted) {
            if ($statsBuffer.Count -gt 0) {
                if ($csvInitialized) {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter) -Append
                }
                else {
                    $statsBuffer | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                }
            }
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

function Export-MboxDeletedItemSize {
    <#
    .SYNOPSIS
        Exports mailbox deleted item store usage.
    .DESCRIPTION
        Ensures an Exchange Online session, retrieves all user mailboxes or a selected subset,
        calculates the deleted item size for each mailbox, and exports the report to CSV by default.
    .PARAMETER UserPrincipalName
        Optional mailbox identity or identities. Accepts pipeline input.
    .PARAMETER CsvFolder
        Destination folder for the CSV file when exporting the report.
    .PARAMETER Csv
        When present, export the report to CSV. Defaults to on.
    .PARAMETER BatchSize
        Number of processed mailboxes before flushing partial CSV output.
    .PARAMETER Resume
        Resume from the latest matching CSV in the target folder or from -CsvPath.
    .PARAMETER CsvPath
        Explicit CSV file to resume. When omitted, the most recent matching CSV in the target folder is used.
    .PARAMETER MaxConsecutiveErrors
        Stop after this many consecutive mailbox-level failures.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'Identity', 'Mailbox', 'SourceMailbox')]
        [string[]]$UserPrincipalName,
        [string]$CsvFolder,
        [bool]$Csv = $true,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25,
        [switch]$Resume,
        [string]$CsvPath,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5
    )

    begin {
        Set-ProgressAndInfoPreferences
        $requestedMailboxes = [System.Collections.Generic.List[string]]::new()
        $report = [System.Collections.Generic.List[object]]::new()
        $processedSinceFlush = 0
        $consecutiveErrors = 0
        $aborted = $false
    }

    process {
        foreach ($entry in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $requestedMailboxes.Add($entry.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $mailboxes = @()
            if ($requestedMailboxes.Count -gt 0) {
                foreach ($mailboxId in ($requestedMailboxes | Select-Object -Unique)) {
                    try {
                        $mailboxes += @(Get-Mailbox -Identity $mailboxId -ErrorAction Stop)
                    }
                    catch {
                        Write-NCMessage "Mailbox '$mailboxId' not found. $($_.Exception.Message)" -Level WARNING
                    }
                }
            }
            else {
                $mailboxes = @(Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue |
                    Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' })
            }

            if (-not $mailboxes -or $mailboxes.Count -eq 0) {
                Write-NCMessage "No user mailboxes matched the provided criteria." -Level WARNING
                return
            }

            $folder = if ($CsvFolder) { Test-Folder $CsvFolder } else { Test-Folder $null }
            $defaultCsvPath = New-File "$folder\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-DeletedItemSize.csv"
            $csv = $defaultCsvPath
            $processedMailboxes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            if ($Resume) {
                $resumePath = $null
                if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                    $resumePath = $CsvPath
                }
                else {
                    $existingCsv = Get-ChildItem -LiteralPath $folder -File -Filter "*_M365-DeletedItemSize.csv" |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
                    if ($existingCsv) {
                        $resumePath = $existingCsv.FullName
                    }
                }

                if ($resumePath) {
                    $csv = $resumePath
                    if (Test-Path -LiteralPath $csv) {
                        try {
                            foreach ($row in (Import-CSV -LiteralPath $csv -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                                if ($row.UserPrincipalName) {
                                    $null = $processedMailboxes.Add(([string]$row.UserPrincipalName).Trim())
                                }
                            }
                            Write-NCMessage ("Resuming deleted item export from {0}; {1} mailbox(es) already recorded." -f $csv, $processedMailboxes.Count) -Level INFO
                        }
                        catch {
                            Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csv, $_.Exception.Message) -Level WARNING
                            $processedMailboxes.Clear()
                            $csv = $defaultCsvPath
                        }
                    }
                    else {
                        Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csv) -Level INFO
                    }
                }
                else {
                    Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $csv) -Level INFO
                }
            }

            Write-NCMessage ("Deleted item export will flush every {0} mailbox(es). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
            Write-NCMessage "Saving report to $csv" -Level DEBUG

            $totalMailboxes = $mailboxes.Count
            $processedCount = 0

            foreach ($mailbox in $mailboxes) {
                $processedCount++
                $Percentage = Get-NCProgressPercent -Current $processedCount -Total $totalMailboxes
                Write-Progress -Activity "Processing $($mailbox.DisplayName)" -Status "$processedCount of $totalMailboxes - $Percentage%" -PercentComplete $Percentage

                if ($Resume -and $mailbox.UserPrincipalName -and $processedMailboxes.Contains($mailbox.UserPrincipalName)) {
                    Write-Verbose "Skipping $($mailbox.UserPrincipalName), already processed."
                    continue
                }

                $stats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName
                if (-not $stats) {
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        if ($report.Count -gt 0) {
                            if ((Test-Path -LiteralPath $csv) -and ((Get-Item -LiteralPath $csv).Length -gt 0)) {
                                $report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                            }
                            else {
                                $report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                            }
                            $report.Clear()
                        }

                        Write-NCMessage ("Stopping export after {0} consecutive mailbox error(s). Partial report kept at {1}." -f $consecutiveErrors, $csv) -Level ERROR
                        $aborted = $true
                        break
                    }
                    continue
                }

                $report.Add([pscustomobject][ordered]@{
                        UserPrincipalName      = $mailbox.UserPrincipalName
                        DisplayName            = $mailbox.DisplayName
                        PrimarySmtpAddress     = $mailbox.PrimarySmtpAddress
                        TotalDeletedItemSizeGB = Convert-MbxSizeToGB -SizeObject $stats.TotalDeletedItemSize
                    }) | Out-Null

                $null = $processedMailboxes.Add($mailbox.UserPrincipalName)
                $consecutiveErrors = 0
                $processedSinceFlush++

                if ($processedSinceFlush -ge $BatchSize -and $report.Count -gt 0) {
                    if ((Test-Path -LiteralPath $csv) -and ((Get-Item -LiteralPath $csv).Length -gt 0)) {
                        $report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                    }
                    else {
                        $report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                    }
                    Write-Verbose "Processed $processedCount / $totalMailboxes mailboxes, flushed batch to CSV."
                    $report.Clear()
                    $processedSinceFlush = 0
                }
            }

            if ($Csv -and $report.Count -gt 0) {
                if ((Test-Path -LiteralPath $csv) -and ((Get-Item -LiteralPath $csv).Length -gt 0)) {
                    $report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                }
                else {
                    $report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                }
            }

            if ($Csv) {
                if ($aborted) {
                    Write-NCMessage "Deleted item size report export stopped early. Partial data kept at $csv." -Level ERROR
                }
                else {
                    Write-NCMessage "Deleted item size report exported to $csv." -Level SUCCESS
                }
            }
            else {
                $report
            }
        }
        finally {
            Write-Progress -Activity "Processing deleted item size" -Completed
            Restore-ProgressAndInfoPreferences
        }
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
        When present, includes archive size, archive quota, and archive usage percentage (if available).
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

    begin {
        Set-ProgressAndInfoPreferences
        $pipelineUpns = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            [void]$pipelineUpns.Add($UserPrincipalName)
        }
    }

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $mailboxes = @()
            try {
                if ($pipelineUpns.Count -eq 0) {
                    $mailboxes = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue
                }
                else {
                    foreach ($upn in ($pipelineUpns | Select-Object -Unique)) {
                        try {
                            $mailboxes += @(Get-Mailbox -Identity $upn -ErrorAction Stop)
                        }
                        catch {
                            Write-NCMessage "Mailbox not found for '$upn'. Skipping. $($_.Exception.Message)" -Level WARNING
                        }
                    }
                }
            }
            catch {
                Write-NCMessage "Failed to retrieve mailbox information: $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $mailboxes -or $mailboxes.Count -eq 0) {
                Write-NCMessage "No mailboxes matched the provided criteria." -Level WARNING
                return
            }

            $processedCount = 0
            $totalMailboxes = $mailboxes.Count

            foreach ($mailbox in $mailboxes) {
                $processedCount++
                $Percentage = Get-NCProgressPercent -Current $processedCount -Total $totalMailboxes
                Write-Progress -Activity "Processing $($mailbox.DisplayName)" -Status "$processedCount of $totalMailboxes - $Percentage%" -PercentComplete $Percentage

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
                            Where-Object { $null -ne $_.OldestItemReceivedDate } |
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
                $archiveQuota = $null
                $hasArchive = ($mailbox.ArchiveStatus -eq 'Active') -or ($mailbox.ArchiveGuid -and $mailbox.ArchiveGuid -ne [guid]::Empty)
                if ($IncludeArchive -and $hasArchive) {
                    $archiveQuota = Resolve-MbxQuotaValue -RawValue $mailbox.ArchiveQuota -Round:$Round
                    $archiveStats = Get-MailboxStatisticsSafe -Identity $mailbox.UserPrincipalName -Archive
                    if ($archiveStats) {
                        $archiveSize = Convert-MbxSizeToGB -SizeObject $archiveStats.TotalItemSize
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
                    ArchiveEnabled       = [bool]$hasArchive
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
                    $record.ArchiveQuotaGB = $archiveQuota
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
}


