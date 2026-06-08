#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Calendar helpers =====================================================================================================================

function Copy-OoOMessage {
    <#
    .SYNOPSIS
        Copies out-of-office settings from one mailbox to another.
    .DESCRIPTION
        Validates Exchange Online connectivity, reads the source mailbox auto-reply configuration,
        and applies the same messages to the destination mailbox. Optionally forces the destination
        to be enabled immediately instead of preserving the source state/schedule.
    .PARAMETER SourceMailbox
        Mailbox from which to read out-of-office configuration. Accepts pipeline input.
    .PARAMETER DestinationMailbox
        Mailbox to update with the cloned configuration.
    .PARAMETER ForceEnable
        Enable auto-replies immediately on the destination, ignoring the source AutoReplyState.
    .PARAMETER PassThru
        Emit the updated auto-reply configuration for the destination mailbox.
    .EXAMPLE
        Copy-OoOMessage -SourceMailbox user1@contoso.com -DestinationMailbox user2@contoso.com
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Source', 'Identity')]
        [string]$SourceMailbox,
        [Parameter(Mandatory)]
        [Alias('Destination')]
        [string]$DestinationMailbox,
        [switch]$ForceEnable,
        [switch]$PassThru
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
            return
        }

        try {
            $sourceConfig = Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to read auto-reply configuration from '$SourceMailbox'. $($_.Exception.Message)" -Level ERROR
            return
        }

        if ([string]::IsNullOrWhiteSpace($sourceConfig.InternalMessage) -and [string]::IsNullOrWhiteSpace($sourceConfig.ExternalMessage)) {
            Write-NCMessage "Source mailbox '$SourceMailbox' has no out-of-office message defined. Destination will still be updated." -Level WARNING
        }

        $targetState = if ($ForceEnable.IsPresent) { 'Enabled' } else { $sourceConfig.AutoReplyState }
        $setParams = @{
            Identity         = $DestinationMailbox
            AutoReplyState   = $targetState
            InternalMessage  = $sourceConfig.InternalMessage
            ExternalMessage  = $sourceConfig.ExternalMessage
            ExternalAudience = $sourceConfig.ExternalAudience
        }

        if (-not $ForceEnable.IsPresent -and $sourceConfig.AutoReplyState -eq 'Scheduled' -and $sourceConfig.StartTime -and $sourceConfig.EndTime) {
            $setParams.StartTime = $sourceConfig.StartTime
            $setParams.EndTime = $sourceConfig.EndTime
        }

        $action = "apply out-of-office settings from $SourceMailbox"
        if (-not $PSCmdlet.ShouldProcess($DestinationMailbox, $action)) {
            return
        }

        try {
            Set-MailboxAutoReplyConfiguration @setParams -ErrorAction Stop
            Write-NCMessage ("Copied out-of-office configuration from {0} to {1}." -f $SourceMailbox, $DestinationMailbox) -Level SUCCESS
        }
        catch {
            Write-NCMessage "Unable to update '$DestinationMailbox'. $($_.Exception.Message)" -Level ERROR
            return
        }

        try {
            $updated = Get-MailboxAutoReplyConfiguration -Identity $DestinationMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Destination updated, but verification failed. $($_.Exception.Message)" -Level WARNING
            return
        }

        if ($PassThru.IsPresent) {
            return $updated
        }

        $updated | Select-Object Identity, AutoReplyState, StartTime, EndTime, ExternalAudience
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Export-CalendarPermission {
    <#
    .SYNOPSIS
        Exports calendar permissions for selected mailboxes.
    .DESCRIPTION
        Collects calendar permissions for specific mailboxes, domains, or all mailboxes, and writes
        the results to a CSV report. Returns the CSV path, or the permission objects when -PassThru
        is specified.
    .PARAMETER SourceMailbox
        Mailbox identities to analyze. Accepts pipeline input.
    .PARAMETER SourceDomain
        Domain to analyze (e.g. contoso.com). All matching mailboxes are included.
    .PARAMETER OutputFolder
        Destination folder for the CSV file. Defaults to the current directory.
    .PARAMETER All
        Scan every mailbox in the tenant (excluding GuestMailUser). Implies CSV export.
    .PARAMETER PassThru
        Emit the collected permission objects in addition to writing the CSV report.
    .PARAMETER BatchSize
        Number of processed mailboxes before flushing partial CSV output.
    .PARAMETER Resume
        Resume from the latest matching CSV in the target folder or from -CsvPath.
    .PARAMETER CsvPath
        Explicit CSV file to resume. When omitted, the most recent matching CSV in the target folder is used.
    .PARAMETER MaxConsecutiveErrors
        Stop after this many consecutive mailbox-level failures.
    .EXAMPLE
        Export-CalendarPermission -SourceMailbox user@contoso.com -OutputFolder C:\Temp
    .EXAMPLE
        Export-CalendarPermission -SourceDomain contoso.com -PassThru
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity', 'Mailbox')]
        [string[]]$SourceMailbox,
        [string[]]$SourceDomain,
        [string]$OutputFolder,
        [switch]$All,
        [switch]$PassThru,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25,
        [switch]$Resume,
        [string]$CsvPath,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5
    )

    begin {
        Set-ProgressAndInfoPreferences
        $mailboxInputs = [System.Collections.Generic.List[string]]::new()
        $domainInputs = [System.Collections.Generic.List[string]]::new()
        $results = [System.Collections.Generic.List[object]]::new()
        $processedSinceFlush = 0
        $consecutiveErrors = 0
        $aborted = $false
    }

    process {
        foreach ($entry in $SourceMailbox) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $mailboxInputs.Add($entry) | Out-Null
            }
        }

        foreach ($domain in $SourceDomain) {
            if (-not [string]::IsNullOrWhiteSpace($domain)) {
                $domainInputs.Add($domain) | Out-Null
            }
        }
    }

    end {
        try {
            try {
                $reportFolder = Test-Folder -Path $OutputFolder
            }
            catch {
                Write-NCMessage "Destination folder is not valid. $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $mailboxInputs.Count -and -not $domainInputs.Count -and -not $All.IsPresent) {
                Write-NCMessage "No mailbox or domain specified; scanning all mailboxes. This may take a while." -Level WARNING
                $All = $true
            }

            if (-not (Test-EOLConnection)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $targets = [System.Collections.Generic.List[object]]::new()
            $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            $addMailbox = {
                param($mailbox)
                if (-not $mailbox) { return }
                $key = if ($mailbox.PrimarySmtpAddress) { $mailbox.PrimarySmtpAddress } else { $mailbox.Identity }
                if ([string]::IsNullOrWhiteSpace($key)) { return }
                if ($seen.Add($key)) {
                    $targets.Add($mailbox) | Out-Null
                }
            }

            if ($All.IsPresent) {
                $allMailboxes = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue
                foreach ($mbx in $allMailboxes) { & $addMailbox $mbx }
            }

            foreach ($domain in $domainInputs) {
                Write-NCMessage ("Analyzing mailboxes in {0} ..." -f $domain) -Level INFO
                $domainMailboxes = Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue | Where-Object {
                    $_.RecipientTypeDetails -ne 'GuestMailUser' -and $_.EmailAddresses -like "*@$domain"
                }
                foreach ($mbx in $domainMailboxes) { & $addMailbox $mbx }
            }

            foreach ($identity in $mailboxInputs) {
                try {
                    $resolved = Get-Mailbox -Identity $identity -ErrorAction Stop
                    & $addMailbox $resolved
                }
                catch {
                    Write-NCMessage "Mailbox '$identity' not found. $($_.Exception.Message)" -Level ERROR
                }
            }

            if ($targets.Count -eq 0) {
                Write-NCMessage "No mailboxes found for the specified filters." -Level WARNING
                return
            }

            $normalizeText = {
                param($value)
                return Get-NormalizedText -Value $value
            }
            $buildPermissionKey = {
                param($row)

                $mailboxId = & $normalizeText $row.MailboxIdentity
                if (-not $mailboxId) {
                    $mailboxId = & $normalizeText $row.PrimarySmtpAddress
                }
                $user = & $normalizeText $row.User
                $permissions = & $normalizeText $row.Permissions
                return "{0}|{1}|{2}" -f $mailboxId, $user, $permissions
            }
            $existingPermissionKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $csvPathResolved = $null
            $folderForExport = $reportFolder

            $defaultCsvPath = New-File (Join-Path -Path $folderForExport -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-CalendarPermissions-Report.csv")
            $csvPathResolved = $defaultCsvPath

            if ($Resume) {
                $resumePath = $null
                if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                    $resumePath = $CsvPath
                }
                else {
                    $existingCsv = Get-ChildItem -LiteralPath $folderForExport -File -Filter "*_M365-CalendarPermissions-Report.csv" |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
                    if ($existingCsv) {
                        $resumePath = $existingCsv.FullName
                    }
                }

                if ($resumePath) {
                    $csvPathResolved = $resumePath
                    if (Test-Path -LiteralPath $csvPathResolved) {
                        try {
                            foreach ($row in (Import-CSV -LiteralPath $csvPathResolved -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                                $null = $existingPermissionKeys.Add((& $buildPermissionKey $row))
                            }
                            Write-NCMessage ("Resuming calendar permission export from {0}; {1} row(s) already recorded." -f $csvPathResolved, $existingPermissionKeys.Count) -Level INFO
                        }
                        catch {
                            Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csvPathResolved, $_.Exception.Message) -Level WARNING
                            $existingPermissionKeys.Clear()
                            $csvPathResolved = $defaultCsvPath
                        }
                    }
                    else {
                        Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csvPathResolved) -Level INFO
                    }
                }
                else {
                    Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $csvPathResolved) -Level INFO
                }
            }

            Write-NCMessage ("Calendar permission export will flush every {0} mailbox(es). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
            Write-NCMessage "Saving report to $csvPathResolved" -Level DEBUG

            $results = [System.Collections.Generic.List[object]]::new()
            $counter = 0

            foreach ($mailbox in $targets) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $targets.Count
                Write-Progress -Activity "Processing $($mailbox.PrimarySmtpAddress)" -Status "$counter of $($targets.Count) - $Percentage%" -PercentComplete $Percentage

                try {
                    $exoMailbox = Get-EXOMailbox -Identity $mailbox.Identity -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to load mailbox '$($mailbox.Identity)'. $($_.Exception.Message)" -Level ERROR
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        $aborted = $true
                        break
                    }
                    continue
                }

                try {
                    $calendarFolder = Get-MailboxFolderStatistics $exoMailbox.PrimarySmtpAddress -FolderScope Calendar -ErrorAction Stop | Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
                }
                catch {
                    Write-NCMessage "Unable to read calendar folder for '$($exoMailbox.PrimarySmtpAddress)'. $($_.Exception.Message)" -Level ERROR
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        $aborted = $true
                        break
                    }
                    continue
                }

                if (-not $calendarFolder) {
                    Write-NCMessage "Calendar folder not found for '$($exoMailbox.PrimarySmtpAddress)'." -Level WARNING
                    continue
                }

                try {
                    $folderIdentity = "{0}:{1}" -f $exoMailbox.PrimarySmtpAddress, $calendarFolder.FolderId
                    $folderPerms = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop | Where-Object {
                        $_.AccessRights -notlike 'AvailabilityOnly' -and $_.AccessRights -notlike 'None'
                    }
                }
                catch {
                    Write-NCMessage "Unable to retrieve calendar permissions for '$($exoMailbox.PrimarySmtpAddress)'. $($_.Exception.Message)" -Level ERROR
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        $aborted = $true
                        break
                    }
                    continue
                }

                foreach ($perm in $folderPerms) {
                    $row = [pscustomobject]@{
                        MailboxIdentity     = $exoMailbox.Identity
                        PrimarySmtpAddress  = $exoMailbox.PrimarySmtpAddress
                        User                = $perm.User
                        Permissions         = ($perm.AccessRights -join ', ')
                    }
                    $rowKey = & $buildPermissionKey $row
                    if ($Resume -and $existingPermissionKeys.Contains($rowKey)) {
                        continue
                    }
                    $results.Add($row) | Out-Null
                }

                $consecutiveErrors = 0
                $processedSinceFlush++
                if ($BatchSize -gt 0 -and $results.Count -gt 0 -and ($processedSinceFlush -ge $BatchSize)) {
                    if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                        $results | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                    }
                    else {
                        $results | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                    }
                    $results.Clear()
                    $processedSinceFlush = 0
                }
            }

            if ($aborted -and $results.Count -gt 0) {
                if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    $results | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                }
                else {
                    $results | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                }
                $results.Clear()
            }

            if ($results.Count -eq 0 -and (-not (Test-Path -LiteralPath $csvPathResolved) -or (Get-Item -LiteralPath $csvPathResolved).Length -eq 0)) {
                Write-NCMessage "No calendar permissions found for the selected scope." -Level WARNING
                return
            }

            try {
                if ($results.Count -gt 0) {
                    if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                        $results | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                    }
                    else {
                        $results | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                    }
                }
                if ($aborted) {
                    Write-NCMessage ("Calendar permission export stopped early. Partial data kept at {0}" -f $csvPathResolved) -Level ERROR
                }
                else {
                    Write-NCMessage ("Calendar permission report exported to {0}" -f $csvPathResolved) -Level SUCCESS
                }
            }
            catch {
                Write-NCMessage "Unable to write CSV report. $($_.Exception.Message)" -Level ERROR
                return
            }

            if ($PassThru.IsPresent) {
                $results
            }
        }
        finally {
            Write-Progress -Activity "Processing calendar permissions" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Get-RoomDetails {
    <#
    .SYNOPSIS
        Lists room list members with capacity and location details.
    .DESCRIPTION
        Ensures Exchange Online connectivity, enumerates room lists (optionally filtered by City),
        expands member rooms, and returns/export details. Supports CSV export and grid view.
    .PARAMETER City
        Optional city/name filter applied to room list name or display name.
    .PARAMETER Csv
        Export results to CSV.
    .PARAMETER OutputFolder
        Destination folder for CSV (defaults to current directory).
    .PARAMETER GridView
        Show the results in Out-GridView.
    .PARAMETER PassThru
        Emit the room details objects to the pipeline (also when exporting).
    .EXAMPLE
        Get-RoomDetails -City Milan -Csv -OutputFolder C:\Temp
    .EXAMPLE
        Get-RoomDetails -GridView
    #>
    [CmdletBinding()]
    param(
        [string[]]$City,
        [switch]$Csv,
        [string]$OutputFolder,
        [switch]$GridView,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
    }

    process {}

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $roomGroups = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited -WarningAction SilentlyContinue
            if ($City -and $City.Count -gt 0) {
                $roomGroups = $roomGroups | Where-Object {
                    foreach ($c in $City) {
                        if ($_.Name -like "*$c*" -or $_.DisplayName -like "*$c*") { return $true }
                    }
                    return $false
                }
            }

            if (-not $roomGroups -or $roomGroups.Count -eq 0) {
                Write-NCMessage "No room lists found with the specified filters." -Level WARNING
                return
            }

            $counter = 0
            foreach ($group in $roomGroups) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $roomGroups.Count
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $($roomGroups.Count) - $Percentage%" -PercentComplete $Percentage

                try {
                    $members = Get-DistributionGroupMember -Identity $group.PrimarySmtpAddress -ResultSize Unlimited -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to retrieve members for room list '$($group.PrimarySmtpAddress)'. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                foreach ($member in $members) {
                    try {
                        $mailbox = Get-Mailbox -Identity $member.PrimarySmtpAddress -ErrorAction Stop
                    }
                    catch {
                        Write-NCMessage "Unable to load mailbox '$($member.PrimarySmtpAddress)'. $($_.Exception.Message)" -Level ERROR
                        continue
                    }

                    $results.Add([pscustomobject]@{
                            Location                    = $group.Name
                            LocationPrimarySmtpAddress  = $group.PrimarySmtpAddress
                            RoomDisplayName             = $member.DisplayName
                            RoomPrimarySmtpAddress      = $member.PrimarySmtpAddress
                            RoomCapacity                = $mailbox.ResourceCapacity
                        }) | Out-Null
                }
            }

            Write-Progress -Activity "Processing room lists" -Completed

            if ($results.Count -eq 0) {
                Write-NCMessage "No room details found. Check filters or RoomList definitions." -Level WARNING
                return
            }

            $csvPath = $null
            if ($Csv.IsPresent) {
                try {
                    $folder = Test-Folder -Path $OutputFolder
                    $csvPath = New-File (Join-Path -Path $folder -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-Rooms.csv")
                    $results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                    Write-NCMessage ("Export completed: {0}" -f $csvPath) -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Unable to export CSV. $($_.Exception.Message)" -Level ERROR
                }
            }

            if ($GridView.IsPresent) {
                try {
                    $results | Out-GridView -Title "M365 Rooms Details"
                }
                catch {
                    Write-NCMessage "Unable to show grid view. $($_.Exception.Message)" -Level WARNING
                }
            }

            if ($PassThru.IsPresent -or -not $Csv.IsPresent -or $GridView.IsPresent) {
                $results
            }
            elseif ($csvPath) {
                $csvPath
            }
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Set-OoO {
    <#
    .SYNOPSIS
        Enables, schedules, or disables out-of-office replies for a mailbox.
    .DESCRIPTION
        Ensures Exchange Online connectivity, applies the provided internal/external messages,
        optionally schedules a start/end interval, or disables auto-replies entirely.
    .PARAMETER SourceMailbox
        Mailbox on which to configure out-of-office. Accepts pipeline input.
    .PARAMETER ChooseDayFromCalendar
        Opens a calendar popup to pick start and end dates for scheduled auto-replies.
    .PARAMETER InternalMessage
        HTML/text used for internal recipients. Defaults to the current configuration or a template.
    .PARAMETER ExternalMessage
        HTML/text used for external recipients. Defaults to InternalMessage when omitted.
    .PARAMETER StartTime
        Optional start time for scheduled auto-replies. Requires -EndTime.
    .PARAMETER EndTime
        Optional end time for scheduled auto-replies. Requires -StartTime.
    .PARAMETER ExternalAudience
        Audience for external replies: None, Known, All. Defaults to All.
    .PARAMETER Disable
        Disable auto-replies on the specified mailbox.
    .PARAMETER PassThru
        Emit the updated auto-reply configuration.
    .EXAMPLE
        Set-OoO -SourceMailbox user@contoso.com -InternalMessage "<p>Back on Monday</p>" -ExternalMessage "<p>Back soon</p>"
    .EXAMPLE
        Set-OoO -SourceMailbox user@contoso.com -Disable
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Enable')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox,
        [Parameter(ParameterSetName = 'Enable')]
        [switch]$ChooseDayFromCalendar,
        [Parameter(ParameterSetName = 'Enable')]
        [string]$InternalMessage,
        [Parameter(ParameterSetName = 'Enable')]
        [string]$ExternalMessage,
        [Parameter(ParameterSetName = 'Enable')]
        [Nullable[datetime]]$StartTime,
        [Parameter(ParameterSetName = 'Enable')]
        [Nullable[datetime]]$EndTime,
        [Parameter(ParameterSetName = 'Enable')]
        [ValidateSet('None', 'Known', 'All')]
        [string]$ExternalAudience = 'All',
        [Parameter(ParameterSetName = 'Disable', Mandatory)]
        [switch]$Disable,
        [switch]$PassThru
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
            return
        }

        try {
            $currentConfig = Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to read current auto-reply configuration for '$SourceMailbox'. $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($Disable.IsPresent) {
            if (-not $PSCmdlet.ShouldProcess($SourceMailbox, "Disable auto-replies")) {
                return
            }

            try {
                Set-MailboxAutoReplyConfiguration -Identity $SourceMailbox -AutoReplyState Disabled -ErrorAction Stop
                Write-NCMessage ("Disabled out-of-office for {0}." -f $SourceMailbox) -Level SUCCESS
            }
            catch {
                Write-NCMessage "Unable to disable out-of-office for '$SourceMailbox'. $($_.Exception.Message)" -Level ERROR
                return
            }

            $updatedDisable = Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox -ErrorAction SilentlyContinue
            if ($PassThru.IsPresent -and $updatedDisable) { $updatedDisable }
            return
        }

        if (($StartTime.HasValue -and -not $EndTime.HasValue) -or (-not $StartTime.HasValue -and $EndTime.HasValue)) {
            Write-NCMessage "Both -StartTime and -EndTime must be provided to schedule auto-replies." -Level ERROR
            return
        }

        if ($StartTime.HasValue -and $EndTime.HasValue -and $StartTime.Value -ge $EndTime.Value) {
            Write-NCMessage "StartTime must be earlier than EndTime." -Level ERROR
            return
        }

        if ($ChooseDayFromCalendar.IsPresent -and ($StartTime.HasValue -or $EndTime.HasValue)) {
            Write-NCMessage "Use either -ChooseDayFromCalendar or -StartTime/-EndTime, not both." -Level ERROR
            return
        }

        $defaultTemplate = "I'm out of office and will have limited access to my mailbox.<br />I will reply to your e-mail as soon as possible.<br /><br />Have a nice day."
        $internal = if ($PSBoundParameters.ContainsKey('InternalMessage')) { $InternalMessage } else { $currentConfig.InternalMessage }
        if ([string]::IsNullOrWhiteSpace($internal)) { $internal = $defaultTemplate }

        $external = if ($PSBoundParameters.ContainsKey('ExternalMessage')) { $ExternalMessage } else { $currentConfig.ExternalMessage }
        if ([string]::IsNullOrWhiteSpace($external)) { $external = $internal }

        $state = 'Enabled'
        if ($ChooseDayFromCalendar.IsPresent) {
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

            function Get-DateFromCalendar {
                param([string]$Title)

                $form = New-Object Windows.Forms.Form
                $form.Size = New-Object Drawing.Size @(220, 210)
                $form.StartPosition = "CenterScreen"
                $form.KeyPreview = $true

                $calendar = New-Object System.Windows.Forms.MonthCalendar
                $calendar.ShowTodayCircle = $true
                $calendar.MaxSelectionCount = 1
                $form.Controls.Add($calendar)
                $form.Topmost = $true
                $form.Text = $Title

                $selectedDate = $null
                $form.Add_KeyDown({
                        if ($_.KeyCode -eq "Enter") {
                            Set-Variable -Name selectedDate -Value $calendar.SelectionStart -Scope 1
                            $form.Close()
                        }
                        elseif ($_.KeyCode -eq "Escape") {
                            $form.Close()
                        }
                    })

                [void]$form.ShowDialog()
                return $selectedDate
            }

            Write-NCMessage "Select the first day of absence in the popup and press Enter." -Level INFO
            $StartTime = Get-DateFromCalendar -Title "Select OoO start date"
            if (-not $StartTime) {
                Write-NCMessage "You must select at least one day from the calendar." -Level ERROR
                return
            }

            Write-NCMessage "Select the last day of absence in the popup and press Enter." -Level INFO
            $EndTime = Get-DateFromCalendar -Title "Select OoO end date"
            if (-not $EndTime) {
                Write-NCMessage "You must select at least one day from the calendar." -Level ERROR
                return
            }

            $state = 'Scheduled'
        }
        elseif ($StartTime.HasValue -and $EndTime.HasValue) {
            $state = 'Scheduled'
        }

        $setParams = @{
            Identity         = $SourceMailbox
            AutoReplyState   = $state
            InternalMessage  = $internal
            ExternalMessage  = $external
            ExternalAudience = $ExternalAudience
        }

        if ($state -eq 'Scheduled') {
            $setParams.StartTime = $StartTime.Value
            $setParams.EndTime = $EndTime.Value
        }

        $action = if ($state -eq 'Scheduled') {
            "Schedule out-of-office from $($StartTime.Value) to $($EndTime.Value)"
        }
        else {
            "Enable out-of-office replies"
        }

        if (-not $PSCmdlet.ShouldProcess($SourceMailbox, $action)) {
            return
        }

        try {
            Set-MailboxAutoReplyConfiguration @setParams -ErrorAction Stop
            Write-NCMessage ("Out-of-office {0} for {1}." -f $state.ToLowerInvariant(), $SourceMailbox) -Level SUCCESS
        }
        catch {
            Write-NCMessage "Unable to configure out-of-office for '$SourceMailbox'. $($_.Exception.Message)" -Level ERROR
            return
        }

        $updated = Get-MailboxAutoReplyConfiguration -Identity $SourceMailbox -ErrorAction SilentlyContinue
        if ($PassThru.IsPresent -and $updated) { $updated }
    }

    end { Restore-ProgressAndInfoPreferences }
}


