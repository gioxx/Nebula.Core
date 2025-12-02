#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Quarantine ===========================================================================================================================

function Export-QuarantineEml {
    <#
    .SYNOPSIS
        Exports a quarantined message as an EML file.
    .DESCRIPTION
        Retrieves the quarantined message by MessageId, writes the decoded EML to the specified folder,
        optionally opens it, and can release the message to all recipients.
    .PARAMETER MessageId
        MessageId of the quarantined e-mail (with or without angle brackets).
    .PARAMETER DestinationFolder
        Folder where the EML file will be written. Defaults to the current directory.
    .PARAMETER OpenFile
        Open the exported file after saving it.
    .PARAMETER ReleaseToAll
        Release the message to all recipients after export.
    .PARAMETER ReportFalsePositive
        Also report the message as a false positive when releasing.
    .EXAMPLE
        Export-QuarantineEml -MessageId 20230617142935.F5B74194B266E458@contoso.com -DestinationFolder C:\Temp -OpenFile -ReleaseToAll
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$MessageId,
        [string]$DestinationFolder,
        [switch]$OpenFile,
        [switch]$ReleaseToAll,
        [switch]$ReportFalsePositive
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $folder = Test-Folder -Path $DestinationFolder
        }
        catch {
            Write-NCMessage "Destination folder is not valid. $($_.Exception.Message)" -Level ERROR
            return
        }

        $normalizedId = ConvertTo-QuarantineMessageId -MessageId $MessageId

        try {
            $message = Get-QuarantineMessage -MessageId $normalizedId -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to find quarantined message '$normalizedId'. $($_.Exception.Message)" -Level ERROR
            return
        }

        try {
            $exported = $message | Export-QuarantineMessage
            $bytes = [Convert]::FromBase64String($exported.eml)
        }
        catch {
            Write-NCMessage "Unable to export quarantined message '$normalizedId'. $($_.Exception.Message)" -Level ERROR
            return
        }

        $invalidChars = [Regex]::Escape(([IO.Path]::GetInvalidFileNameChars() -join ''))
        $safeNameSource = if ($message.Subject) { $message.Subject } else { $message.MessageId }
        $safeBaseName = [Regex]::Replace((Format-OutputString -Value $safeNameSource -MaxLength 60), "[$invalidChars]", '_')
        if ([string]::IsNullOrWhiteSpace($safeBaseName)) {
            $safeBaseName = 'QuarantineMessage'
        }
        $emlPath = New-File (Join-Path -Path $folder -ChildPath "$safeBaseName.eml")

        try {
            [IO.File]::WriteAllBytes($emlPath, $bytes)
            Write-NCMessage ("Saved quarantined message to {0}" -f $emlPath) -Level SUCCESS
        }
        catch {
            Write-NCMessage "Unable to write EML file. $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($OpenFile.IsPresent) {
            try {
                Invoke-Item -LiteralPath $emlPath
            }
            catch {
                Write-NCMessage "File exported but could not be opened automatically. $($_.Exception.Message)" -Level WARNING
            }
        }

        if ($ReleaseToAll.IsPresent) {
            $action = "Release quarantined message '$($message.Subject)'"
            if ($PSCmdlet.ShouldProcess($message.MessageId, $action)) {
                try {
                    $releaseParams = @{
                        Identity            = $message.Identity
                        ReleaseToAll        = $true
                        Confirm             = $false
                        ReportFalsePositive = $ReportFalsePositive.IsPresent
                    }
        Release-QuarantineMessage @releaseParams | Out-Null
                    Write-NCMessage "Released quarantined message $($message.MessageId) to all recipients." -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Unable to release quarantined message. $($_.Exception.Message)" -Level ERROR
                }
            }
        }

        [pscustomobject]@{
            MessageId           = $message.MessageId
            Subject             = $message.Subject
            QuarantineTypes     = $message.QuarantineTypes
            EmlPath             = $emlPath
            ReleasedToAll       = $ReleaseToAll.IsPresent
            ReportFalsePositive = $ReportFalsePositive.IsPresent -and $ReleaseToAll.IsPresent
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Get-QuarantineFrom {
    <#
    .SYNOPSIS
        Lists quarantined messages from specific senders.
    .DESCRIPTION
        Retrieves quarantine entries for the provided sender addresses, expanding message details
        and returning a consistent set of properties.
    .PARAMETER SenderAddress
        One or more sender addresses to query. Accepts pipeline input.
    .PARAMETER IncludeReleased
        Include messages already released (default hides them).
    .EXAMPLE
        Get-QuarantineFrom -SenderAddress mario.rossi@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Sender')]
        [string[]]$SenderAddress,
        [switch]$IncludeReleased
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        foreach ($currentSender in $SenderAddress) {
            if ([string]::IsNullOrWhiteSpace($currentSender)) { continue }
            Write-NCMessage ("Searching quarantined messages from {0} ..." -f $currentSender) -Level INFO

            try {
                $messages = Get-QuarantineMessage -SenderAddress $currentSender -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Unable to retrieve messages for '$currentSender'. $($_.Exception.Message)" -Level ERROR
                continue
            }

            foreach ($msg in $messages) {
                try {
                    $details = Get-QuarantineMessage -Identity $msg.Identity -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to load message details for '$($msg.Identity)'. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                if (-not $IncludeReleased.IsPresent -and $details.Released) {
                    continue
                }

                $results.Add([pscustomobject]@{
                        Subject          = Format-OutputString -Value $details.Subject
                        SenderAddress    = $details.SenderAddress
                        RecipientAddress = $details.RecipientAddress
                        ReceivedTime     = $details.ReceivedTime
                        QuarantineTypes  = $details.QuarantineTypes
                        Released         = $details.Released
                        ReleasedUser     = $details.ReleasedUser
                        MessageId        = $details.MessageId
                        Identity         = $details.Identity
                    }) | Out-Null
            }
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
        $results
    }
}

function Get-QuarantineFromDomain {
    <#
    .SYNOPSIS
        Lists quarantined messages from specific sender domains.
    .DESCRIPTION
        Retrieves quarantine entries where the sender's domain matches the provided values.
    .PARAMETER SenderDomain
        One or more domains (e.g. contoso.com). Accepts pipeline input.
    .PARAMETER IncludeReleased
        Include messages already released (default hides them).
    .EXAMPLE
        Get-QuarantineFromDomain -SenderDomain contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$SenderDomain,
        [switch]$IncludeReleased
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        foreach ($domain in $SenderDomain) {
            if ([string]::IsNullOrWhiteSpace($domain)) { continue }
            Write-NCMessage ("Searching quarantined messages from *@{0} ..." -f $domain) -Level INFO

            try {
                $messages = Get-QuarantineMessage -ErrorAction Stop | Where-Object { $_.SenderAddress -like "*@$domain" }
            }
            catch {
                Write-NCMessage "Unable to retrieve messages for domain '$domain'. $($_.Exception.Message)" -Level ERROR
                continue
            }

            foreach ($msg in $messages) {
                try {
                    $details = Get-QuarantineMessage -Identity $msg.Identity -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to load message details for '$($msg.Identity)'. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                if (-not $IncludeReleased.IsPresent -and $details.Released) {
                    continue
                }

                $results.Add([pscustomobject]@{
                        Subject          = Format-OutputString -Value $details.Subject
                        SenderAddress    = $details.SenderAddress
                        RecipientAddress = $details.RecipientAddress
                        ReceivedTime     = $details.ReceivedTime
                        QuarantineTypes  = $details.QuarantineTypes
                        Released         = $details.Released
                        ReleasedUser     = $details.ReleasedUser
                        MessageId        = $details.MessageId
                        Identity         = $details.Identity
                    }) | Out-Null
            }
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
        $results
    }
}

function Get-QuarantineToRelease {
    <#
    .SYNOPSIS
        Retrieves quarantine messages pending release.
    .DESCRIPTION
        Pulls quarantined messages within a date range, optionally shows a grid for selection,
        exports reports, and can release or delete the selected entries.
    .PARAMETER ChooseDayFromCalendar
        Pick a single day using a calendar popup.
    .PARAMETER Interval
        Number of days back from today to search (1-30). Ignored when using -ChooseDayFromCalendar.
    .PARAMETER GridView
        Display results in Out-GridView and return only the selected rows.
    .PARAMETER Csv
        Export all retrieved entries to CSV in the chosen folder (or current directory).
    .PARAMETER Html
        Export all retrieved entries to HTML using PSWriteHTML if available.
    .PARAMETER OutputFolder
        Target folder for CSV/HTML exports.
    .PARAMETER ReleaseSelected
        Release selected (or all) entries. Requires confirmation (supports -WhatIf).
    .PARAMETER DeleteSelected
        Delete selected (or all) entries. Requires confirmation (supports -WhatIf).
    .PARAMETER ReportFalsePositive
        When releasing, also report messages as false positives.
    .EXAMPLE
        Get-QuarantineToRelease -Interval 7 -GridView -ReleaseSelected -ReportFalsePositive
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interval', SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Calendar')]
        [switch]$ChooseDayFromCalendar,
        [Parameter(ParameterSetName = 'Interval', Mandatory)]
        [ValidateRange(1, 30)]
        [int]$Interval,
        [switch]$GridView,
        [switch]$Html,
        [switch]$Csv,
        [string]$OutputFolder,
        [switch]$ReleaseSelected,
        [switch]$DeleteSelected,
        [switch]$ReportFalsePositive
    )

    begin {
        if ($ReleaseSelected.IsPresent -and $DeleteSelected.IsPresent) {
            throw "Specify either -ReleaseSelected or -DeleteSelected, not both."
        }
        Set-ProgressAndInfoPreferences
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        $startDate = $null
        $endDate = $null

        if ($ChooseDayFromCalendar.IsPresent) {
            [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
            [void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")

            $form = New-Object Windows.Forms.Form
            $form.Size = New-Object Drawing.Size @(200, 190)
            $form.StartPosition = "CenterScreen"
            $form.KeyPreview = $true

            $calendar = New-Object System.Windows.Forms.MonthCalendar
            $calendar.ShowTodayCircle = $true
            $calendar.MaxSelectionCount = 1
            $form.Controls.Add($calendar)
            $form.Topmost = $true
            $form.Text = "Select the day to be analyzed"

            $selectedDate = $null
            $form.Add_KeyDown({
                    if ($_.KeyCode -eq "Enter") {
                        $script:selectedDate = $calendar.SelectionStart
                        $form.Close()
                    }
                    elseif ($_.KeyCode -eq "Escape") {
                        $form.Close()
                    }
                })

            [void]$form.ShowDialog()

            if ($selectedDate) {
                $startDate = $selectedDate.Date.AddDays(-1)
                $endDate = $selectedDate.Date
            }
            else {
                Write-NCMessage "You must select at least one day from the calendar." -Level ERROR
                return
            }
        }
        else {
            $startDate = (Get-Date).AddDays(-$Interval)
            $endDate = Get-Date
        }

        $page = 1
        $quarantined = @()
        do {
            try {
                $pageData = Get-QuarantineMessage -StartReceivedDate $startDate.Date -EndReceivedDate $endDate -PageSize 1000 -ReleaseStatus NotReleased -Page $page
            }
            catch {
                Write-NCMessage "Unable to retrieve quarantine page $page. $($_.Exception.Message)" -Level ERROR
                break
            }

            $page++
            if ($pageData) {
                $quarantined += $pageData
            }
        } until (-not $pageData)

        if (-not $quarantined -or $quarantined.Count -eq 0) {
            Write-NCMessage "No quarantined messages found in the selected interval." -Level WARNING
            return
        }

        $items = $quarantined | ForEach-Object {
            [pscustomobject]@{
                SenderAddress    = $_.SenderAddress
                RecipientAddress = $_.RecipientAddress
                Subject          = $_.Subject
                ReceivedTime     = $_.ReceivedTime
                QuarantineTypes  = $_.QuarantineTypes
                Released         = $_.Released
                MessageId        = $_.MessageId
                Identity         = $_.Identity
            }
        }

        Write-NCMessage ("Retrieved {0} quarantined items from {1:d} to {2:d}." -f $items.Count, $startDate, $endDate) -Level INFO

        if ($Csv.IsPresent) {
            try {
                $folder = Test-Folder -Path $OutputFolder
                $csvPath = New-File (Join-Path -Path $folder -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-QuarantineToRelease-Report.csv")
                $items | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                Write-NCMessage ("CSV exported to {0}" -f $csvPath) -Level SUCCESS
            }
            catch {
                Write-NCMessage "Unable to export CSV. $($_.Exception.Message)" -Level ERROR
            }
        }

        if ($Html.IsPresent) {
            if (-not (Get-Module -Name PSWriteHTML -ListAvailable)) {
                Write-NCMessage "PSWriteHTML module is not available. Install it to use -Html output." -Level WARNING
            }
            else {
                try {
                    Import-Module PSWriteHTML -ErrorAction Stop
                    $folder = Test-Folder -Path $OutputFolder
                    $htmlPath = New-File (Join-Path -Path $folder -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-QuarantineToRelease-Report.html")
                    $items | Out-GridHtml | Set-Content -LiteralPath $htmlPath -Encoding UTF8
                    Write-NCMessage ("HTML exported to {0}" -f $htmlPath) -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Unable to export HTML report. $($_.Exception.Message)" -Level ERROR
                }
            }
        }

        $selection = $items
        if ($GridView.IsPresent) {
            $title = "{0} to {1} - {2} items" -f $startDate.Date, $endDate.Date, $items.Count
            $selection = $items | Sort-Object -Descending ReceivedTime | Out-GridView -Title $title -PassThru
            if (-not $selection) {
                Write-NCMessage "No items selected." -Level WARNING
                return
            }
        }

        if (-not $ReleaseSelected.IsPresent -and -not $DeleteSelected.IsPresent) {
            return $selection | Sort-Object -Property Subject
        }

        $processed = [System.Collections.Generic.List[object]]::new()
        $counter = 0
        foreach ($item in $selection) {
            $counter++
            $percentComplete = (($counter / $selection.Count) * 100)
            Write-Progress -Activity "Processing $($item.Subject)" -Status "$counter of $($selection.Count) ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            if ($ReleaseSelected.IsPresent) {
                if ($PSCmdlet.ShouldProcess($item.Subject, "Release quarantined message")) {
                    try {
                        $releaseParams = @{
                            Identity            = $item.Identity
                            ReleaseToAll        = $true
                            Confirm             = $false
                            ReportFalsePositive = $ReportFalsePositive.IsPresent
                        }
                        Release-QuarantineMessage @releaseParams | Out-Null
                        $details = Get-QuarantineMessage -Identity $item.Identity
                        $processed.Add([pscustomobject]@{
                                Subject       = Format-OutputString -Value $details.Subject
                                SenderAddress = Format-OutputString -Value $details.SenderAddress
                                Released      = $details.Released
                                ReleasedUser  = $details.ReleasedUser
                            }) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to release message '$($item.Subject)'. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }
            elseif ($DeleteSelected.IsPresent) {
                if ($PSCmdlet.ShouldProcess($item.Subject, "Delete quarantined message permanently")) {
                    try {
                        Delete-QuarantineMessage -Identity $item.Identity -Confirm:$false
                        $processed.Add([pscustomobject]@{
                                Subject       = Format-OutputString -Value $item.Subject
                                SenderAddress = Format-OutputString -Value $item.SenderAddress
                                Deleted       = $true
                            }) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to delete message '$($item.Subject)'. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }
        }

        Write-Progress -Activity "Processing quarantined messages" -Completed

        if ($processed.Count -gt 0) {
            Write-NCMessage ("{0} item(s) processed." -f $processed.Count) -Level SUCCESS
            return $processed
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Unlock-QuarantineFrom {
    <#
    .SYNOPSIS
        Releases quarantined messages from specific senders.
    .DESCRIPTION
        Retrieves messages for the given senders and releases them to all recipients,
        optionally reporting them as false positives.
    .PARAMETER SenderAddress
        One or more sender addresses. Accepts pipeline input.
    .PARAMETER ReportFalsePositive
        Also report the released messages as false positives.
    .EXAMPLE
        Unlock-QuarantineFrom -SenderAddress mario.rossi@contoso.com -ReportFalsePositive
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Sender')]
        [string[]]$SenderAddress,
        [switch]$ReportFalsePositive
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        foreach ($currentSender in $SenderAddress) {
            if ([string]::IsNullOrWhiteSpace($currentSender)) { continue }
            Write-NCMessage ("Search for quarantined messages from {0} ..." -f $currentSender) -Level INFO

            try {
                $messages = Get-QuarantineMessage -SenderAddress $currentSender -ErrorAction Stop | Where-Object { $_.ReleaseStatus -ne "Released" -and $null -ne $_.QuarantinedUser }
                Write-NCMessage "Found $($messages.Count) message(s) from $currentSender not yet released." -Level VERBOSE
            }
            catch {
                Write-NCMessage "Unable to retrieve messages for '$currentSender'. $($_.Exception.Message)" -Level ERROR
                continue
            }

            foreach ($msg in $messages) {
                if ($PSCmdlet.ShouldProcess($msg.Identity, "Release quarantined message")) {
                    try {
                        Write-NCMessage "Trying to release $($msg.Identity) to $($msg.RecipientAddress)  ..." -Level VERBOSE
                        $releaseParams = @{
                            Identity            = $msg.Identity
                            ReleaseToAll        = $true
                            Confirm             = $false
                            ReportFalsePositive = $ReportFalsePositive.IsPresent
                        }
                        Release-QuarantineMessage @releaseParams | Out-Null
                        $details = Get-QuarantineMessage -Identity $msg.Identity
                        $results.Add([pscustomobject]@{
                            Subject       = Format-OutputString -Value $details.Subject 40
                            SenderAddress = $details.SenderAddress
                            ReceivedTime  = $details.ReceivedTime
                            Released      = $details.Released
                            ReleasedUser  = $details.ReleasedUser
                        }) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to release message '$($msg.Identity)'. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
        $results | Select-Object Subject, SenderAddress, ReceivedTime, ReleasedUser | Format-Table -AutoSize
    }
}

Set-Alias -Name rqf -Value Unlock-QuarantineFrom -Description "Release Quarantine from (function)"

function Unlock-QuarantineMessageId {
    <#
    .SYNOPSIS
        Releases quarantined messages by MessageId.
    .DESCRIPTION
        Accepts MessageId values (with or without angle brackets), releases the messages to all recipients,
        and returns the release status.
    .PARAMETER MessageId
        One or more MessageId values. Accepts pipeline input.
    .PARAMETER ReportFalsePositive
        Also report the released messages as false positives.
    .EXAMPLE
        Unlock-QuarantineMessageId -MessageId CAH_w85uSio_cz4HsFxJAGQDd-kzxGijLaMagZU95m3A1G8hWBA@mail.contoso.com
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Id')]
        [string[]]$MessageId,
        [switch]$ReportFalsePositive
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        foreach ($id in $MessageId) {
            $normalizedId = ConvertTo-QuarantineMessageId -MessageId $id

            try {
                $messages = Get-QuarantineMessage -MessageId $normalizedId -ErrorAction Stop | Where-Object { $_.ReleaseStatus -ne "Released" -and $_.QuarantinedUser }
            }
            catch {
                Write-NCMessage "Unable to retrieve quarantined message '$normalizedId'. $($_.Exception.Message)" -Level ERROR
                continue
            }

            if (-not $messages -or $messages.Count -eq 0) {
                Write-NCMessage "No quarantined messages to release with id $normalizedId (already released or not found yet)." -Level WARNING
                continue
            }

            foreach ($msg in $messages) {
                if ($PSCmdlet.ShouldProcess($msg.Identity, "Release quarantined message")) {
                    try {
                        $releaseParams = @{
                            Identity            = $msg.Identity
                            ReleaseToAll        = $true
                            Confirm             = $false
                            ReportFalsePositive = $ReportFalsePositive.IsPresent
                        }
                        Release-QuarantineMessage @releaseParams | Out-Null
                        $details = Get-QuarantineMessage -Identity $msg.Identity
                        $results.Add([pscustomobject]@{
                                Subject       = Format-OutputString -Value $details.Subject 40
                                SenderAddress = $details.SenderAddress
                                ReceivedTime  = $details.ReceivedTime
                                Released      = $details.Released
                                ReleasedUser  = $details.ReleasedUser
                            }) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to release message '$($msg.Identity)'. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
        $results
    }
}
