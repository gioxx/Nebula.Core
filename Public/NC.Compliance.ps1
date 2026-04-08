#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Compliance helpers ===================================================================================================================

function Search-MboxCutoffWindow {
    <#
    .SYNOPSIS
        Creates or reuses a Purview Compliance Search to isolate mailbox items by date criteria.
    .DESCRIPTION
        Builds a content query for a target mailbox (items before a cutoff date, or within a fixed date range),
        runs a compliance estimate, and can optionally run a Preview action with sampled output lines.
        Useful to isolate candidate items before export/cleanup workflows.
    .PARAMETER Mailbox
        Target mailbox (UPN or SMTP address). Accepts pipeline input.
    .PARAMETER Mode
        Query mode:
        - BeforeCutoff: items older than CutoffDate
        - Range: items in [StartDate, EndDate) (end exclusive)
    .PARAMETER CutoffDate
        Cutoff date used when Mode is BeforeCutoff.
    .PARAMETER StartDate
        Start date used when Mode is Range.
    .PARAMETER EndDate
        End date (exclusive) used when Mode is Range.
    .PARAMETER Preview
        Create a Preview action and return a limited sample of preview items.
    .PARAMETER PreviewCount
        Number of preview entries to sample.
    .PARAMETER ExistingSearchName
        Explicit compliance search name to reuse.
    .PARAMETER UseExistingOnly
        Do not create/modify search definition; only run estimate/preview on ExistingSearchName.
    .PARAMETER PollingSeconds
        Polling interval in seconds while waiting for Compliance Search/Action completion.
    .PARAMETER MaxWaitMinutes
        Maximum wait time before aborting search/action polling.
    .EXAMPLE
        Search-MboxCutoffWindow -Mailbox 'user@contoso.com' -Mode BeforeCutoff -CutoffDate '2025-01-01'
    .EXAMPLE
        Search-MboxCutoffWindow -Mailbox 'user@contoso.com' -Mode Range -StartDate '2025-01-01' -EndDate '2025-02-01' -Preview -PreviewCount 25
    .EXAMPLE
        Search-MboxCutoffWindow -Mailbox 'user@contoso.com' -ExistingSearchName 'Isolate_Pre_20250101_140530' -UseExistingOnly
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity', 'UserPrincipalName', 'SourceMailbox')]
        [string]$Mailbox,
        [ValidateSet('BeforeCutoff', 'Range')]
        [string]$Mode = 'BeforeCutoff',
        [datetime]$CutoffDate = [datetime]'2025-01-01',
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$Preview,
        [ValidateRange(1, 500)]
        [int]$PreviewCount = 50,
        [string]$ExistingSearchName,
        [switch]$UseExistingOnly,
        [ValidateRange(5, 300)]
        [int]$PollingSeconds = 10,
        [ValidateRange(1, 240)]
        [int]$MaxWaitMinutes = 60
    )

    begin {
        Set-ProgressAndInfoPreferences
    }

    process {
        if (-not (Test-EOLConnection)) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
            return
        }

        try {
            if (-not (Get-Command -Name Get-ComplianceSearch -ErrorAction SilentlyContinue)) {
                Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop | Out-Null
            }
            else {
                try {
                    Get-ComplianceSearch -ErrorAction Stop | Select-Object -First 1 | Out-Null
                }
                catch {
                    Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop | Out-Null
                }
            }
        }
        catch {
            Write-NCMessage "Unable to connect to Microsoft Purview Compliance PowerShell. $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($Mode -eq 'Range') {
            if (-not $PSBoundParameters.ContainsKey('StartDate') -or -not $PSBoundParameters.ContainsKey('EndDate')) {
                Write-NCMessage "Range mode requires both -StartDate and -EndDate." -Level ERROR
                return
            }
            if ($EndDate -le $StartDate) {
                Write-NCMessage "EndDate must be greater than StartDate." -Level ERROR
                return
            }
        }

        $query = if ($Mode -eq 'BeforeCutoff') {
            $cutoff = $CutoffDate.ToString('MM/dd/yyyy')
            "(Received<$cutoff) OR (Sent<$cutoff)"
        }
        else {
            $start = $StartDate.ToString('MM/dd/yyyy')
            $end = $EndDate.ToString('MM/dd/yyyy')
            "((Received>=$start AND Received<$end) OR (Sent>=$start AND Sent<$end))"
        }

        $searchName = if (-not [string]::IsNullOrWhiteSpace($ExistingSearchName)) {
            $ExistingSearchName
        }
        else {
            $prefix = if ($Mode -eq 'BeforeCutoff') { "Isolate_Pre_$($CutoffDate.ToString('yyyyMMdd'))" } else { 'Isolate_Range' }
            "{0}_{1}" -f $prefix, (Get-Date -Format 'yyyyMMdd_HHmmss')
        }

        Write-NCMessage "Mailbox: $Mailbox" -Level INFO
        Write-NCMessage "Mode: $Mode" -Level INFO
        Write-NCMessage "Query: $query" -Level INFO
        Write-NCMessage "Search: $searchName" -Level INFO

        if ($UseExistingOnly.IsPresent -and [string]::IsNullOrWhiteSpace($ExistingSearchName)) {
            Write-NCMessage "UseExistingOnly requires -ExistingSearchName." -Level ERROR
            return
        }

        if (-not $UseExistingOnly.IsPresent) {
            $existing = Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue
            if (-not $existing) {
                if ($PSCmdlet.ShouldProcess($searchName, "Create Compliance Search for mailbox '$Mailbox'")) {
                    try {
                        New-ComplianceSearch -Name $searchName -ExchangeLocation $Mailbox -ContentMatchQuery $query -ErrorAction Stop | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to create compliance search '$searchName'. $($_.Exception.Message)" -Level ERROR
                        return
                    }
                }
                else {
                    return
                }
            }
            else {
                Write-NCMessage "Compliance search '$searchName' already exists. Reusing it." -Level WARNING
            }
        }
        else {
            $existing = Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue
            if (-not $existing) {
                Write-NCMessage "Existing search '$searchName' not found." -Level ERROR
                return
            }
        }

        if (-not $PSCmdlet.ShouldProcess($searchName, "Run compliance estimate")) {
            return
        }

        try {
            Start-ComplianceSearch -Identity $searchName -ErrorAction Stop | Out-Null
        }
        catch {
            Write-NCMessage "Unable to start compliance search '$searchName'. $($_.Exception.Message)" -Level ERROR
            return
        }

        $deadline = (Get-Date).AddMinutes($MaxWaitMinutes)
        $searchStatus = $null
        $search = $null
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds $PollingSeconds
            $search = Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue
            if (-not $search) {
                continue
            }

            $searchStatus = [string]$search.Status
            if ($searchStatus -in @('Completed', 'PartiallySucceeded', 'PartiallyCompleted', 'Failed')) {
                break
            }
        }

        if (-not $search) {
            Write-NCMessage "Unable to read compliance search '$searchName' status." -Level ERROR
            return
        }

        if ($searchStatus -notin @('Completed', 'PartiallySucceeded', 'PartiallyCompleted')) {
            if ($searchStatus -eq 'Failed') {
                Write-NCMessage "Compliance search '$searchName' failed." -Level ERROR
            }
            else {
                Write-NCMessage "Timeout while waiting for compliance search '$searchName' completion." -Level ERROR
            }
            return
        }

        $estimatedItems = [int]$search.Items
        $unindexedItems = $search.UnindexedItems
        Write-NCMessage ("Search completed. Estimated items: {0}" -f $estimatedItems) -Level SUCCESS
        if ($null -ne $unindexedItems) {
            Write-NCMessage ("Estimated unindexed items: {0}" -f $unindexedItems) -Level WARNING
        }

        $previewStatus = $null
        $previewSample = @()
        if ($Preview.IsPresent) {
            if (-not $PSCmdlet.ShouldProcess($searchName, "Create Preview action")) {
                return
            }

            try {
                $previewAction = New-ComplianceSearchAction -SearchName $searchName -Preview -Force -Confirm:$false -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Unable to create preview action for '$searchName'. $($_.Exception.Message)" -Level ERROR
                return
            }

            $actionDeadline = (Get-Date).AddMinutes($MaxWaitMinutes)
            $actionResult = $null
            while ((Get-Date) -lt $actionDeadline) {
                Start-Sleep -Seconds ([Math]::Max($PollingSeconds, 10))
                $actionResult = Get-ComplianceSearchAction -Identity $previewAction.Identity -ErrorAction SilentlyContinue
                if (-not $actionResult) {
                    continue
                }

                $previewStatus = [string]$actionResult.Status
                if ($previewStatus -in @('Completed', 'PartiallyCompleted', 'Failed')) {
                    break
                }
            }

            if (-not $actionResult) {
                Write-NCMessage "Unable to read preview action status for '$searchName'." -Level ERROR
                return
            }

            if ($previewStatus -eq 'Failed') {
                Write-NCMessage ("Preview action failed for '{0}'. {1}" -f $searchName, [string]$actionResult.Errors) -Level ERROR
                return
            }

            if ($previewStatus -notin @('Completed', 'PartiallyCompleted')) {
                Write-NCMessage "Timeout while waiting for preview action completion for '$searchName'." -Level ERROR
                return
            }

            $rawResults = [string]$actionResult.Results
            if (-not [string]::IsNullOrWhiteSpace($rawResults)) {
                $previewSample = @($rawResults -split ",\s*(?=Location:)" | Select-Object -First $PreviewCount)
            }

            Write-NCMessage ("Preview action status: {0}" -f $previewStatus) -Level SUCCESS
            if ($previewSample.Count -gt 0) {
                Write-NCMessage ("Preview sample lines returned: {0}" -f $previewSample.Count) -Level INFO
            }
            else {
                Write-NCMessage "Preview completed, but no sample lines were returned in PowerShell output. Check Purview portal for details." -Level WARNING
            }
        }

        Write-NCMessage "Purview portal: https://purview.microsoft.com" -Level INFO
        Write-NCMessage "Path: eDiscovery -> Content search -> open the search -> Actions/Export" -Level INFO

        [pscustomobject]@{
            Mailbox        = $Mailbox
            Mode           = $Mode
            Query          = $query
            SearchName     = $searchName
            SearchStatus   = $searchStatus
            EstimatedItems = $estimatedItems
            UnindexedItems = $unindexedItems
            PreviewStatus  = $previewStatus
            PreviewSample  = $previewSample
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
    }
}

function Set-MboxMrmCleanup {
    <#
    .SYNOPSIS
        Applies a one-shot MRM cleanup policy to a mailbox.
    .DESCRIPTION
        Computes a safe retention age from a fixed cutoff date plus a safety buffer, then creates/updates
        a retention tag and policy, assigns the policy to the mailbox, and optionally triggers Managed Folder Assistant.
        Intended for temporary cleanup workflows where older items should be targeted while preserving recent data.
    .PARAMETER Mailbox
        Target mailbox identity (UPN or SMTP). Accepts pipeline input.
    .PARAMETER FixedCutoffDate
        Fixed cutoff date used to compute AgeLimitForRetention in days.
    .PARAMETER SafetyBufferDays
        Additional safety days added to the computed retention age.
    .PARAMETER RetentionAction
        Retention action for the tag (`DeleteAndAllowRecovery` or `PermanentlyDelete`).
    .PARAMETER TagName
        Retention tag name. If omitted, an automatic name based on cutoff date is used.
    .PARAMETER PolicyName
        Retention policy name. If omitted, an automatic name based on cutoff date is used.
    .PARAMETER RunAssistant
        Trigger Managed Folder Assistant (FullCrawl) after policy assignment.
    .EXAMPLE
        Set-MboxMrmCleanup -Mailbox 'user@contoso.com' -FixedCutoffDate '2025-01-01' -SafetyBufferDays 7
    .EXAMPLE
        Set-MboxMrmCleanup -Mailbox 'user@contoso.com' -RetentionAction PermanentlyDelete -RunAssistant -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity', 'UserPrincipalName', 'SourceMailbox')]
        [string]$Mailbox,
        [datetime]$FixedCutoffDate = [datetime]'2025-01-01',
        [ValidateRange(0, 60)]
        [int]$SafetyBufferDays = 7,
        [ValidateSet('DeleteAndAllowRecovery', 'PermanentlyDelete')]
        [string]$RetentionAction = 'DeleteAndAllowRecovery',
        [string]$TagName,
        [string]$PolicyName,
        [switch]$RunAssistant
    )

    begin {
        Set-ProgressAndInfoPreferences
    }

    process {
        if (-not (Test-EOLConnection)) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
            return
        }

        $now = Get-Date
        $ageDays = [int]([math]::Ceiling(($now - $FixedCutoffDate).TotalDays)) + $SafetyBufferDays
        if ($ageDays -lt 1) { $ageDays = 1 }

        if ([string]::IsNullOrWhiteSpace($TagName)) {
            $TagName = "OneShot_PreCutoff_$($FixedCutoffDate.ToString('yyyyMMdd'))"
        }
        if ([string]::IsNullOrWhiteSpace($PolicyName)) {
            $PolicyName = "OneShot_PreCutoff_$($FixedCutoffDate.ToString('yyyyMMdd'))"
        }

        Write-NCMessage "Mailbox: $Mailbox" -Level INFO
        Write-NCMessage ("Fixed cutoff date: {0:yyyy-MM-dd}" -f $FixedCutoffDate) -Level INFO
        Write-NCMessage ("Safety buffer (days): {0}" -f $SafetyBufferDays) -Level INFO
        Write-NCMessage ("Computed AgeLimitForRetention (days): {0}" -f $ageDays) -Level SUCCESS
        Write-NCMessage ("Retention action: {0}" -f $RetentionAction) -Level INFO
        Write-NCMessage ("Tag name: {0}" -f $TagName) -Level INFO
        Write-NCMessage ("Policy name: {0}" -f $PolicyName) -Level INFO

        try {
            $tag = Get-RetentionPolicyTag -Identity $TagName -ErrorAction SilentlyContinue
            if (-not $tag) {
                if ($PSCmdlet.ShouldProcess($TagName, "Create retention policy tag")) {
                    New-RetentionPolicyTag -Name $TagName -Type All -RetentionEnabled $true -AgeLimitForRetention $ageDays -RetentionAction $RetentionAction -ErrorAction Stop | Out-Null
                    Write-NCMessage "Retention policy tag '$TagName' created." -Level SUCCESS
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess($TagName, "Update retention policy tag settings")) {
                    Set-RetentionPolicyTag -Identity $TagName -RetentionEnabled $true -AgeLimitForRetention $ageDays -RetentionAction $RetentionAction -ErrorAction Stop | Out-Null
                    Write-NCMessage "Retention policy tag '$TagName' updated." -Level SUCCESS
                }
            }
        }
        catch {
            Write-NCMessage "Unable to create/update retention policy tag '$TagName'. $($_.Exception.Message)" -Level ERROR
            return
        }

        try {
            $policy = Get-RetentionPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
            if (-not $policy) {
                if ($PSCmdlet.ShouldProcess($PolicyName, "Create retention policy with tag '$TagName'")) {
                    New-RetentionPolicy -Name $PolicyName -RetentionPolicyTagLinks $TagName -ErrorAction Stop | Out-Null
                    Write-NCMessage "Retention policy '$PolicyName' created." -Level SUCCESS
                }
            }
            else {
                $links = @($policy.RetentionPolicyTagLinks)
                if ($links -notcontains $TagName) {
                    if ($PSCmdlet.ShouldProcess($PolicyName, "Add retention policy tag link '$TagName'")) {
                        Set-RetentionPolicy -Identity $PolicyName -RetentionPolicyTagLinks ($links + $TagName) -ErrorAction Stop | Out-Null
                        Write-NCMessage "Retention policy '$PolicyName' updated with tag '$TagName'." -Level SUCCESS
                    }
                }
                else {
                    Write-NCMessage "Retention policy '$PolicyName' already includes '$TagName'." -Level INFO
                }
            }
        }
        catch {
            Write-NCMessage "Unable to create/update retention policy '$PolicyName'. $($_.Exception.Message)" -Level ERROR
            return
        }

        try {
            if ($PSCmdlet.ShouldProcess($Mailbox, "Assign retention policy '$PolicyName'")) {
                Set-Mailbox -Identity $Mailbox -RetentionPolicy $PolicyName -ErrorAction Stop | Out-Null
                Write-NCMessage "Retention policy '$PolicyName' assigned to '$Mailbox'." -Level SUCCESS
            }
        }
        catch {
            Write-NCMessage "Unable to assign retention policy '$PolicyName' to '$Mailbox'. $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($RunAssistant.IsPresent) {
            try {
                if ($PSCmdlet.ShouldProcess($Mailbox, 'Trigger Managed Folder Assistant (FullCrawl)')) {
                    Start-ManagedFolderAssistant -Identity $Mailbox -FullCrawl -ErrorAction Stop
                    Write-NCMessage "Managed Folder Assistant triggered for '$Mailbox'." -Level SUCCESS
                }
            }
            catch {
                Write-NCMessage "Unable to trigger Managed Folder Assistant for '$Mailbox'. $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            Write-NCMessage "Managed Folder Assistant not triggered. Use -RunAssistant to start it." -Level INFO
        }

        [pscustomobject]@{
            Mailbox            = $Mailbox
            FixedCutoffDate    = $FixedCutoffDate
            SafetyBufferDays   = $SafetyBufferDays
            AgeLimitDays       = $ageDays
            RetentionAction    = $RetentionAction
            TagName            = $TagName
            PolicyName         = $PolicyName
            RollbackCommand    = "Set-Mailbox -Identity '$Mailbox' -RetentionPolicy `$null"
            RemovePolicyHint   = "Remove-RetentionPolicy -Identity '$PolicyName'"
            RemoveTagHint      = "Remove-RetentionPolicyTag -Identity '$TagName'"
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
    }
}

