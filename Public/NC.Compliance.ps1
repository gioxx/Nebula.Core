#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Compliance helpers ===================================================================================================================

function Get-MboxMrmCleanup {
    <#
    .SYNOPSIS
        Lists retention tags, policies, and mailbox assignments for MRM cleanup workflows.
    .DESCRIPTION
        Connects to Exchange Online and inventories retention policies. The output defaults to the temporary
        Nebula.Core cleanup objects created by Set-MboxMrmCleanup. Use -AllTenantObjects to list every
        retention policy in the tenant. Each row includes the linked tag details inline and the mailbox count,
        so you can spot temporary cleanup objects and decide what can be removed safely.
    .PARAMETER Identity
        Retention policy name or linked tag name to inspect. When omitted, temporary cleanup policies are returned.
    .PARAMETER AllTenantObjects
        Include every retention policy in the tenant instead of limiting the inventory to temporary Nebula.Core
        cleanup objects.
    .PARAMETER Detailed
        Include the linked tag names and mailbox lists in the output.
    .EXAMPLE
        Get-MboxMrmCleanup
    .EXAMPLE
        Get-MboxMrmCleanup -Detailed
    .EXAMPLE
        Get-MboxMrmCleanup -AllTenantObjects
    .EXAMPLE
        Get-MboxMrmCleanup -Identity OneShot_PreCutoff_20250101
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'TagName', 'PolicyName')]
        [string[]]$Identity,
        [switch]$AllTenantObjects,
        [switch]$Detailed
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($entry in $Identity) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $targets.Add($entry.Trim()) | Out-Null
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

            $tagObjects = @()
            $policyObjects = @()

            try {
                $tagObjects = @(Get-RetentionPolicyTag -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to retrieve retention policy tags. $($_.Exception.Message)" -Level ERROR
                return
            }

            try {
                $policyObjects = @(Get-RetentionPolicy -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to retrieve retention policies. $($_.Exception.Message)" -Level ERROR
                return
            }

            $mailboxObjects = @()
            try {
                if (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue) {
                    $mailboxObjects = @(Get-EXOMailbox -ResultSize Unlimited -Properties RetentionPolicy, RecipientTypeDetails, DisplayName, PrimarySmtpAddress -ErrorAction Stop)
                }
                else {
                    $mailboxObjects = @(Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction Stop)
                }
            }
            catch {
                Write-NCMessage "Unable to retrieve mailbox assignments for retention policies. $($_.Exception.Message)" -Level WARNING
                $mailboxObjects = @()
            }

            $mailboxesByPolicy = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($mailbox in $mailboxObjects) {
                $retentionPolicyName = [string]$mailbox.RetentionPolicy
                if ([string]::IsNullOrWhiteSpace($retentionPolicyName)) {
                    continue
                }

                if (-not $mailboxesByPolicy.ContainsKey($retentionPolicyName)) {
                    $mailboxesByPolicy[$retentionPolicyName] = [System.Collections.Generic.List[object]]::new()
                }

                $mailboxesByPolicy[$retentionPolicyName].Add($mailbox) | Out-Null
            }

            $policyLinksByTag = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($policy in $policyObjects) {
                foreach ($tagLink in @($policy.RetentionPolicyTagLinks)) {
                    if ([string]::IsNullOrWhiteSpace([string]$tagLink)) {
                        continue
                    }

                    $tagName = [string]$tagLink
                    if (-not $policyLinksByTag.ContainsKey($tagName)) {
                        $policyLinksByTag[$tagName] = [System.Collections.Generic.List[string]]::new()
                    }

                    $policyLinksByTag[$tagName].Add([string]$policy.Name) | Out-Null
                }
            }

            $policyResults = foreach ($policy in $policyObjects) {
                if ($null -eq $policy) {
                    continue
                }

                $policyName = [string]$policy.Name
                $linkedTags = @($policy.RetentionPolicyTagLinks | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $primaryTagName = $linkedTags | Select-Object -First 1
                $primaryTag = $null
                if (-not [string]::IsNullOrWhiteSpace($primaryTagName)) {
                    $primaryTag = $tagObjects | Where-Object { [string]$_.Name -ieq $primaryTagName } | Select-Object -First 1
                }

                if (-not $AllTenantObjects.IsPresent -and $policyName -notlike 'OneShot_PreCutoff_*') {
                    continue
                }

                if ($targets.Count -gt 0) {
                    $matchesPolicy = $targets -contains $policyName
                    $matchesLinkedTag = $false
                    foreach ($tagName in $linkedTags) {
                        if ($targets -contains $tagName) {
                            $matchesLinkedTag = $true
                            break
                        }
                    }

                    if (-not $matchesPolicy -and -not $matchesLinkedTag) {
                        continue
                    }
                }

                $assignedMailboxes = @()
                if ($mailboxesByPolicy.ContainsKey($policyName)) {
                    $assignedMailboxes = @($mailboxesByPolicy[$policyName] | ForEach-Object {
                            if ($_.PrimarySmtpAddress) { [string]$_.PrimarySmtpAddress }
                            elseif ($_.UserPrincipalName) { [string]$_.UserPrincipalName }
                            else { [string]$_.Identity }
                        } | Sort-Object -Unique)
                }

                $ageLimitDays = if ($primaryTag -and ($primaryTag.AgeLimitForRetention -is [TimeSpan])) {
                    [int][math]::Floor($primaryTag.AgeLimitForRetention.TotalDays)
                }
                elseif ($primaryTag) {
                    $primaryTag.AgeLimitForRetention
                }
                else {
                    $null
                }

                $row = [ordered]@{
                    ObjectType           = 'Policy'
                    Identity             = $policy.Identity
                    Name                 = $policyName
                    LinkedTagName        = $primaryTagName
                    TagType              = $primaryTag.Type
                    RetentionEnabled     = $primaryTag.RetentionEnabled
                    AgeLimitForRetentionDays = $ageLimitDays
                    RetentionAction      = $primaryTag.RetentionAction
                    ConditionSummary     = if ($primaryTag) {
                        "Type={0}; AgeLimit={1}d; Action={2}; Enabled={3}" -f $primaryTag.Type, $ageLimitDays, $primaryTag.RetentionAction, $primaryTag.RetentionEnabled
                    }
                    elseif ($linkedTags.Count -gt 0) {
                        "TagLinks={0}" -f ($linkedTags -join ', ')
                    }
                    else {
                        'TagLinks='
                    }
                    LinkedTagCount       = $linkedTags.Count
                    TagLinkCount         = $linkedTags.Count
                    AssignedMailboxCount = $assignedMailboxes.Count
                }

                if ($Detailed.IsPresent) {
                    $row.LinkedTagNames = $linkedTags
                    $row.AssignedMailboxes = $assignedMailboxes
                }

                [pscustomobject]$row
            }

            $results = @($policyResults)
            if ($results.Count -eq 0) {
                Write-NCMessage "No retention policies matched the requested filters." -Level WARNING
                return
            }

            $results | Sort-Object ObjectType, Name
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Remove-MboxMrmCleanup {
    <#
    .SYNOPSIS
        Removes temporary MRM cleanup tags and policies.
    .DESCRIPTION
        Finds the specified retention policy tag and policy, moves the affected mailboxes back to the
        default retention policy or to a specific standard policy, and then removes the temporary policy
        and tag when they are no longer needed.
    .PARAMETER Identity
        Retention tag or retention policy name to remove. When omitted, all Nebula.Core temporary cleanup
        objects (names starting with OneShot_PreCutoff_) are targeted.
    .PARAMETER Mailbox
        Optional mailbox list to restore before deleting the cleanup policy. When omitted, every mailbox
        currently using the target policy is restored.
    .PARAMETER RestorePolicyName
        Explicit standard retention policy to assign back to the mailbox or mailboxes.
    .PARAMETER ClearRetentionPolicy
        Clear the mailbox retention policy instead of assigning a specific standard policy.
    .PARAMETER RemoveTag
        Remove the retention policy tag after the policy has been cleaned up.
    .PARAMETER RemovePolicy
        Remove the retention policy after the target mailboxes have been restored.
    .EXAMPLE
        Remove-MboxMrmCleanup -Identity OneShot_PreCutoff_20250101
    .EXAMPLE
        Remove-MboxMrmCleanup -PolicyName OneShot_PreCutoff_20250101 -RestorePolicyName 'Default MRM Policy'
    .EXAMPLE
        'user@contoso.com' | Remove-MboxMrmCleanup -Identity OneShot_PreCutoff_20250101 -ClearRetentionPolicy
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'TagName', 'PolicyName')]
        [string[]]$Identity,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [Alias('UserPrincipalName', 'IdentityMailbox', 'SourceMailbox')]
        [string[]]$Mailbox,
        [string]$RestorePolicyName,
        [switch]$ClearRetentionPolicy,
        [bool]$RemoveTag = $true,
        [bool]$RemovePolicy = $true
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
        $mailboxTargets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($entry in $Identity) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $targets.Add($entry.Trim()) | Out-Null
            }
        }

        foreach ($entry in $Mailbox) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $mailboxTargets.Add($entry.Trim()) | Out-Null
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

            if (-not [string]::IsNullOrWhiteSpace($RestorePolicyName) -and $ClearRetentionPolicy.IsPresent) {
                Write-NCMessage "Use either -RestorePolicyName or -ClearRetentionPolicy, not both." -Level ERROR
                return
            }

            if (-not [string]::IsNullOrWhiteSpace($RestorePolicyName)) {
                $restorePolicyCheck = Get-RetentionPolicy -Identity $RestorePolicyName -ErrorAction SilentlyContinue
                if (-not $restorePolicyCheck) {
                    Write-NCMessage "Restore policy '$RestorePolicyName' was not found. Use the exact retention policy name." -Level ERROR
                    return
                }
            }

            $tagObjects = @()
            $policyObjects = @()
            try {
                $tagObjects = @(Get-RetentionPolicyTag -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to retrieve retention policy tags. $($_.Exception.Message)" -Level ERROR
                return
            }

            try {
                $policyObjects = @(Get-RetentionPolicy -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to retrieve retention policies. $($_.Exception.Message)" -Level ERROR
                return
            }

            $targetTagNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $targetPolicyNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            if ($targets.Count -eq 0) {
                foreach ($tag in $tagObjects) {
                    if ($tag.Name -like 'OneShot_PreCutoff_*') {
                        $null = $targetTagNames.Add([string]$tag.Name)
                    }
                }

                foreach ($policy in $policyObjects) {
                    if ($policy.Name -like 'OneShot_PreCutoff_*') {
                        $null = $targetPolicyNames.Add([string]$policy.Name)
                    }
                }
            }
            else {
                foreach ($entry in $targets) {
                    $tagMatch = $tagObjects | Where-Object { [string]$_.Name -ieq $entry } | Select-Object -First 1
                    if ($tagMatch) {
                        $null = $targetTagNames.Add([string]$tagMatch.Name)
                    }

                    $policyMatch = $policyObjects | Where-Object { [string]$_.Name -ieq $entry } | Select-Object -First 1
                    if ($policyMatch) {
                        $null = $targetPolicyNames.Add([string]$policyMatch.Name)
                    }
                }
            }

            if ($targetTagNames.Count -eq 0 -and $targetPolicyNames.Count -eq 0) {
                Write-NCMessage "No matching retention tags or policies were found." -Level WARNING
                return
            }

            $mailboxesToRestore = [System.Collections.Generic.List[object]]::new()
            $policyNamesToRemove = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($policyName in $targetPolicyNames) {
                $policy = $policyObjects | Where-Object { [string]$_.Name -ieq $policyName } | Select-Object -First 1
                if ($policy) {
                    $null = $policyNamesToRemove.Add([string]$policy.Name)
                }
            }

            foreach ($tagName in $targetTagNames) {
                $tagLinkedPolicies = @($policyObjects | Where-Object { @($_.RetentionPolicyTagLinks) -contains $tagName })
                foreach ($policy in $tagLinkedPolicies) {
                    $null = $policyNamesToRemove.Add([string]$policy.Name)
                }
            }

            if ($mailboxTargets.Count -gt 0) {
                foreach ($mailboxIdentity in $mailboxTargets) {
                    try {
                        $mailboxObj = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
                        $mailboxesToRestore.Add($mailboxObj) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to read mailbox '$mailboxIdentity'. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }
            else {
                try {
                    $allMailboxes = @(Get-Mailbox -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Unable to enumerate mailboxes using retention policies. $($_.Exception.Message)" -Level WARNING
                    $allMailboxes = @()
                }

                foreach ($mailbox in $allMailboxes) {
                    if ($policyNamesToRemove.Contains([string]$mailbox.RetentionPolicy)) {
                        $mailboxesToRestore.Add($mailbox) | Out-Null
                    }
                }
            }

            $restoreTargets = [System.Collections.Generic.List[object]]::new()
            foreach ($mailbox in $mailboxesToRestore) {
                if ($null -eq $mailbox) {
                    continue
                }

                $mailboxName = if ($mailbox.PrimarySmtpAddress) { [string]$mailbox.PrimarySmtpAddress } else { [string]$mailbox.Identity }
                $targetPolicy = if (-not [string]::IsNullOrWhiteSpace($RestorePolicyName)) { $RestorePolicyName } else { $null }
                $action = if ($null -eq $targetPolicy) { 'Clear retention policy back to default' } else { "Restore retention policy '$targetPolicy'" }

                if (-not $PSCmdlet.ShouldProcess($mailboxName, $action)) {
                    continue
                }

                try {
                    Set-Mailbox -Identity $mailbox.Identity -RetentionPolicy $targetPolicy -ErrorAction Stop | Out-Null
                    $restoreTargets.Add([pscustomobject]@{
                            Mailbox        = $mailboxName
                            PreviousPolicy = [string]$mailbox.RetentionPolicy
                            NewPolicy      = $targetPolicy
                            Action         = if ($null -eq $targetPolicy) { 'Cleared' } else { 'Restored' }
                        }) | Out-Null
                }
                catch {
                    Write-NCMessage "Unable to restore retention policy for '$mailboxName'. $($_.Exception.Message)" -Level ERROR
                }
            }

            foreach ($policyName in $policyNamesToRemove) {
                $policy = $policyObjects | Where-Object { [string]$_.Name -ieq $policyName } | Select-Object -First 1
                if (-not $policy) {
                    continue
                }

                $linkedTags = @($policy.RetentionPolicyTagLinks)
                if ($linkedTags.Count -gt 1) {
                    Write-NCMessage ("Policy '{0}' has multiple tag links. Review it manually before deletion." -f $policyName) -Level WARNING
                    continue
                }

                if ($RemovePolicy -and $PSCmdlet.ShouldProcess($policyName, "Remove retention policy")) {
                    try {
                        Remove-RetentionPolicy -Identity $policyName -ErrorAction Stop
                        Write-NCMessage "Retention policy '$policyName' removed." -Level SUCCESS
                    }
                    catch {
                        Write-NCMessage "Unable to remove retention policy '$policyName'. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }

            if ($RemoveTag) {
                foreach ($tagName in $targetTagNames) {
                    $tag = $tagObjects | Where-Object { [string]$_.Name -ieq $tagName } | Select-Object -First 1
                    if (-not $tag) {
                        continue
                    }

                    $linkedPolicies = @($policyObjects | Where-Object { @($_.RetentionPolicyTagLinks) -contains $tagName })
                    if ($linkedPolicies.Count -gt 0) {
                        foreach ($policy in $linkedPolicies) {
                            if ($policyNamesToRemove.Contains([string]$policy.Name)) {
                                continue
                            }

                            Write-NCMessage ("Tag '{0}' is still linked to policy '{1}'. Review that policy before deleting the tag." -f $tagName, $policy.Name) -Level WARNING
                        }

                        if (($linkedPolicies | Where-Object { -not $policyNamesToRemove.Contains([string]$_.Name) }).Count -gt 0) {
                            continue
                        }
                    }

                    if ($PSCmdlet.ShouldProcess($tagName, "Remove retention policy tag")) {
                        try {
                            Remove-RetentionPolicyTag -Identity $tagName -ErrorAction Stop
                            Write-NCMessage "Retention policy tag '$tagName' removed." -Level SUCCESS
                        }
                        catch {
                            Write-NCMessage "Unable to remove retention policy tag '$tagName'. $($_.Exception.Message)" -Level ERROR
                        }
                    }
                }
            }

            if ($restoreTargets.Count -gt 0) {
                $restoreCount = $restoreTargets.Count
                Write-NCMessage ("Restored retention policy for {0} mailbox(es)." -f $restoreCount) -Level SUCCESS
            }

            if ($restoreTargets.Count -gt 0 -or $policyNamesToRemove.Count -gt 0 -or $targetTagNames.Count -gt 0) {
                [pscustomobject]@{
                    RestoredMailboxes = @($restoreTargets)
                    RemovedPolicies   = @($policyNamesToRemove | Sort-Object)
                    RemovedTags       = @($targetTagNames | Sort-Object)
                    RestorePolicyName = $RestorePolicyName
                    ClearedToDefault  = $ClearRetentionPolicy.IsPresent -or [string]::IsNullOrWhiteSpace($RestorePolicyName)
                }
            }
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

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

        $existingMailboxRetentionPolicy = $null
        try {
            $mailboxState = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
            $existingMailboxRetentionPolicy = [string]$mailboxState.RetentionPolicy
        }
        catch {
            Write-NCMessage "Unable to read current retention policy for '$Mailbox'. Cleanup hints will assume the default policy." -Level WARNING
        }

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

        $escapedPolicyName = $PolicyName -replace "'", "''"
        $escapedTagName = $TagName -replace "'", "''"
        $escapedExistingPolicy = if ([string]::IsNullOrWhiteSpace($existingMailboxRetentionPolicy)) {
            $null
        }
        else {
            $existingMailboxRetentionPolicy -replace "'", "''"
        }

        [pscustomobject]@{
            Mailbox            = $Mailbox
            FixedCutoffDate    = $FixedCutoffDate
            SafetyBufferDays   = $SafetyBufferDays
            AgeLimitDays       = $ageDays
            RetentionAction    = $RetentionAction
            TagName            = $TagName
            PolicyName         = $PolicyName
            ExistingPolicy     = $existingMailboxRetentionPolicy
            RollbackCommand    = if ([string]::IsNullOrWhiteSpace($existingMailboxRetentionPolicy)) {
                "Set-Mailbox -Identity '$Mailbox' -RetentionPolicy `$null"
            }
            else {
                "Set-Mailbox -Identity '$Mailbox' -RetentionPolicy '$escapedExistingPolicy'"
            }
            RemovePolicyHint   = "Remove-RetentionPolicy -Identity '$escapedPolicyName'"
            RemoveTagHint      = "Remove-RetentionPolicyTag -Identity '$escapedTagName'"
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
    }
}
