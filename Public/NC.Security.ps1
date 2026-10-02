#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Security helpers =====================================================================================================================

function Disable-UserDevices {
    <#
    .SYNOPSIS
        Disables all registered devices for specified users.
    .DESCRIPTION
        Ensures Microsoft Graph connectivity, resolves the target users, and sets AccountEnabled = $false
        on each registered device. Skips missing users and reports progress.
    .PARAMETER UserPrincipalName
        One or more user principal names. Accepts pipeline input.
    .PARAMETER PassThru
        Emit the impacted devices as objects.
    .EXAMPLE
        Disable-UserDevices -UserPrincipalName user1@contoso.com,user2@contoso.com -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $targets.Add($upn.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if ($targets.Count -eq 0) {
                Write-NCMessage "No user principal names provided." -Level WARNING
                return
            }

            $scopes = @('Directory.ReadWrite.All', 'Device.ReadWrite.All')
            if (-not (Test-MgGraphConnection -Scopes $scopes -EnsureExchangeOnline:$false)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $queue = @(foreach ($entry in $targets) { if ($dedup.Add($entry)) { $entry } })

            Write-NCGraphBatchNotice -Count ($queue.Count) -Noun 'user(s)'

            for ($offset = 0; $offset -lt $queue.Count; $offset += 20) {
                $chunk = @($queue[$offset..([Math]::Min($offset + 20, $queue.Count) - 1)])
                $Percentage = Get-NCProgressPercent -Current $offset -Total $queue.Count
                Write-Progress -Activity "Resolving users" -Status "$offset of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                # (a) Resolve the whole chunk.
                $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $chunk -Property @('id', 'userPrincipalName', 'displayName') -FailedIdentifier $failedUsers

                $users = [System.Collections.Generic.List[object]]::new()
                foreach ($upn in $chunk) {
                    $user = $resolvedUsers[$upn.Trim()]
                    if (-not $user -or -not $user.id) {
                        if (-not $failedUsers.Contains($upn.Trim())) {
                            Write-NCMessage "Can't find Azure AD account for user $upn." -Level ERROR
                        }
                        continue
                    }
                    $users.Add($user) | Out-Null
                }
                if ($users.Count -eq 0) { continue }

                # (b) Read the registered devices of every resolved user.
                $deviceRequests = @(for ($i = 0; $i -lt $users.Count; $i++) {
                        @{
                            Id     = "m$i"
                            Method = 'GET'
                            Url    = "/users/$([uri]::EscapeDataString([string]$users[$i].id))/registeredDevices?`$select=id,displayName,accountEnabled"
                        }
                    })
                $deviceResponses = @(Invoke-NCGraphBatchCollection -Requests $deviceRequests -Activity 'Reading registered devices')

                # (c) Confirm per device.
                $approved = [System.Collections.Generic.List[object]]::new()
                for ($i = 0; $i -lt $users.Count; $i++) {
                    $user = $users[$i]
                    $response = $deviceResponses[$i]
                    if (-not $response.Success) {
                        Write-NCMessage "Unable to retrieve registered devices for $($user.userPrincipalName). $($response.ErrorMessage)" -Level ERROR
                        continue
                    }

                    $devices = @($response.Items)
                    if ($devices.Count -eq 0) {
                        Write-NCMessage ("No registered devices found for {0}." -f $user.userPrincipalName) -Level WARNING
                        continue
                    }

                    foreach ($device in $devices) {
                        $deviceLabel = if ($device.displayName) { $device.displayName } else { $device.id }
                        if ($PSCmdlet.ShouldProcess($deviceLabel, "Disable device for user $($user.userPrincipalName)")) {
                            $approved.Add([pscustomobject]@{ User = $user; Device = $device; Label = $deviceLabel }) | Out-Null
                        }
                    }
                }
                if ($approved.Count -eq 0) { continue }

                # (d) Disable the approved devices in batches.
                $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                        @{
                            Id     = "m$i"
                            Method = 'PATCH'
                            Url    = "/devices/$([uri]::EscapeDataString([string]$approved[$i].Device.id))"
                            Body   = @{ accountEnabled = $false }
                        }
                    })
                $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity 'Disabling devices')
                for ($i = 0; $i -lt $approved.Count; $i++) {
                    $user = $approved[$i].User
                    $device = $approved[$i].Device
                    $response = $responses[$i]
                    if ($response.Success) {
                        $results.Add([pscustomobject]@{
                                UserPrincipalName = $user.userPrincipalName
                                UserDisplayName   = $user.displayName
                                DeviceId          = $device.id
                                DeviceDisplayName = $device.displayName
                                Action            = 'Disabled'
                            }) | Out-Null
                    }
                    else {
                        Write-NCMessage "Failed to disable device $($approved[$i].Label) for $($user.userPrincipalName). $($response.ErrorMessage)" -Level ERROR
                    }
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($results.Count -gt 0) {
                $deviceLabel = if ($results.Count -eq 1) { 'device' } else { 'devices' }
                Write-NCMessage ("Disabled {0} {1}." -f $results.Count, $deviceLabel) -Level SUCCESS
            }
        }
        finally {
            Write-Progress -Activity "Resolving user" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Disable-UserSignIn {
    <#
    .SYNOPSIS
        Blocks sign-in for specified users.
    .DESCRIPTION
        Ensures Microsoft Graph connectivity, resolves the target users, and sets AccountEnabled = $false.
        Skips missing users and supports WhatIf/Confirm.
    .PARAMETER UserPrincipalName
        One or more user principal names. Accepts pipeline input.
    .PARAMETER PassThru
        Emit the impacted users as objects.
    .EXAMPLE
        Disable-UserSignIn -UserPrincipalName user1@contoso.com,user2@contoso.com -Confirm:$false
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $targets.Add($upn.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if ($targets.Count -eq 0) {
                Write-NCMessage "No user principal names provided." -Level WARNING
                return
            }

            $scopes = @('Directory.ReadWrite.All')
            if (-not (Test-MgGraphConnection -Scopes $scopes -EnsureExchangeOnline:$false)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $queue = @(foreach ($entry in $targets) { if ($dedup.Add($entry)) { $entry } })

            Write-NCGraphBatchNotice -Count ($queue.Count) -Noun 'user(s)'

            for ($offset = 0; $offset -lt $queue.Count; $offset += 20) {
                $chunk = @($queue[$offset..([Math]::Min($offset + 20, $queue.Count) - 1)])
                $Percentage = Get-NCProgressPercent -Current $offset -Total $queue.Count
                Write-Progress -Activity "Processing users" -Status "$offset of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                # (a) Resolve the whole chunk.
                $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $chunk -Property @('id', 'userPrincipalName', 'displayName') -FailedIdentifier $failedUsers

                # (b) Confirm per user.
                $approved = [System.Collections.Generic.List[object]]::new()
                foreach ($upn in $chunk) {
                    $user = $resolvedUsers[$upn.Trim()]
                    if (-not $user -or -not $user.id) {
                        if (-not $failedUsers.Contains($upn.Trim())) {
                            Write-NCMessage "Can't find Azure AD account for user $upn." -Level ERROR
                        }
                        continue
                    }

                    if ($PSCmdlet.ShouldProcess($user.userPrincipalName, "Disable sign-in")) {
                        $approved.Add($user) | Out-Null
                    }
                }
                if ($approved.Count -eq 0) { continue }

                # (c) Disable the approved users in one batch.
                $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                        @{
                            Id     = "m$i"
                            Method = 'PATCH'
                            Url    = "/users/$([uri]::EscapeDataString([string]$approved[$i].id))"
                            Body   = @{ accountEnabled = $false }
                        }
                    })
                $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity 'Disabling sign-in')
                for ($i = 0; $i -lt $approved.Count; $i++) {
                    $user = $approved[$i]
                    $response = $responses[$i]
                    if ($response.Success) {
                        $results.Add([pscustomobject]@{
                                UserPrincipalName = $user.userPrincipalName
                                DisplayName       = $user.displayName
                                Action            = 'SignInDisabled'
                            }) | Out-Null
                    }
                    else {
                        Write-NCMessage "Failed to disable sign-in for $($user.userPrincipalName). $($response.ErrorMessage)" -Level ERROR
                    }
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($results.Count -gt 0) {
                $userLabel = if ($results.Count -eq 1) { 'user' } else { 'users' }
                Write-NCMessage ("Sign-in disabled for {0} {1}." -f $results.Count, $userLabel) -Level SUCCESS
            }
        }
        finally {
            Write-Progress -Activity "Processing users" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Get-ContentFilterPolicy {
    <#
    .SYNOPSIS
        Reads hosted content filter policy configuration.
    .DESCRIPTION
        Connects to Exchange Online and returns one or more hosted content filter policies. If no
        policy name is provided, the function lists all available policies. The output includes the
        current allow/block lists so you can inspect how each policy is configured before editing it.
    .PARAMETER Identity
        Hosted content filter policy name. If omitted, all policies are returned.
    .PARAMETER Detailed
        Include the resolved allow/block lists in the output.
    .EXAMPLE
        Get-ContentFilterPolicy
    .EXAMPLE
        Get-ContentFilterPolicy -Identity Contoso
    .EXAMPLE
        Get-ContentFilterPolicy -Detailed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('SpamFilter', 'PolicyName')]
        [string[]]$Identity,
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

            $policyObjects = [System.Collections.Generic.List[object]]::new()

            if ($targets.Count -eq 0) {
                try {
                    $policyObjects.AddRange(@(Get-HostedContentFilterPolicy -ErrorAction Stop))
                }
                catch {
                    Write-NCMessage "Unable to retrieve hosted content filter policies. $($_.Exception.Message)" -Level ERROR
                    return
                }

                if ($policyObjects.Count -eq 0) {
                    Write-NCMessage "No hosted content filter policies were found." -Level WARNING
                    return
                }
            }
            else {
                $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($policyName in $targets) {
                    if (-not $dedup.Add($policyName)) {
                        continue
                    }

                    try {
                        $policyObjects.Add((Get-HostedContentFilterPolicy -Identity $policyName -ErrorAction Stop)) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Unable to read hosted content filter policy '$policyName'. $($_.Exception.Message)" -Level ERROR
                    }
                }

                if ($policyObjects.Count -eq 0) {
                    Write-NCMessage "No hosted content filter policies matched the requested identity or identities." -Level WARNING
                    return
                }
            }
            $policyObjects |
                Sort-Object Name |
                ForEach-Object {
                    $blockedSenders = @(Get-NCContentFilterPolicyValues -PolicyObject $_ -PropertyName 'BlockedSenders' -PreferredValueProperty 'Sender')
                    $blockedSenderDomains = @(Get-NCContentFilterPolicyValues -PolicyObject $_ -PropertyName 'BlockedSenderDomains' -PreferredValueProperty 'Domain')
                    $allowedSenders = @(Get-NCContentFilterPolicyValues -PolicyObject $_ -PropertyName 'AllowedSenders' -PreferredValueProperty 'Sender')
                    $allowedSenderDomains = @(Get-NCContentFilterPolicyValues -PolicyObject $_ -PropertyName 'AllowedSenderDomains' -PreferredValueProperty 'Domain')

                    $summary = [ordered]@{
                        Identity            = $_.Identity
                        Name                = $_.Name
                        Enabled             = $_.Enabled
                        Priority            = $_.Priority
                        BlockedSenderCount  = $blockedSenders.Count
                        BlockedDomainCount  = $blockedSenderDomains.Count
                        AllowedSenderCount  = $allowedSenders.Count
                        AllowedDomainCount  = $allowedSenderDomains.Count
                    }

                    if ($Detailed.IsPresent) {
                        $summary.BlockedSenders = $blockedSenders
                        $summary.BlockedSenderDomains = $blockedSenderDomains
                        $summary.AllowedSenders = $allowedSenders
                        $summary.AllowedSenderDomains = $allowedSenderDomains
                    }

                    [pscustomobject]$summary
                }
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Edit-ContentFilterPolicy {
    <#
    .SYNOPSIS
        Updates a hosted content filter policy allow/block list.
    .DESCRIPTION
        Connects to Exchange Online, loads a hosted content filter policy, and adds or removes
        blocked senders, blocked domains, allowed senders, or allowed domains. When allowed senders
        are updated, the helper also keeps the configured allowed-senders group in sync and creates
        missing mail contacts. When allowed domains are updated, the helper also synchronizes the
        sender-domain exceptions on the configured transport rules.
    .PARAMETER Identity
        Hosted content filter policy name. Accepts the legacy SpamFilter alias.
    .PARAMETER BlockedSender
        Sender address to add or remove from the blocked senders list.
    .PARAMETER BlockedDomain
        Domain to add or remove from the blocked sender domains list.
    .PARAMETER AllowedSender
        Sender address to add or remove from the allowed senders list.
    .PARAMETER AllowedDomain
        Domain to add or remove from the allowed sender domains list.
    .PARAMETER AllowedSendersGroup
        Distribution group used to keep the allowed senders group in sync.
    .PARAMETER TransportRuleNames
        Transport rules that should mirror allowed-domain exceptions.
    .PARAMETER Remove
        Remove the provided values instead of adding them.
    .EXAMPLE
        Edit-ContentFilterPolicy -Identity Contoso -BlockedSender user@contoso.com
    .EXAMPLE
        Edit-ContentFilterPolicy -Identity Contoso -AllowedDomain contoso.com -Remove
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('SpamFilter', 'PolicyName')]
        [string]$Identity,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$BlockedSender,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$BlockedDomain,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$AllowedSender,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$AllowedDomain,
        [string]$AllowedSendersGroup,
        [string[]]$TransportRuleNames,
        [switch]$Remove
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

        $blockedSendersQueue = @(
            foreach ($entry in @($BlockedSender)) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    $entry.Trim()
                }
            }
        ) | Sort-Object -Unique

        $blockedDomainsQueue = @(
            foreach ($entry in @($BlockedDomain)) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    $entry.Trim()
                }
            }
        ) | Sort-Object -Unique

        $allowedSendersQueue = @(
            foreach ($entry in @($AllowedSender)) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    $entry.Trim()
                }
            }
        ) | Sort-Object -Unique

        $allowedDomainsQueue = @(
            foreach ($entry in @($AllowedDomain)) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    $entry.Trim()
                }
            }
        ) | Sort-Object -Unique

        $modeLabel = if ($Remove.IsPresent) { 'Remove' } else { 'Add' }

        if ($blockedSendersQueue.Count -eq 0 -and $blockedDomainsQueue.Count -eq 0 -and $allowedSendersQueue.Count -eq 0 -and $allowedDomainsQueue.Count -eq 0) {
            Write-NCMessage "No changes requested. Returning the current policy state." -Level INFO
        }

        try {
            $policy = Get-HostedContentFilterPolicy -Identity $Identity -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to read hosted content filter policy '$Identity'. $($_.Exception.Message)" -Level ERROR
            return
        }

        $transportRules = [System.Collections.Generic.List[object]]::new()
        foreach ($ruleName in @($TransportRuleNames)) {
            if ([string]::IsNullOrWhiteSpace($ruleName)) {
                continue
            }

            try {
                $rule = Get-TransportRule -Identity $ruleName -ErrorAction Stop
                $transportRules.Add($rule) | Out-Null
            }
            catch {
                Write-NCMessage "Transport rule '$ruleName' was not found or could not be read. $($_.Exception.Message)" -Level WARNING
            }
        }

        $contactGroup = $null
        if (-not [string]::IsNullOrWhiteSpace($AllowedSendersGroup)) {
            try {
                $contactGroup = Get-DistributionGroup -Identity $AllowedSendersGroup -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Allowed senders group '$AllowedSendersGroup' was not found or could not be read. $($_.Exception.Message)" -Level WARNING
            }
        }

        if ($allowedSendersQueue.Count -gt 0 -and -not $contactGroup) {
            Write-NCMessage "Allowed sender updates will skip group synchronization because -AllowedSendersGroup was not provided or could not be resolved." -Level INFO
        }

        if ($allowedDomainsQueue.Count -gt 0 -and $transportRules.Count -eq 0) {
            Write-NCMessage "Allowed domain updates will skip transport-rule synchronization because -TransportRuleNames was not provided or no rules could be resolved." -Level INFO
        }

        $changes = [System.Collections.Generic.List[object]]::new()
        $updateDomainList = {
            param(
                [object]$CurrentValues,
                [string]$Entry,
                [switch]$Remove
            )

            $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($value in @($CurrentValues)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    [void]$set.Add(([string]$value).Trim())
                }
            }

            if ($Remove) {
                [void]$set.Remove($Entry)
            }
            else {
                [void]$set.Add($Entry)
            }

            return @($set)
        }

        Write-NCMessage "Policy: $Identity" -Level INFO
        Write-NCMessage ("Mode: {0}" -f $modeLabel) -Level INFO

        foreach ($entry in $blockedSendersQueue) {
            $actionText = if ($Remove.IsPresent) { "remove blocked sender '$entry'" } else { "add blocked sender '$entry'" }
            if (-not $PSCmdlet.ShouldProcess($Identity, $actionText)) {
                continue
            }

            try {
                if ($Remove.IsPresent) {
                    Set-HostedContentFilterPolicy -Identity $Identity -BlockedSenders @{ remove = $entry } -ErrorAction Stop | Out-Null
                }
                else {
                    Set-HostedContentFilterPolicy -Identity $Identity -BlockedSenders @{ add = $entry } -ErrorAction Stop | Out-Null
                }

                $changes.Add([pscustomobject]@{
                        Scope  = 'BlockedSenders'
                        Value  = $entry
                        Action = $(if ($Remove.IsPresent) { 'Removed' } else { 'Added' })
                    }) | Out-Null
            }
            catch {
                Write-NCMessage "Unable to update blocked sender '$entry' on '$Identity'. $($_.Exception.Message)" -Level ERROR
            }
        }

        foreach ($entry in $blockedDomainsQueue) {
            $actionText = if ($Remove.IsPresent) { "remove blocked domain '$entry'" } else { "add blocked domain '$entry'" }
            if (-not $PSCmdlet.ShouldProcess($Identity, $actionText)) {
                continue
            }

            try {
                if ($Remove.IsPresent) {
                    Set-HostedContentFilterPolicy -Identity $Identity -BlockedSenderDomains @{ remove = $entry } -ErrorAction Stop | Out-Null
                }
                else {
                    Set-HostedContentFilterPolicy -Identity $Identity -BlockedSenderDomains @{ add = $entry } -ErrorAction Stop | Out-Null
                }

                $changes.Add([pscustomobject]@{
                        Scope  = 'BlockedSenderDomains'
                        Value  = $entry
                        Action = $(if ($Remove.IsPresent) { 'Removed' } else { 'Added' })
                    }) | Out-Null
            }
            catch {
                Write-NCMessage "Unable to update blocked sender domain '$entry' on '$Identity'. $($_.Exception.Message)" -Level ERROR
            }
        }

        foreach ($entry in $allowedSendersQueue) {
            $actionText = if ($Remove.IsPresent) { "remove allowed sender '$entry'" } else { "add allowed sender '$entry'" }
            if (-not $PSCmdlet.ShouldProcess($Identity, $actionText)) {
                continue
            }

            try {
                if (-not $Remove.IsPresent) {
                    $contact = Get-MailContact -Identity $entry -ErrorAction SilentlyContinue
                    if (-not $contact) {
                        New-MailContact -DisplayName $entry -Name $entry -ExternalEmailAddress $entry -ErrorAction Stop | Out-Null
                        Write-NCMessage "Created mail contact for '$entry'." -Level SUCCESS
                    }

                    Set-MailContact -Identity $entry -HiddenFromAddressListsEnabled $true -ErrorAction Stop | Out-Null

                    if ($contactGroup) {
                        try {
                            Add-DistributionGroupMember -Identity $contactGroup.Identity -Member $entry -ErrorAction Stop | Out-Null
                        }
                        catch {
                            Write-NCMessage "Unable to add '$entry' to allowed senders group '$AllowedSendersGroup'. $($_.Exception.Message)" -Level WARNING
                        }
                    }
                }
                else {
                    if ($contactGroup) {
                        try {
                            Remove-DistributionGroupMember -Identity $contactGroup.Identity -Member $entry -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                        catch {
                            Write-NCMessage "Unable to remove '$entry' from allowed senders group '$AllowedSendersGroup'. $($_.Exception.Message)" -Level WARNING
                        }
                    }
                }

                if ($Remove.IsPresent) {
                    Set-HostedContentFilterPolicy -Identity $Identity -AllowedSenders @{ remove = $entry } -ErrorAction Stop | Out-Null
                }
                else {
                    Set-HostedContentFilterPolicy -Identity $Identity -AllowedSenders @{ add = $entry } -ErrorAction Stop | Out-Null
                }

                $changes.Add([pscustomobject]@{
                        Scope  = 'AllowedSenders'
                        Value  = $entry
                        Action = $(if ($Remove.IsPresent) { 'Removed' } else { 'Added' })
                    }) | Out-Null
            }
            catch {
                Write-NCMessage "Unable to update allowed sender '$entry' on '$Identity'. $($_.Exception.Message)" -Level ERROR
            }
        }

        foreach ($entry in $allowedDomainsQueue) {
            $actionText = if ($Remove.IsPresent) { "remove allowed domain '$entry'" } else { "add allowed domain '$entry'" }
            if (-not $PSCmdlet.ShouldProcess($Identity, $actionText)) {
                continue
            }

            try {
                if ($Remove.IsPresent) {
                    Set-HostedContentFilterPolicy -Identity $Identity -AllowedSenderDomains @{ remove = $entry } -ErrorAction Stop | Out-Null
                }
                else {
                    Set-HostedContentFilterPolicy -Identity $Identity -AllowedSenderDomains @{ add = $entry } -ErrorAction Stop | Out-Null
                }

                foreach ($rule in $transportRules) {
                    $currentDomains = @($rule.ExceptIfSenderDomainIs)
                    $updatedDomains = & $updateDomainList $currentDomains $entry -Remove:$Remove.IsPresent
                    $ruleActionText = if ($Remove.IsPresent) {
                        "remove sender-domain exception '$entry' from transport rule '$($rule.Name)'"
                    }
                    else {
                        "add sender-domain exception '$entry' to transport rule '$($rule.Name)'"
                    }

                    if (-not $PSCmdlet.ShouldProcess($rule.Name, $ruleActionText)) {
                        continue
                    }

                    try {
                        if ($updatedDomains.Count -gt 0) {
                            Set-TransportRule -Identity $rule.Identity -ExceptIfSenderDomainIs $updatedDomains -ErrorAction Stop | Out-Null
                        }
                        else {
                            Set-TransportRule -Identity $rule.Identity -ExceptIfSenderDomainIs $null -ErrorAction Stop | Out-Null
                        }
                    }
                    catch {
                        Write-NCMessage "Unable to update transport rule '$($rule.Name)' for domain '$entry'. $($_.Exception.Message)" -Level WARNING
                    }
                }

                $changes.Add([pscustomobject]@{
                        Scope  = 'AllowedSenderDomains'
                        Value  = $entry
                        Action = $(if ($Remove.IsPresent) { 'Removed' } else { 'Added' })
                    }) | Out-Null
            }
            catch {
                Write-NCMessage "Unable to update allowed sender domain '$entry' on '$Identity'. $($_.Exception.Message)" -Level ERROR
            }
        }

        try {
            $updatedPolicy = Get-HostedContentFilterPolicy -Identity $Identity -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Policy '$Identity' was updated, but the refreshed policy could not be read back. $($_.Exception.Message)" -Level WARNING
            $updatedPolicy = $policy
        }

        $summary = [pscustomobject]@{
            Identity             = $updatedPolicy.Identity
            DisplayName           = $updatedPolicy.Name
            AllowedSenders        = @(Get-NCContentFilterPolicyValues -PolicyObject $updatedPolicy -PropertyName 'AllowedSenders' -PreferredValueProperty 'Sender')
            AllowedSenderDomains  = @(Get-NCContentFilterPolicyValues -PolicyObject $updatedPolicy -PropertyName 'AllowedSenderDomains' -PreferredValueProperty 'Domain')
            BlockedSenders        = @(Get-NCContentFilterPolicyValues -PolicyObject $updatedPolicy -PropertyName 'BlockedSenders' -PreferredValueProperty 'Sender')
            BlockedSenderDomains  = @(Get-NCContentFilterPolicyValues -PolicyObject $updatedPolicy -PropertyName 'BlockedSenderDomains' -PreferredValueProperty 'Domain')
            AllowedSendersGroup   = $AllowedSendersGroup
            TransportRuleNames    = @($transportRules | Select-Object -ExpandProperty Name)
            Changes               = @($changes)
        }

        if ($changes.Count -gt 0) {
            Write-NCMessage ("Updated content filter policy '{0}' with {1} change(s)." -f $Identity, $changes.Count) -Level SUCCESS
        }
        else {
            Write-NCMessage "No policy changes were applied." -Level INFO
        }

        return $summary
    }

    end {
        Restore-ProgressAndInfoPreferences
    }
}

function Revoke-UserSessions {
    <#
    .SYNOPSIS
        Forces sign-out by revoking refresh tokens for users.
    .DESCRIPTION
        Ensures Microsoft Graph connectivity, targets all users or a selection (with optional exclusions),
        and calls Revoke-MgUserSignInSession. Supports WhatIf/Confirm.
    .PARAMETER All
        Target every user in the tenant.
    .PARAMETER UserPrincipalName
        Specific users to target. Accepts pipeline input.
    .PARAMETER Exclude
        Users to skip when using -All or a list.
    .PARAMETER PassThru
        Emit the impacted users as objects.
    .EXAMPLE
        Revoke-UserSessions -UserPrincipalName user1@contoso.com,user2@contoso.com
    .EXAMPLE
        Revoke-UserSessions -All -Exclude user@contoso.com -Confirm:$false
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [switch]$All,
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [string[]]$Exclude,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
        $exclusions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $targets.Add($upn.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if (-not $All.IsPresent -and $targets.Count -eq 0) {
                Write-NCMessage "No target specified. Use -All or provide user principal names." -Level WARNING
                return
            }

            foreach ($ex in $Exclude) {
                if (-not [string]::IsNullOrWhiteSpace($ex)) { $exclusions.Add($ex.Trim()) | Out-Null }
            }

            $scopes = @('Directory.ReadWrite.All')
            if (-not (Test-MgGraphConnection -Scopes $scopes -EnsureExchangeOnline:$false)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            $queue = [System.Collections.Generic.List[object]]::new()

            if ($All.IsPresent) {
                try {
                    $allUsers = Get-MgUser -All -ConsistencyLevel eventual -ErrorAction Stop
                    foreach ($u in $allUsers) {
                        $queue.Add([pscustomobject]@{ id = $u.Id; userPrincipalName = $u.UserPrincipalName; displayName = $u.DisplayName }) | Out-Null
                    }
                }
                catch {
                    Write-NCMessage "Unable to retrieve all users. $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            else {
                $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $uniqueTargets = @(foreach ($entry in $targets) { if ($dedup.Add($entry)) { $entry } })

                Write-NCGraphBatchNotice -Count ($uniqueTargets.Count) -Noun 'user(s)'

                for ($offset = 0; $offset -lt $uniqueTargets.Count; $offset += 20) {
                    $chunk = @($uniqueTargets[$offset..([Math]::Min($offset + 20, $uniqueTargets.Count) - 1)])
                    $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $chunk -Property @('id', 'userPrincipalName', 'displayName') -FailedIdentifier $failedUsers
                    foreach ($upn in $chunk) {
                        $user = $resolvedUsers[$upn.Trim()]
                        if ($user -and $user.id) {
                            $queue.Add($user) | Out-Null
                        }
                        elseif (-not $failedUsers.Contains($upn.Trim())) {
                            Write-NCMessage "Can't find Azure AD account for user $upn." -Level ERROR
                        }
                    }
                }
            }

            if ($queue.Count -eq 0) {
                Write-NCMessage "No users to process." -Level WARNING
                return
            }

            if ($All.IsPresent) {
                Write-NCGraphBatchNotice -Count ($queue.Count) -Noun 'user(s)'
            }

            $results = [System.Collections.Generic.List[object]]::new()

            for ($offset = 0; $offset -lt $queue.Count; $offset += 20) {
                $chunk = @($queue[$offset..([Math]::Min($offset + 20, $queue.Count) - 1)])
                $Percentage = Get-NCProgressPercent -Current $offset -Total $queue.Count
                Write-Progress -Activity "Revoking sessions" -Status "$offset of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                $approved = [System.Collections.Generic.List[object]]::new()
                foreach ($user in $chunk) {
                    if ($exclusions.Contains([string]$user.userPrincipalName)) {
                        Write-NCMessage ("Skipping user {0}" -f $user.userPrincipalName) -Level INFO
                        continue
                    }

                    if ($PSCmdlet.ShouldProcess($user.userPrincipalName, "Revoke sign-in sessions")) {
                        $approved.Add($user) | Out-Null
                    }
                }
                if ($approved.Count -eq 0) { continue }

                $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                        @{
                            Id     = "m$i"
                            Method = 'POST'
                            Url    = "/users/$([uri]::EscapeDataString([string]$approved[$i].id))/revokeSignInSessions"
                            Body   = @{}
                        }
                    })
                $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity 'Revoking sessions')
                for ($i = 0; $i -lt $approved.Count; $i++) {
                    $user = $approved[$i]
                    $response = $responses[$i]
                    if ($response.Success) {
                        $results.Add([pscustomobject]@{
                                UserPrincipalName = $user.userPrincipalName
                                DisplayName       = $user.displayName
                                Action            = 'SessionsRevoked'
                            }) | Out-Null
                    }
                    else {
                        Write-NCMessage "Failed to revoke sessions for $($user.userPrincipalName). $($response.ErrorMessage)" -Level ERROR
                    }
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($results.Count -gt 0) {
                $userLabel = if ($results.Count -eq 1) { 'user' } else { 'users' }
                Write-NCMessage ("Revoked sessions for {0} {1}." -f $results.Count, $userLabel) -Level SUCCESS
            }
        }
        finally {
            Write-Progress -Activity "Revoking sessions" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}



