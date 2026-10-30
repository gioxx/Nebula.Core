#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Licenses helpers =====================================================================================================================

function Add-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Assigns licenses to a user by friendly name or SKU identifier.
    .DESCRIPTION
        Resolves the provided license names using the cached license catalog and the tenant's subscribed SKUs,
        then assigns them to the target user (preserving existing licenses). Accepts friendly product names,
        SKU part numbers, or SKU IDs.
    .PARAMETER UserPrincipalName
        Target user UPN or object ID.
    .PARAMETER License
        One or more license identifiers: friendly name (as resolved by the catalog), SKU part number, or SKU ID (GUID).
    .PARAMETER ForceLicenseCatalogRefresh
        Force a refresh of the cached license catalog before resolving friendly names.
    .PARAMETER ShowErrorDetails
        Kept for compatibility. Error messages always include the Microsoft Graph error detail, so this switch has no effect.
    .EXAMPLE
        Add-MsolAccountSku -UserPrincipalName user@contoso.com -License "Microsoft 365 E3"
    .EXAMPLE
        Add-MsolAccountSku -UserPrincipalName user@contoso.com -License "ENTERPRISEPACK","VISIOCLIENT"
    .EXAMPLE
        Add-MsolAccountSku -UserPrincipalName user@contoso.com -License "18181a46-0d4e-45cd-891e-60aabd171b4e"
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN')]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string[]]$License,
        [switch]$ForceLicenseCatalogRefresh,
        [switch]$ShowErrorDetails
    )

    begin {
        Set-ProgressAndInfoPreferences

        $graphConnected = Test-MgGraphConnection
        $ready = $false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $queue = [System.Collections.Generic.List[object]]::new()
        $state = @{ Started = $false }
        # The same user can arrive under different identifiers, in any batch: count it, and its seats, once
        $seenUserIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $maxAttempts = 3

        $defaultUsageLocation = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) {
            [string]$NCVars.UsageLocation
        }
        else { 'US' }

        $normalizeUsageLocation = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return $value.Trim().ToUpperInvariant()
        }

        $normalizeString = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return ($value.Trim().ToUpperInvariant())
        }

        $uniqueAdds = @()
        $resolved = @()
        $requestedList = ''
        $remaining = @{}

        if ($graphConnected) {
            # -License is not pipeline-bound: tenant data and license matching are item-independent.
            try {
                $licenseCatalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent
            }
            catch {
                Write-NCMessage $_ -Level WARNING
                $licenseCatalog = $null
            }

            $licenseLookup = $null
            $customLookup = $null
            if ($licenseCatalog) {
                if ($licenseCatalog.PSObject.Properties['Lookup']) { $licenseLookup = $licenseCatalog.Lookup }
                if ($licenseCatalog.PSObject.Properties['CustomLookup']) { $customLookup = $licenseCatalog.CustomLookup }
            }

            $tenantSkus = $null
            $tenantSkusFailed = $false
            try {
                $tenantSkus = Invoke-NCRetry -Action {
                    Get-MgSubscribedSku -All -ErrorAction Stop
                } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve tenant licenses" -OnError {
                    param($attempt, $max, $err)
                    $currentAttempt = if ($attempt) { $attempt } else { '?' }
                    $currentMax = if ($max) { $max } else { $maxAttempts }
                    Write-NCMessage "Failed to retrieve tenant licenses, attempt $currentAttempt of $currentMax." -Level ERROR
                }
            }
            catch {
                Write-NCMessage "Unable to retrieve tenant licenses after $maxAttempts attempts." -Level ERROR
                $tenantSkusFailed = $true
            }

            if (-not $tenantSkusFailed) {
                if (-not $tenantSkus -or @($tenantSkus).Count -eq 0) {
                    Write-NCMessage "No tenant licenses available to assign." -Level WARNING
                }
                else {
                    $unmatched = @()
                    $inputLicenses = $License | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique

                    foreach ($entry in $inputLicenses) {
                        $target = & $normalizeString $entry
                        $match = $null

                        foreach ($sku in $tenantSkus) {
                            $skuIdString = [string]$sku.SkuId
                            $skuPart = & $normalizeString $sku.SkuPartNumber

                            $display = $null
                            $matchSource = $null
                            if ($licenseLookup) {
                                $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $sku.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
                            }
                            $displayNormalized = if ($display) { & $normalizeString $display } else { $null }

                            if ($target -eq $skuPart -or $target -eq ($skuIdString.ToUpperInvariant()) -or ($displayNormalized -and $target -eq $displayNormalized)) {
                                $prepaidUnits = $sku.PrepaidUnits
                                $enabledUnits = if ($prepaidUnits) { [int]$prepaidUnits.Enabled } else { 0 }
                                $consumedUnits = if ($sku.ConsumedUnits -is [int]) { [int]$sku.ConsumedUnits } else { [int]0 }
                                $availableUnits = [Math]::Max($enabledUnits - $consumedUnits, 0)
                                $match = @{
                                    SkuId         = $sku.SkuId
                                    SkuPartNumber = $sku.SkuPartNumber
                                    Name          = if ($display) { $display } else { $sku.SkuPartNumber }
                                    Available     = $availableUnits
                                }
                                break
                            }
                        }

                        if ($match) {
                            $resolved += $match
                        }
                        else {
                            $unmatched += $entry
                        }
                    }

                    if ($unmatched.Count -gt 0) {
                        Write-NCMessage ("Unable to resolve license(s): {0}" -f ($unmatched -join ', ')) -Level ERROR
                    }
                    else {
                        $uniqueAdds = @($resolved | Group-Object SkuId | ForEach-Object {
                                $_.Group | Select-Object -First 1
                            })
                        $requestedList = ($resolved | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '
                        # Local per-SKU counter seeded from the tenant snapshot; consumed as users are approved.
                        foreach ($item in $uniqueAdds) {
                            $remaining[[string]$item.SkuId] = [int]$item.Available
                        }
                        $ready = $true
                    }
                }
            }
        }

        # Resolves the queued users and assigns licenses in Graph batches.
        $flush = {
            $entries = @($queue)
            $queue.Clear()
            if ($entries.Count -eq 0) { return }

            if (-not $state.Started) {
                $state.Started = $true
                Write-NCGraphBatchNotice -Count ($entries.Count) -Noun 'users' -Streaming
            }

            # (a) Resolve every queued user.
            $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $entries -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation', 'assignedLicenses') -FailedIdentifier $failedUsers

            # (b) Check availability and ask for confirmation per user, in input order.
            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $entries) {
                $user = $resolvedUsers[$entry.Trim()]
                if (-not $user -or -not $user.id) {
                    if (-not $failedUsers.Contains($entry.Trim())) {
                        Write-NCMessage "Unable to resolve user recipient for $entry" -Level ERROR
                    }
                    continue
                }

                if (-not $seenUserIds.Add([string]$user.id)) {
                    Write-Verbose "Skipping $($entry): $($user.userPrincipalName) is already in this request."
                    continue
                }

                $upn = [string]$user.userPrincipalName
                $normalizedCurrentUsage = & $normalizeUsageLocation $user.usageLocation
                $normalizedTargetUsage = & $normalizeUsageLocation $defaultUsageLocation
                $targetUsage = if ($normalizedTargetUsage -and $normalizedTargetUsage -ne $normalizedCurrentUsage) { $defaultUsageLocation } else { $null }

                # Skip SKUs the user already has: assigning them again uses no seat and would reset their disabled plans
                $ownedSkuIds = @($user.assignedLicenses | ForEach-Object { [string]$_.skuId })
                $namesAlreadyAssigned = @($uniqueAdds | Where-Object { $ownedSkuIds -contains [string]$_.SkuId } | ForEach-Object { $_.Name })
                if ($namesAlreadyAssigned.Count -gt 0) {
                    Write-NCMessage ("{0} already has license(s): {1}" -f $upn, (($namesAlreadyAssigned | Select-Object -Unique) -join ', ')) -Level WARNING
                }

                $assignableItems = @()
                $namesNoAvailability = @()
                foreach ($item in $uniqueAdds) {
                    if ($ownedSkuIds -contains [string]$item.SkuId) { continue }
                    $available = $remaining[[string]$item.SkuId]
                    if ($available -le 0) {
                        Write-NCMessage ("No available units for license {0} ({1}) (available: {2})" -f $item.Name, $item.SkuPartNumber, $available) -Level WARNING
                        $namesNoAvailability += $item.Name
                        continue
                    }
                    $assignableItems += $item
                }

                if ($assignableItems.Count -eq 0) {
                    if ($namesNoAvailability.Count -gt 0) {
                        Write-NCMessage ("No licenses to assign: none available. Requested: {0}" -f $requestedList) -Level ERROR
                    }
                    continue
                }

                $assignableList = ($assignableItems | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '
                $summary = if ($targetUsage) {
                    "Set usage location to {0} and assign license(s): {1} to {2}" -f $targetUsage, $assignableList, $upn
                }
                else {
                    "Assign license(s): {0} to {1}" -f $assignableList, $upn
                }

                if (-not $PSCmdlet.ShouldProcess($upn, $summary)) {
                    continue
                }

                # Reserve the units now so later users in the same chunk see the reduced availability.
                foreach ($item in $assignableItems) { $remaining[[string]$item.SkuId] = [int]$remaining[[string]$item.SkuId] - 1 }

                $approved.Add([pscustomobject]@{
                        Id                  = [string]$user.id
                        UserPrincipalName   = $upn
                        TargetUsage         = $targetUsage
                        AssignableItems     = $assignableItems
                        AssignableList      = $assignableList
                        NamesNoAvailability = $namesNoAvailability
                        Ready               = $true
                    })
            }
            if ($approved.Count -eq 0) { return }

            $releaseUnits = {
                param($Approved)
                foreach ($item in $Approved.AssignableItems) { $remaining[[string]$item.SkuId] = [int]$remaining[[string]$item.SkuId] + 1 }
            }

            # (c) Pass 1: usage location for the users that need it.
            $patchIndexes = @(for ($i = 0; $i -lt $approved.Count; $i++) { if ($approved[$i].TargetUsage) { $i } })
            if ($patchIndexes.Count -gt 0) {
                $patchRequests = @(foreach ($i in $patchIndexes) {
                        @{ Id = "p$i"; Method = 'PATCH'; Url = "/users/$([uri]::EscapeDataString($approved[$i].Id))"; Body = @{ usageLocation = $approved[$i].TargetUsage } }
                    })
                $patchResponses = @(Invoke-NCGraphBatch -Requests $patchRequests -Activity 'Setting usage location')
                for ($k = 0; $k -lt $patchIndexes.Count; $k++) {
                    $i = $patchIndexes[$k]
                    $response = $patchResponses[$k]
                    if ($response.Success) {
                        Write-Verbose "Usage location set to $($approved[$i].TargetUsage) for $($approved[$i].UserPrincipalName)."
                    }
                    else {
                        Write-NCMessage "Unable to set usage location ($($approved[$i].TargetUsage)) for $($approved[$i].UserPrincipalName): $($response.ErrorMessage)" -Level ERROR
                        $approved[$i].Ready = $false
                        & $releaseUnits $approved[$i]
                    }
                }
            }

            # (d) Pass 2: assign licenses for users that are ready.
            $assignIndexes = @(for ($i = 0; $i -lt $approved.Count; $i++) { if ($approved[$i].Ready) { $i } })
            if ($assignIndexes.Count -eq 0) { return }
            $assignRequests = @(foreach ($i in $assignIndexes) {
                    @{
                        Id     = "m$i"
                        Method = 'POST'
                        Url    = "/users/$([uri]::EscapeDataString($approved[$i].Id))/assignLicense"
                        Body   = @{
                            addLicenses    = @(foreach ($item in $approved[$i].AssignableItems) { @{ skuId = [string]$item.SkuId; disabledPlans = @() } })
                            removeLicenses = @()
                        }
                    }
                })
            $assignResponses = @(Invoke-NCGraphBatch -Requests $assignRequests -Activity 'Assigning licenses')
            for ($k = 0; $k -lt $assignIndexes.Count; $k++) {
                $i = $assignIndexes[$k]
                $response = $assignResponses[$k]
                if ($response.Success) {
                    Write-NCMessage ("Assigned license(s) to {0}: {1}" -f $approved[$i].UserPrincipalName, $approved[$i].AssignableList) -Level SUCCESS
                    if ($approved[$i].NamesNoAvailability.Count -gt 0) {
                        Write-NCMessage ("Skipped license(s) with no available units: {0}" -f (($approved[$i].NamesNoAvailability | Select-Object -Unique) -join ', ')) -Level WARNING
                    }
                }
                else {
                    Write-NCMessage "License assignment failed for $($approved[$i].UserPrincipalName): $($response.ErrorMessage)" -Level ERROR
                    & $releaseUnits $approved[$i]
                }
            }
        }
    }

    process {
        if (-not $graphConnected -or -not $ready) {
            return
        }

        $queue.Add($UserPrincipalName)
        if ($queue.Count -ge 20) {
            & $flush
        }
    }

    end {
        try {
            if ($graphConnected -and $ready) {
                & $flush
            }
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Copy-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Copies all licenses from one user to another without removing them from the source.
    .DESCRIPTION
        Reads source user licenses (including disabled service plans) and assigns them to the destination user
        if they are not already present. Uses Microsoft Graph and the cached license catalog for friendly names.
    .PARAMETER SourceUserPrincipalName
        UserPrincipalName or object ID of the source user.
    .PARAMETER DestinationUserPrincipalName
        UserPrincipalName or object ID of the destination user.
    .EXAMPLE
        Copy-UserMsolAccountSku -SourceUserPrincipalName user1@contoso.com -DestinationUserPrincipalName user2@contoso.com
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias('Source', 'From')]
        [string]$SourceUserPrincipalName,
        [Parameter(Mandatory = $true, Position = 1)]
        [Alias('Destination', 'To')]
        [string]$DestinationUserPrincipalName
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $resolvedUsers = Resolve-NCGraphUserBatch -Identifier @($SourceUserPrincipalName, $DestinationUserPrincipalName) -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation') -FailedIdentifier $failedUsers

        $sourceUser = $resolvedUsers[$SourceUserPrincipalName.Trim()]
        if (-not $sourceUser -or -not $sourceUser.id) {
            if (-not $failedUsers.Contains($SourceUserPrincipalName.Trim())) {
                Write-NCMessage "Unable to retrieve source user $($SourceUserPrincipalName): user not found." -Level ERROR
            }
            return
        }

        $destinationUser = $resolvedUsers[$DestinationUserPrincipalName.Trim()]
        if (-not $destinationUser -or -not $destinationUser.id) {
            if (-not $failedUsers.Contains($DestinationUserPrincipalName.Trim())) {
                Write-NCMessage "Unable to retrieve destination user $($DestinationUserPrincipalName): user not found." -Level ERROR
            }
            return
        }

        if ([string]$sourceUser.id -eq [string]$destinationUser.id) {
            Write-NCMessage "Source and destination users are the same. Aborting." -Level ERROR
            return
        }

        $defaultUsageLocation = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) {
            [string]$NCVars.UsageLocation
        }
        else { 'US' }

        $normalizeUsageLocation = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return $value.Trim().ToUpperInvariant()
        }

        $currentUsage = & $normalizeUsageLocation $destinationUser.usageLocation
        $desiredUsage = & $normalizeUsageLocation $defaultUsageLocation

        if ($desiredUsage -and $desiredUsage -ne $currentUsage) {
            $targetUsage = $defaultUsageLocation
            try {
                Update-MgUser -UserId $destinationUser.id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                Write-Verbose "Usage location set to $targetUsage for $($destinationUser.userPrincipalName)."
            }
            catch {
                Write-NCMessage "Unable to set usage location ($targetUsage) for $($destinationUser.userPrincipalName): $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        try {
            $licenseCatalog = Get-LicenseCatalog
        }
        catch {
            Write-NCMessage "License catalog unavailable: $($_.Exception.Message)" -Level WARNING
            $licenseCatalog = $null
        }

        $licenseLookup = $null
        $customLookup = $null
        if ($licenseCatalog) {
            if ($licenseCatalog.PSObject.Properties['Lookup']) {
                $licenseLookup = $licenseCatalog.Lookup
            }
            if ($licenseCatalog.PSObject.Properties['CustomLookup']) {
                $customLookup = $licenseCatalog.CustomLookup
            }
        }
        $maxAttempts = 3

        $licenseDetailRequests = @(
            @{ Id = 'l0'; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$sourceUser.id))/licenseDetails" }
            @{ Id = 'l1'; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$destinationUser.id))/licenseDetails" }
        )
        $licenseDetailResults = @(Invoke-NCGraphBatchCollection -Requests $licenseDetailRequests -Activity 'Reading user licenses')

        if (-not $licenseDetailResults[0].Success) {
            Write-NCMessage "Failed to retrieve licenses for $($sourceUser.userPrincipalName) after $maxAttempts attempts." -Level ERROR
            Write-Verbose $licenseDetailResults[0].ErrorMessage
            return
        }
        $sourceLicenses = if ($null -ne $licenseDetailResults[0].Items) { @($licenseDetailResults[0].Items) } else { @() }

        if ($sourceLicenses.Count -eq 0) {
            Write-NCMessage "Source user $($sourceUser.userPrincipalName) has no licenses to copy." -Level WARNING
            return
        }

        if (-not $licenseDetailResults[1].Success) {
            Write-NCMessage "Unable to read destination licenses for $($destinationUser.userPrincipalName): $($licenseDetailResults[1].ErrorMessage)" -Level ERROR
            return
        }
        $destinationLicenses = if ($null -ne $licenseDetailResults[1].Items) { @($licenseDetailResults[1].Items) } else { @() }

        try {
            $tenantSkus = Invoke-NCRetry -Action {
                Get-MgSubscribedSku -All -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve tenant licenses" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve tenant licenses, attempt $currentAttempt of $currentMax." -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Unable to retrieve tenant licenses after $maxAttempts attempts." -Level ERROR
            return
        }

        $availabilityBySkuId = @{}
        foreach ($sku in $tenantSkus) {
            $skuIdString = [string]$sku.SkuId
            if ([string]::IsNullOrWhiteSpace($skuIdString)) { continue }
            $prepaidUnits = $sku.PrepaidUnits
            $enabledUnits = if ($prepaidUnits) { [int]$prepaidUnits.Enabled } else { 0 }
            $consumedUnits = if ($sku.ConsumedUnits -is [int]) { [int]$sku.ConsumedUnits } else { [int]0 }
            $availabilityBySkuId[$skuIdString.ToUpperInvariant()] = [Math]::Max($enabledUnits - $consumedUnits, 0)
        }

        $destinationSkuIds = @()
        foreach ($destSku in $destinationLicenses) {
            $parsedSku = [guid]::Empty
            if ([guid]::TryParse([string]$destSku.SkuId, [ref]$parsedSku)) { $destinationSkuIds += $parsedSku }
        }
        $addLicenses = @()
        $licenseNames = @()
        $skippedInvalid = @()
        $namesNoAvailability = @()

        foreach ($lic in $sourceLicenses) {
            $skuIdString = [string]$lic.SkuId

            if ([string]::IsNullOrWhiteSpace($skuIdString)) {
                $skippedInvalid += "empty SkuId ($($lic.SkuPartNumber))"
                continue
            }

            $parsedGuid = [guid]::Empty
            if (-not [guid]::TryParse($skuIdString, [ref]$parsedGuid)) {
                $skippedInvalid += "invalid SkuId '$skuIdString' ($($lic.SkuPartNumber))"
                continue
            }

            if ($destinationSkuIds -contains $parsedGuid) {
                Write-Verbose "Destination already has $($lic.SkuPartNumber); skipping add."
                continue
            }

            $available = $availabilityBySkuId[$skuIdString.ToUpperInvariant()]
            if ($null -eq $available) { $available = 0 }
            if ($available -le 0) {
                Write-NCMessage ("No available units for license {0} (available: {1}); skipping." -f $lic.SkuPartNumber, $available) -Level WARNING
                $namesNoAvailability += $lic.SkuPartNumber
                continue
            }

            $validatedDisabled = @()
            if ($lic.DisabledPlans) {
                foreach ($plan in $lic.DisabledPlans) {
                    $planString = [string]$plan
                    if ([guid]::TryParse($planString, [ref]([guid]::Empty))) {
                        $validatedDisabled += $plan
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace($planString)) {
                        Write-Verbose "Skipping invalid disabled plan '$planString' for $($lic.SkuPartNumber)."
                    }
                }
            }

            $addLicenses += @{
                SkuId         = $parsedGuid
                DisabledPlans = $validatedDisabled
            }

            $matchSource = $null
            $name = if ($licenseLookup) {
                Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $lic.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
            }
            if ($name) {
                $licenseNames += $name
            }
            else {
                $licenseNames += $lic.SkuPartNumber
            }
        }

        if ($addLicenses.Count -eq 0) {
            if ($namesNoAvailability.Count -gt 0) {
                Write-NCMessage ("Nothing to copy: no available units for {0}." -f (($namesNoAvailability | Select-Object -Unique) -join ', ')) -Level WARNING
            }
            else {
                Write-NCMessage "Nothing to copy: destination already has all licenses from $($sourceUser.userPrincipalName)." -Level WARNING
            }
            return
        }

        $uniqueNames = $licenseNames | Select-Object -Unique
        $actionSummary = "Copy licenses ($($uniqueNames -join ', ')) from $($sourceUser.userPrincipalName) to $($destinationUser.userPrincipalName)"

        if ($skippedInvalid.Count -gt 0) {
            Write-NCMessage ("Skipped licenses with invalid IDs: {0}" -f ($skippedInvalid -join '; ')) -Level WARNING
        }

        if (-not $PSCmdlet.ShouldProcess($destinationUser.userPrincipalName, $actionSummary)) {
            return
        }

        try {
            Invoke-NCRetry -Action {
                Set-MgUserLicense -UserId $destinationUser.id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "assign licenses to $($destinationUser.userPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage ("Failed to assign licenses to {0}, attempt {1} of {2}. {3}" -f $destinationUser.userPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
            } | Out-Null
            Write-NCMessage "Copied licenses to $($destinationUser.userPrincipalName): $($uniqueNames -join ', ')." -Level SUCCESS
            if ($namesNoAvailability.Count -gt 0) {
                Write-NCMessage ("Skipped license(s) with no available units: {0}" -f (($namesNoAvailability | Select-Object -Unique) -join ', ')) -Level WARNING
            }
        }
        catch {
            Write-NCMessage "License copy to $($destinationUser.userPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
        }
    }
    finally {
        Restore-ProgressAndInfoPreferences
    }
}

function Export-MsolAccountSku {
    <#
    .SYNOPSIS
        Exports assigned Microsoft 365 licenses to CSV.
    .DESCRIPTION
        Connects to Microsoft Graph, downloads the license catalog, iterates all licensed users,
        maps SKU part numbers to friendly names, and writes/resumes a CSV report.
    .PARAMETER CSVFolder
        Output folder (defaults to the current directory if omitted).
    .PARAMETER Domain
        Optional domain filter. When specified, only users whose Mail, UserPrincipalName,
        or ProxyAddresses belong to that domain are exported.
    .PARAMETER License
        Optional license filter. When specified, only users who have at least one matching
        license are exported, but all of their assigned licenses are still included in the report.
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .PARAMETER BatchSize
        Number of processed users before flushing partial CSV output.
    .PARAMETER Resume
        Resume from the latest matching CSV in the target folder or from -CsvPath.
    .PARAMETER CsvPath
        Explicit CSV file to resume. When omitted, the most recent matching CSV in the target folder is used.
    .PARAMETER MaxConsecutiveErrors
        Stop after this many consecutive user-level failures.
    .EXAMPLE
        Export-MsolAccountSku
    .EXAMPLE
        Export-MsolAccountSku -CsvFolder "C:\Temp"
    .EXAMPLE
        Export-MsolAccountSku -Domain "contoso.com"
    .EXAMPLE
        Export-MsolAccountSku -License "Exchange Online (Plan 1)"
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $True, HelpMessage = "Folder where export CSV file (e.g. C:\Temp)")]
        [string]$CSVFolder,
        [string]$Domain,
        [string[]]$License,
        [switch]$ForceLicenseCatalogRefresh,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 50,
        [switch]$Resume,
        [string]$CsvPath,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $folder = Test-Folder($CSVFolder)
        try {
            $licenseCatalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent
        }
        catch {
            Write-NCMessage $_ -Level ERROR
            return
        }

        $licenseLookup = $licenseCatalog.Lookup
        $customLookup = $licenseCatalog.CustomLookup
        $tenantSkus = @()
        if ($License) {
            try {
                $tenantSkus = Invoke-NCRetry -Action {
                    Get-MgSubscribedSku -All -ErrorAction Stop
                } -MaxAttempts 3 -DelaySeconds 5 -OperationDescription "retrieve tenant licenses" -OnError {
                    param($attempt, $max, $err)
                    $currentAttempt = if ($attempt) { $attempt } else { '?' }
                    $currentMax = if ($max) { $max } else { 3 }
                    Write-NCMessage "Failed to retrieve tenant licenses, attempt $currentAttempt of $currentMax." -Level ERROR
                }
            }
            catch {
                Write-NCMessage "Unable to retrieve tenant licenses for license filtering." -Level ERROR
                return
            }

            if (-not $tenantSkus -or $tenantSkus.Count -eq 0) {
                Write-NCMessage "No tenant licenses available to filter on." -Level WARNING
                return
            }
        }
        $reportBuffer = [System.Collections.Generic.List[object]]::new()
        $processedCount = 0
        $processedSinceFlush = 0
        $consecutiveErrors = 0
        $aborted = $false
        $maxAttempts = 3
        $resolvedViaCustom = @{}
        $unknownSkuTracker = @{}
        $normalizedDomain = $null
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $normalizedDomain = $Domain.Trim().ToLowerInvariant().TrimStart('@')
            if ([string]::IsNullOrWhiteSpace($normalizedDomain)) {
                Write-NCMessage "Domain filter cannot be empty." -Level ERROR
                return
            } else {
                Write-Verbose "Filtering users for domain '$normalizedDomain'."
            }
        }

        $getAddressDomain = {
            param($value)

            if ([string]::IsNullOrWhiteSpace($value)) {
                return $null
            }

            $address = [string]$value
            if ($address.Contains(':')) {
                $address = $address.Substring($address.IndexOf(':') + 1)
            }

            $atIndex = $address.LastIndexOf('@')
            if ($atIndex -lt 0 -or $atIndex -eq ($address.Length - 1)) {
                return $null
            }

            return $address.Substring($atIndex + 1).Trim().ToLowerInvariant()
        }

        $normalizeText = {
            param($value)

            if ([string]::IsNullOrWhiteSpace($value)) {
                return $null
            }

            return $value.Trim().ToUpperInvariant()
        }

        $resolveLicenseFilter = {
            param(
                [string[]]$RequestedLicenses,
                $TenantSkus,
                $Lookup,
                $FallbackLookup
            )

            $resolved = @()
            $unmatched = @()

            $requestedItems = $RequestedLicenses |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim() } |
                Select-Object -Unique

            foreach ($entry in $requestedItems) {
                $target = & $normalizeText $entry
                $match = $null

                foreach ($sku in $TenantSkus) {
                    $skuIdString = [string]$sku.SkuId
                    $skuIdNormalized = if ($skuIdString) { $skuIdString.ToUpperInvariant() } else { $null }
                    $skuPart = & $normalizeText $sku.SkuPartNumber

                    $display = $null
                    $matchSource = $null
                    if ($Lookup) {
                        $display = Get-LicenseDisplayName -Lookup $Lookup -SkuPartNumber $sku.SkuPartNumber -FallbackLookup $FallbackLookup -MatchSource ([ref]$matchSource)
                    }
                    $displayNormalized = if ($display) { & $normalizeText $display } else { $null }

                    if ($target -eq $skuPart -or $target -eq $skuIdNormalized -or ($displayNormalized -and $target -eq $displayNormalized)) {
                        $match = [pscustomobject]@{
                            SkuId         = $sku.SkuId
                            SkuPartNumber = $sku.SkuPartNumber
                            Name          = if ($display) { $display } else { $sku.SkuPartNumber }
                        }
                        break
                    }
                }

                if ($match) {
                    $resolved += $match
                }
                else {
                    $unmatched += $entry
                }
            }

            [pscustomobject]@{
                Resolved  = $resolved
                Unmatched = $unmatched
            }
        }

        $licenseFilter = $null
        $licenseFilterIds = @()
        $licenseFilterParts = @()
        $licenseFilterNames = @()
        if ($License) {
            $licenseFilter = & $resolveLicenseFilter $License $tenantSkus $licenseLookup $customLookup
            if ($licenseFilter.Unmatched.Count -gt 0) {
                Write-NCMessage ("Unable to resolve license filter(s): {0}" -f ($licenseFilter.Unmatched -join ', ')) -Level ERROR
                return
            }

            $licenseFilterIds = @(
                $licenseFilter.Resolved |
                    ForEach-Object { ([string]$_.SkuId).ToUpperInvariant() } |
                    Where-Object { $_ } |
                    Select-Object -Unique
            )
            $licenseFilterParts = @(
                $licenseFilter.Resolved |
                    ForEach-Object { & $normalizeText $_.SkuPartNumber } |
                    Where-Object { $_ } |
                    Select-Object -Unique
            )
            $licenseFilterNames = @(
                $licenseFilter.Resolved |
                    ForEach-Object { $_.Name } |
                    Where-Object { $_ } |
                    Select-Object -Unique
            )

            Write-Verbose ("Filtering users by license(s): {0}" -f ($licenseFilterNames -join ', '))
        }

        $matchesDomain = {
            param($user)

            if (-not $normalizedDomain) {
                return $true
            }

            $domains = @()

            foreach ($candidate in @($user.Mail, $user.UserPrincipalName)) {
                $candidateDomain = & $getAddressDomain $candidate
                if ($candidateDomain) {
                    $domains += $candidateDomain
                }
            }

            if ($user.PSObject.Properties['ProxyAddresses'] -and $user.ProxyAddresses) {
                foreach ($proxy in $user.ProxyAddresses) {
                    $proxyDomain = & $getAddressDomain $proxy
                    if ($proxyDomain) {
                        $domains += $proxyDomain
                    }
                }
            }

            return ($domains | Select-Object -Unique) -contains $normalizedDomain
        }

        $defaultCsvPath = New-File("$($folder)\$((Get-Date -Format $($NCVars.DateTimeString_CSV)).ToString())_M365-User-License-Report.csv")
        $CSV = $defaultCsvPath
        $processedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        if ($Resume) {
            $resumePath = $null
            if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                $resumePath = $CsvPath
            }
            else {
                $existingCsv = Get-ChildItem -LiteralPath $folder -File -Filter "*_M365-User-License-Report.csv" |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1
                if ($existingCsv) {
                    $resumePath = $existingCsv.FullName
                }
            }

            if ($resumePath) {
                $CSV = $resumePath
                if (Test-Path -LiteralPath $CSV) {
                    try {
                        foreach ($row in (Import-CSV -LiteralPath $CSV -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                            $identity = if ($row.UserPrincipalName) { [string]$row.UserPrincipalName } else { $null }
                            if ($identity) {
                                $null = $processedUsers.Add($identity.Trim())
                            }
                        }
                        Write-NCMessage ("Resuming user license export from {0}; {1} user(s) already recorded." -f $CSV, $processedUsers.Count) -Level INFO
                    }
                    catch {
                        Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $CSV, $_.Exception.Message) -Level WARNING
                        $processedUsers.Clear()
                        $CSV = $defaultCsvPath
                    }
                }
                else {
                    Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $CSV) -Level INFO
                }
            }
            else {
                $CSV = $defaultCsvPath
                Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $CSV) -Level INFO
            }
        }

        Write-NCMessage ("User license export will flush every {0} user(s). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
        Write-NCMessage "Saving report to $CSV" -Level DEBUG

        try {
            $Users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable totalUsers -All -Property Id,DisplayName,UserPrincipalName,Mail,ProxyAddresses -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Failed to retrieve users with assigned licenses: $_" -Level ERROR
            return
        }

        if ($normalizedDomain) {
            $Users = $Users | Where-Object { & $matchesDomain $_ }
            $totalUsers = @($Users).Count
            if ($totalUsers -eq 0) {
                Write-NCMessage "No licensed users match domain '$normalizedDomain'." -Level WARNING
                return
            }
        }

        $allUsers = @($Users)
        $licenseDetailCache = @{}
        $batchNoticeShown = $false

        for ($userIndex = 0; $userIndex -lt $allUsers.Count; $userIndex++) {
            $User = $allUsers[$userIndex]
            $processedCount++
            $Percentage = Get-NCProgressPercent -Current $processedCount -Total $totalUsers
            Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$processedCount out of $totalUsers - $Percentage%" -PercentComplete $Percentage

            if ($processedUsers.Contains($User.UserPrincipalName)) {
                Write-NCMessage "Skipping $($User.UserPrincipalName), already processed." -Level WARNING
                continue
            }

            $processedSinceFlush++

            # Read license details for this user and the next pending users in one Graph batch (20 per request).
            if (-not $licenseDetailCache.ContainsKey([string]$User.Id)) {
                if (-not $batchNoticeShown) {
                    $batchNoticeShown = $true
                    Write-NCGraphBatchNotice -Count $totalUsers -Noun 'user(s)'
                }

                $pending = [System.Collections.Generic.List[object]]::new()
                for ($j = $userIndex; $j -lt $allUsers.Count -and $pending.Count -lt 20; $j++) {
                    $candidate = $allUsers[$j]
                    if ($processedUsers.Contains($candidate.UserPrincipalName)) { continue }
                    if ($licenseDetailCache.ContainsKey([string]$candidate.Id)) { continue }
                    $pending.Add($candidate)
                }

                $detailRequests = @(for ($k = 0; $k -lt $pending.Count; $k++) {
                        @{ Id = "l$k"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$pending[$k].Id))/licenseDetails" }
                    })
                $detailResults = @(Invoke-NCGraphBatchCollection -Requests $detailRequests -Activity 'Reading user licenses')
                for ($k = 0; $k -lt $pending.Count; $k++) {
                    $licenseDetailCache[[string]$pending[$k].Id] = $detailResults[$k]
                }
            }

            $detailResult = $licenseDetailCache[[string]$User.Id]
            $licenseDetailCache.Remove([string]$User.Id)

            if (-not $detailResult.Success) {
                Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName) after $maxAttempts attempts. Skipping." -Level ERROR
                $consecutiveErrors++
                if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                    if ($reportBuffer.Count -gt 0) {
                        if ((Test-Path -LiteralPath $CSV) -and ((Get-Item -LiteralPath $CSV).Length -gt 0)) {
                            $reportBuffer | Export-CSV -LiteralPath $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding) -Append
                        }
                        else {
                            $reportBuffer | Export-CSV -LiteralPath $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding)
                        }
                        $reportBuffer.Clear()
                    }

                    Write-NCMessage ("Stopping export after {0} consecutive user error(s). Partial report kept at {1}." -f $consecutiveErrors, $CSV) -Level ERROR
                    $aborted = $true
                    break
                }
                continue
            }

            $GraphLicense = if ($null -ne $detailResult.Items) { @($detailResult.Items) } else { @() }

            $userMatchesLicenseFilter = -not $License
            $userMatchedLicenseNames = @()
            $userRows = @()

            if ($null -ne $GraphLicense) {
                foreach ($licenseSku in $GraphLicense) {
                    $matchSource = $null
                    $skuIdString = if ($licenseSku.SkuId) { [string]$licenseSku.SkuId } else { $null }
                    $skuIdNormalized = if ($skuIdString) { $skuIdString.ToUpperInvariant() } else { $null }
                    $skuPartNormalized = & $normalizeText $licenseSku.SkuPartNumber

                    $productName = Get-LicenseDisplayName -Lookup $licenseLookup `
                        -SkuPartNumber $licenseSku.SkuPartNumber `
                        -FallbackLookup $customLookup `
                        -MatchSource ([ref]$matchSource)

                    if (-not $productName) {
                        Write-Verbose "Unknown license: $($licenseSku.SkuPartNumber) for $($User.UserPrincipalName)"
                        if ($unknownSkuTracker.ContainsKey($licenseSku.SkuPartNumber)) {
                            $unknownSkuTracker[$licenseSku.SkuPartNumber]++
                        }
                        else {
                            $unknownSkuTracker[$licenseSku.SkuPartNumber] = 1
                        }
                        $productName = $licenseSku.SkuPartNumber
                    }
                    elseif ($matchSource -and $matchSource -ne 'Primary') {
                        if ($resolvedViaCustom.ContainsKey($licenseSku.SkuPartNumber)) {
                            $resolvedViaCustom[$licenseSku.SkuPartNumber]++
                        }
                        else {
                            $resolvedViaCustom[$licenseSku.SkuPartNumber] = 1
                        }
                    }

                    if ($License -and -not $userMatchesLicenseFilter) {
                        if (($skuIdNormalized -and ($licenseFilterIds -contains $skuIdNormalized)) -or
                            ($skuPartNormalized -and ($licenseFilterParts -contains $skuPartNormalized))) {
                            $userMatchesLicenseFilter = $true
                            $userMatchedLicenseNames += if ($productName) { $productName } else { $licenseSku.SkuPartNumber }
                        }
                    }
                    elseif ($License) {
                        if (($skuIdNormalized -and ($licenseFilterIds -contains $skuIdNormalized)) -or
                            ($skuPartNormalized -and ($licenseFilterParts -contains $skuPartNormalized))) {
                            $userMatchedLicenseNames += if ($productName) { $productName } else { $licenseSku.SkuPartNumber }
                        }
                    }

                    $userRows += [pscustomobject]@{
                        DisplayName        = $User.DisplayName
                        UserPrincipalName  = $User.UserPrincipalName
                        PrimarySmtpAddress = $User.Mail
                        Licenses           = $productName
                    }
                }
            }

            if ($License -and -not $userMatchesLicenseFilter) {
                continue
            }

            if ($License -and $userMatchedLicenseNames.Count -gt 0) {
                $matchedNames = $userMatchedLicenseNames | Select-Object -Unique
                $userRows = @($userRows | ForEach-Object {
                    $_ | Add-Member -NotePropertyName MatchedLicenses -NotePropertyValue ($matchedNames -join ', ') -PassThru
                })
            }
            elseif ($License) {
                $userRows = @($userRows | ForEach-Object {
                    $_ | Add-Member -NotePropertyName MatchedLicenses -NotePropertyValue $null -PassThru
                })
            }

            if ($userRows.Count -gt 0) {
                $reportBuffer.AddRange($userRows)
                $null = $processedUsers.Add($User.UserPrincipalName)
            }

            $consecutiveErrors = 0

            if ($processedSinceFlush -ge $BatchSize -and $reportBuffer.Count -gt 0) {
                if ((Test-Path -LiteralPath $CSV) -and ((Get-Item -LiteralPath $CSV).Length -gt 0)) {
                    $reportBuffer | Export-CSV -LiteralPath $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding) -Append
                }
                else {
                    $reportBuffer | Export-CSV -LiteralPath $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding)
                }
                Write-Verbose "Processed $processedCount out of $totalUsers, saving partial results ..."
                $reportBuffer.Clear()
                $processedSinceFlush = 0
            }
        }

        if ($reportBuffer.Count -gt 0) {
            if ((Test-Path -LiteralPath $CSV) -and ((Get-Item -LiteralPath $CSV).Length -gt 0)) {
                $reportBuffer | Export-CSV -LiteralPath $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding) -Append
            }
            else {
                $reportBuffer | Export-CSV -LiteralPath $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding)
            }
            $reportBuffer.Clear()
        }

        if ($aborted) {
            Write-NCMessage "User license report export stopped early. Partial data kept at $CSV." -Level ERROR
        }
        else {
            Write-NCMessage "User license report exported to $CSV." -Level SUCCESS
        }

        if ($resolvedViaCustom.Count -gt 0) {
            Write-NCMessage "Licenses not found, but resolved via custom catalog:" -Level WARNING
            foreach ($sku in ($resolvedViaCustom.Keys | Sort-Object)) {
                $count = $resolvedViaCustom[$sku]
                Write-NCMessage (" - {0} ({1} occurrence{2})" -f $sku, $count, $(if ($count -ne 1) { 's' } else { '' })) -Level WARNING
            }
        }

        if ($unknownSkuTracker.Count -gt 0) {
            Write-NCMessage "Licenses still without mappings:" -Level WARNING
            foreach ($sku in ($unknownSkuTracker.Keys | Sort-Object)) {
                $count = $unknownSkuTracker[$sku]
                Write-NCMessage (" - {0} ({1} occurrence{2})" -f $sku, $count, $(if ($count -ne 1) { 's' } else { '' })) -Level WARNING
            }
        }

        Write-Progress -Activity "Export complete" -Completed
    }
    finally {
        Restore-ProgressAndInfoPreferences
    }
}

function Get-TenantMsolAccountSku {
    <#
    .SYNOPSIS
        Lists available tenant licenses with resolved names and usage counts.
    .DESCRIPTION
        Connects to Microsoft Graph, loads the license catalog, retrieves all tenant SKUs, resolves
        part numbers to friendly names (using the same lookup logic as other license functions), and
        returns counts for total, consumed, available (net of suspended), suspended, and warning seats.
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .PARAMETER Filter
        Filters the output to licenses whose name or SkuPartNumber contains the provided text.
    .PARAMETER Domain
        Optional domain filter for sample users. When specified, sample users are limited to
        accounts whose Mail, UserPrincipalName, or ProxyAddresses belong to the selected domain.
    .PARAMETER SampleUsers
        Returns up to N sample users for each matching SKU (requires -Filter).
    .PARAMETER IncludeSampleUsers
        Returns sample users using the default limit of 5 (requires -Filter).
    .PARAMETER AsTable
        Display the result as a formatted table instead of returning objects.
    .PARAMETER GridView
        Show the result in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-TenantMsolAccountSku
    .EXAMPLE
        Get-TenantMsolAccountSku -AsTable
    .EXAMPLE
        Get-TenantMsolAccountSku -Filter "E3"
    .EXAMPLE
        Get-TenantMsolAccountSku -Filter "E3" -SampleUsers 5
    .EXAMPLE
        Get-TenantMsolAccountSku -Filter "E3" -IncludeSampleUsers
    .EXAMPLE
        Get-TenantMsolAccountSku -Filter "E3" -SampleUsers 5 -Domain "contoso.com"
    #>
    [CmdletBinding()]
    param(
        [switch]$ForceLicenseCatalogRefresh,
        [string]$Filter,
        [string]$Domain,
        [int]$SampleUsers = 5,
        [switch]$IncludeSampleUsers,
        [switch]$AsTable,
        [switch]$GridView
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        try {
            $licenseCatalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent
        }
        catch {
            Write-NCMessage $_ -Level ERROR
            return
        }

        $licenseLookup = $licenseCatalog.Lookup
        $customLookup = $licenseCatalog.CustomLookup
        $maxAttempts = 3

        try {
            $skus = Invoke-NCRetry -Action {
                Get-MgSubscribedSku -All -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve tenant licenses" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve tenant licenses, attempt $currentAttempt of $currentMax." -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Failed to retrieve tenant licenses after $maxAttempts attempts." -Level ERROR
            return
        }

        if (-not $skus -or $skus.Count -eq 0) {
            Write-NCMessage "No licenses found for this tenant." -Level WARNING
            return
        }

        $useSampleUsers = $PSBoundParameters.ContainsKey('SampleUsers') -or $IncludeSampleUsers.IsPresent
        $sampleUserLimit = if ($PSBoundParameters.ContainsKey('SampleUsers')) { $SampleUsers } else { 5 }

        if ($useSampleUsers -and -not $Filter) {
            Write-NCMessage "SampleUsers requires -Filter to limit the query scope. Example: Get-TenantMsolAccountSku -Filter \"E3\" -SampleUsers 5" -Level ERROR
            return
        }

        $normalizedDomain = $null
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $normalizedDomain = $Domain.Trim().ToLowerInvariant().TrimStart('@')
            if ([string]::IsNullOrWhiteSpace($normalizedDomain)) {
                Write-NCMessage "Domain filter cannot be empty." -Level ERROR
                return
            }
            else {
                Write-Verbose "Filtering sample users for domain '$normalizedDomain'."
            }
        }

        $getAddressDomain = {
            param($value)

            if ([string]::IsNullOrWhiteSpace($value)) {
                return $null
            }

            $address = [string]$value
            if ($address.Contains(':')) {
                $address = $address.Substring($address.IndexOf(':') + 1)
            }

            $atIndex = $address.LastIndexOf('@')
            if ($atIndex -lt 0 -or $atIndex -eq ($address.Length - 1)) {
                return $null
            }

            return $address.Substring($atIndex + 1).Trim().ToLowerInvariant()
        }

        $matchesDomain = {
            param($user)

            if (-not $normalizedDomain) {
                return $true
            }

            $domains = @()
            foreach ($candidate in @($user.Mail, $user.UserPrincipalName)) {
                $candidateDomain = & $getAddressDomain $candidate
                if ($candidateDomain) {
                    $domains += $candidateDomain
                }
            }

            if ($user.PSObject.Properties['ProxyAddresses'] -and $user.ProxyAddresses) {
                foreach ($proxy in $user.ProxyAddresses) {
                    $proxyDomain = & $getAddressDomain $proxy
                    if ($proxyDomain) {
                        $domains += $proxyDomain
                    }
                }
            }

            return ($domains | Select-Object -Unique) -contains $normalizedDomain
        }

        $results = foreach ($sku in $skus) {
            $matchSource = $null
            $display = Get-LicenseDisplayName -Lookup $licenseLookup `
                -SkuPartNumber $sku.SkuPartNumber `
                -FallbackLookup $customLookup `
                -MatchSource ([ref]$matchSource)

            $prepaid = $sku.PrepaidUnits
            $enabled = if ($prepaid) { [int]$prepaid.Enabled } else { 0 }
            $suspended = if ($prepaid) { [int]$prepaid.Suspended } else { 0 }
            $warning = if ($prepaid) { [int]$prepaid.Warning } else { 0 }
            $totalCount = $enabled + $suspended + $warning
            $total = "{0} (Enabled: {1}, Suspended: {2})" -f $totalCount, $enabled, $suspended
            $consumed = if ($sku.ConsumedUnits -is [int]) { [int]$sku.ConsumedUnits } else { [int]0 }
            $available = [Math]::Max($enabled - $consumed, 0)
            $nameSource = if ($matchSource) { $matchSource } elseif ($display) { 'Primary' } else { 'Unknown' }

            [pscustomobject][ordered]@{
                Name          = if ($display) { $display } else { $sku.SkuPartNumber }
                SkuPartNumber = $sku.SkuPartNumber
                SkuId         = $sku.SkuId
                Total         = $total
                TotalCount    = $totalCount
                Consumed      = $consumed
                Available     = $available
                Enabled       = $enabled
                Suspended     = $suspended
                Warning       = $warning
                Source        = $nameSource
            }
        }

        $sorted = $results | Sort-Object Name

        if ($Filter) {
            $filterPattern = [regex]::Escape($Filter)
            $sorted = $sorted | Where-Object {
                $_.Name -match $filterPattern -or $_.SkuPartNumber -match $filterPattern
            }

            if (-not $sorted -or $sorted.Count -eq 0) {
                Write-NCMessage "No licenses match filter '$Filter'." -Level WARNING
                return
            }
        }

        if ($useSampleUsers) {
            if ($sampleUserLimit -le 0) {
                Write-NCMessage "SampleUsers must be greater than 0." -Level ERROR
                return
            }

            $sorted = foreach ($sku in $sorted) {
                $sampleUserList = @()
                try {
                    $examples = Invoke-NCRetry -Action {
                        if ($normalizedDomain) {
                            Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($sku.SkuId))" `
                                -All `
                                -ConsistencyLevel eventual `
                                -Property Id, UserPrincipalName, DisplayName, Mail, ProxyAddresses `
                                -ErrorAction Stop |
                            Where-Object { & $matchesDomain $_ } |
                            Select-Object -First $sampleUserLimit
                        }
                        else {
                            Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($sku.SkuId))" `
                                -Top $sampleUserLimit `
                                -ConsistencyLevel eventual `
                                -Property Id, UserPrincipalName, DisplayName `
                                -ErrorAction Stop |
                            Select-Object -First $sampleUserLimit
                        }
                    } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve sample users for $($sku.SkuPartNumber)" -OnError {
                        param($attempt, $max, $err)
                        $currentAttempt = if ($attempt) { $attempt } else { '?' }
                        $currentMax = if ($max) { $max } else { $maxAttempts }
                        Write-NCMessage "Failed to retrieve sample users for $($sku.SkuPartNumber), attempt $currentAttempt of $currentMax." -Level ERROR
                    }
                }
                catch {
                    $examples = @()
                }

                if ($examples) {
                    foreach ($entry in ($examples | Select-Object -First $sampleUserLimit)) {
                        if ($entry.UserPrincipalName) {
                            if ($entry.DisplayName) {
                                # $sampleUserList += ("{0} <{1}>" -f $entry.DisplayName, $entry.UserPrincipalName)
                                $sampleUserList += ("{0}" -f $entry.UserPrincipalName)
                            }
                            else {
                                $sampleUserList += $entry.UserPrincipalName
                            }
                        }
                    }
                }

                [pscustomobject][ordered]@{
                    Name            = $sku.Name
                    SkuPartNumber   = $sku.SkuPartNumber
                    SkuId           = $sku.SkuId
                    Total           = $sku.Total
                    TotalCount      = $sku.TotalCount
                    Consumed        = $sku.Consumed
                    Available       = $sku.Available
                    Enabled         = $sku.Enabled
                    Suspended       = $sku.Suspended
                    Warning         = $sku.Warning
                    Source          = $sku.Source
                    SampleUsers     = $sampleUserList
                    SampleUsersText = if ($sampleUserList.Count -gt 0) { $sampleUserList -join [Environment]::NewLine } else { $null }
                }
            }
        }

        if ($GridView.IsPresent) {
            $summaryRows = $sorted | Select-Object Name, SkuPartNumber, Total, Consumed, Available
            $summaryRows | Out-NCGridView -Title "M365 Tenant Licenses"

            if ($useSampleUsers) {
                $sampleRows = foreach ($sku in $sorted) {
                    if ($sku.PSObject.Properties['SampleUsers'] -and $sku.SampleUsers) {
                        foreach ($sampleUser in $sku.SampleUsers) {
                            [pscustomobject][ordered]@{
                                Name          = $sku.Name
                                SkuPartNumber = $sku.SkuPartNumber
                                SampleUser    = $sampleUser
                            }
                        }
                    }
                }

                if ($sampleRows) {
                    $sampleRows | Out-NCGridView -Title "M365 Tenant License Sample Users"
                }
            }
        }
        elseif ($AsTable.IsPresent) {
            $limited = $sorted | Select-Object @{
                Name       = 'Name'
                Expression = { Format-OutputString -Value $_.Name -MaxLength $NCVars.MaxFieldLength }
            }, SkuPartNumber, Total, Consumed, Available
            Show-Table -Rows $limited -AsTable

            if ($useSampleUsers) {
                Add-EmptyLine
                foreach ($sku in $sorted) {
                    if (-not ($sku.PSObject.Properties['SampleUsers'] -and $sku.SampleUsers)) {
                        continue
                    }

                    Write-NCMessage ("Sample users for {0} ({1}):" -f $sku.Name, $sku.SkuPartNumber) -Level INFO
                    foreach ($sampleUser in $sku.SampleUsers) {
                        Write-NCMessage ("  - {0}" -f $sampleUser) -Level INFO
                    }
                    Add-EmptyLine
                }
            }
        }
        else {
            $sorted
        }
    }
    finally {
        Restore-ProgressAndInfoPreferences
    }
}

function Get-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Shows licenses assigned to a specific user.
    .DESCRIPTION
        Downloads the license catalog, fetches the target user via Microsoft Graph, and prints each
        assigned SKU with the mapped product name (when available).
        When -Clipboard is specified, copies a quoted, comma-separated list of the licenses to the clipboard.
    .PARAMETER UserPrincipalName
        Target user UPN or object ID.
    .PARAMETER Clipboard
        Copies the resolved license names (fallback: SkuPartNumber) to the clipboard as: "License1","License2"
    .PARAMETER CheckAvailability
        Show available seat counts for the user's assigned SKUs (uses tenant license data).
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .PARAMETER ShowErrorDetails
        Kept for compatibility. Error messages always include the Microsoft Graph error detail, so this switch has no effect.
    .EXAMPLE
        Get-UserMsolAccountSku -UserPrincipalName "user@contoso.com"
    .EXAMPLE
        Get-UserMsolAccountSku -UserPrincipalName "user@contoso.com" -Clipboard
    .EXAMPLE
        Get-UserMsolAccountSku -UserPrincipalName "user@contoso.com" -CheckAvailability
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "User Principal Name (e.g. user@contoso.com)")]
        [Alias('User', 'UPN')]
        [string] $UserPrincipalName,
        [switch] $Clipboard,
        [switch] $CheckAvailability,
        [switch] $ForceLicenseCatalogRefresh,
        [switch] $ShowErrorDetails
    )

    begin {
        Set-ProgressAndInfoPreferences
        $initSucceeded = $true
        $clipboardLines = [System.Collections.Generic.List[string]]::new()
        $state = @{ Started = $false; ClipboardHasContent = $false }
        $queue = [System.Collections.Generic.List[object]]::new()
        $availabilityBySkuId = @{}

        try {
            $GraphConnection = Test-MgGraphConnection
            if (-not $GraphConnection) {
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                $initSucceeded = $false
            }
            else {
                $licenseCatalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent
            }
        }
        catch {
            Write-NCMessage $_ -Level ERROR
            $initSucceeded = $false
        }

        if ($initSucceeded) {
            $licenseLookup = $licenseCatalog.Lookup
            $customLookup = $licenseCatalog.CustomLookup
            $maxAttempts = 3

            $catalogSource = $licenseCatalog.Source
            $catalogUpdated = if ($licenseCatalog.LastCommitUtc) {
                Format-NCDateTime -Value $licenseCatalog.LastCommitUtc -AsLocalTime
            }
            else { $null }
            $catalogInfo = if ($catalogSource -or $catalogUpdated) {
                $parts = @()
                if ($catalogSource) { $parts += $catalogSource }
                if ($catalogUpdated) { $parts += "last updated: $catalogUpdated" }
                " (source: {0})" -f ($parts -join ', ')
            }
            else { '' }

            if ($CheckAvailability.IsPresent) {
                try {
                    $tenantSkus = Invoke-NCRetry -Action {
                        Get-MgSubscribedSku -All -ErrorAction Stop
                    } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve tenant licenses" -OnError {
                        param($attempt, $max, $err)
                        $currentAttempt = if ($attempt) { $attempt } else { '?' }
                        $currentMax = if ($max) { $max } else { $maxAttempts }
                        Write-NCMessage "Failed to retrieve tenant licenses, attempt $currentAttempt of $currentMax." -Level ERROR
                    }
                }
                catch {
                    $tenantSkus = @()
                }

                if ($tenantSkus -and $tenantSkus.Count -gt 0) {
                    foreach ($sku in $tenantSkus) {
                        $matchSource = $null
                        $display = Get-LicenseDisplayName -Lookup $licenseLookup `
                            -SkuPartNumber $sku.SkuPartNumber `
                            -FallbackLookup $customLookup `
                            -MatchSource ([ref]$matchSource)

                        $prepaid = $sku.PrepaidUnits
                        $enabled = if ($prepaid) { [int]$prepaid.Enabled } else { 0 }
                        $consumed = if ($sku.ConsumedUnits -is [int]) { [int]$sku.ConsumedUnits } else { [int]0 }
                        $available = [Math]::Max($enabled - $consumed, 0)
                        $skuKey = [string]$sku.SkuId

                        $availabilityBySkuId[$skuKey] = [pscustomobject]@{
                            Name          = if ($display) { $display } else { $sku.SkuPartNumber }
                            SkuPartNumber = $sku.SkuPartNumber
                            SkuId         = $sku.SkuId
                            Available     = $available
                            Source        = if ($matchSource) { $matchSource } elseif ($display) { 'Primary' } else { 'Unknown' }
                        }
                    }
                }
                else {
                    Write-NCMessage "Tenant license availability data not available." -Level WARNING
                }
            }
        }

        # Resolves the queued users, reads their licenses in Graph batches and prints the results.
        $flush = {
            $entries = @($queue)
            $queue.Clear()
            if ($entries.Count -eq 0) { return }

            if (-not $state.Started) {
                $state.Started = $true
                Write-NCGraphBatchNotice -Count ($entries.Count) -Noun 'users' -Streaming
            }

            $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $entries -Property @('id', 'displayName', 'userPrincipalName', 'mail') -FailedIdentifier $failedUsers

            $targets = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $entries) {
                $resolved = $resolvedUsers[$entry.Trim()]
                if (-not $resolved -or -not $resolved.id) {
                    if (-not $failedUsers.Contains($entry.Trim())) {
                        Write-NCMessage "Unable to resolve user recipient for $entry" -Level ERROR
                    }
                    continue
                }
                $targets.Add($resolved)
            }
            if ($targets.Count -eq 0) { return }

            $detailRequests = @(for ($i = 0; $i -lt $targets.Count; $i++) {
                    @{ Id = "l$i"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$targets[$i].id))/licenseDetails" }
                })
            $detailResults = @(Invoke-NCGraphBatchCollection -Requests $detailRequests -Activity 'Reading user licenses')

            for ($i = 0; $i -lt $targets.Count; $i++) {
                $user = $targets[$i]
                $detail = $detailResults[$i]

                Write-NCMessage ("`nProcessing user: {0} <{1}>{2}`n" -f $user.displayName, $user.userPrincipalName, $catalogInfo) -Level SUCCESS

                if (-not $detail.Success) {
                    Write-NCMessage "Failed to retrieve licenses for $($user.userPrincipalName) after $maxAttempts attempts." -Level ERROR
                    continue
                }
                $GraphLicense = if ($null -ne $detail.Items) { @($detail.Items) } else { @() }

                if ($GraphLicense -and $GraphLicense.Count -gt 0) {
                    $licensesForClipboard = @()
                    $availabilityLines = [System.Collections.Generic.List[string]]::new()

                    foreach ($lic in $GraphLicense) {
                        $skuPart = $lic.SkuPartNumber
                        $skuId = $lic.SkuId
                        $matchSource = $null
                        $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $skuPart -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
                        $licensesForClipboard += if ($display) { $display } else { $skuPart }
                        if ($display) {
                            $suffix = if ($matchSource -and $matchSource -ne 'Primary') { ' (custom)' } else { '' }
                            Write-NCMessage ("  - {0}{2} ({1})" -f $display, $skuId, $suffix) -Level INFO
                        }
                        else {
                            Write-Verbose ("  - Unknown license: {0} ({1})" -f $skuPart, $skuId)
                            Write-NCMessage ("  - {0} ({1})" -f $skuPart, $skuId) -Level WARNING
                        }

                        if ($CheckAvailability.IsPresent) {
                            $skuKey = [string]$skuId
                            $availability = if ($availabilityBySkuId.ContainsKey($skuKey)) { $availabilityBySkuId[$skuKey].Available } else { $null }
                            $availabilityDisplay = if ($null -ne $availability) { $availability } else { 'N/A' }
                            $availabilityLabel = if ($display) { $display } else { $skuPart }
                            $availabilityLines.Add(("  - {0}: {1}" -f $availabilityLabel, $availabilityDisplay)) | Out-Null
                        }
                    }

                    if ($CheckAvailability.IsPresent -and $availabilityLines.Count -gt 0) {
                        Add-EmptyLine
                        Write-NCMessage "Tenant License Availability:" -Level INFO
                        foreach ($line in $availabilityLines) {
                            Write-NCMessage $line -Level INFO
                        }
                    }

                    if ($Clipboard.IsPresent) {
                        $normalized = $licensesForClipboard | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                        $quoted = $normalized | ForEach-Object { "`"$($_.Replace('"', '\"'))`"" }
                        if ((@($quoted)).Count -gt 0) {
                            $clipboardLines.Add(($quoted -join ","))
                            $state.ClipboardHasContent = $true
                        }
                    }
                }
                else {
                    Write-Verbose "No licenses assigned to this user."
                    Write-NCMessage ("No licenses assigned to user {0}." -f $user.UserPrincipalName) -Level WARNING
                }
            }
        }
    }

    process {
        if (-not $initSucceeded) { return }

        $queue.Add($UserPrincipalName)
        if ($queue.Count -ge 20) {
            & $flush
        }
    }

    end {
        if ($initSucceeded) { & $flush }

        if ($Clipboard.IsPresent -and $state.ClipboardHasContent) {
            $clipboardText = ($clipboardLines -join [Environment]::NewLine)
            try {
                $clipboardText | Set-Clipboard
                Add-EmptyLine
                Write-NCMessage "Copied license list to clipboard." -Level INFO
            }
            catch {
                Write-NCMessage "Unable to copy license list to clipboard: $($_.Exception.Message)" -Level WARNING
            }
        }
        elseif ($Clipboard.IsPresent) {
            Add-EmptyLine
            Write-NCMessage "No license data available to copy to clipboard." -Level WARNING
        }

        Add-EmptyLine
        Restore-ProgressAndInfoPreferences
    }
}

function Get-UserUsageLocation {
    <#
    .SYNOPSIS
        Reads the current usage location for one or more users.
    .DESCRIPTION
        Resolves users from pipeline or explicit input and returns their current UsageLocation
        from Microsoft Graph. The output also includes the configured NebulaCore default so the
        current value can be compared against the environment setting at a glance.
    .PARAMETER UserPrincipalName
        User principal name, object ID, or short identifier. Accepts pipeline input.
    .EXAMPLE
        Get-UserUsageLocation -UserPrincipalName user@contoso.com
    .EXAMPLE
        'user1@contoso.com','user2@contoso.com' | Get-UserUsageLocation
    .EXAMPLE
        Get-MgUser -Filter "endsWith(userPrincipalName,'@contoso.com')" | Get-UserUsageLocation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($entry in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $targets.Add($entry.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if ($targets.Count -eq 0) {
                Write-NCMessage "No user principal names provided." -Level WARNING
                return
            }

            $GraphConnection = Test-MgGraphConnection -Scopes @('User.Read.All') -EnsureExchangeOnline:$false
            if (-not $GraphConnection) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            $defaultUsageLocation = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) {
                [string]$NCVars.UsageLocation
            }
            else {
                'US'
            }

            $normalizeUsageLocation = {
                param($value)
                if ([string]::IsNullOrWhiteSpace($value)) { return $null }
                return $value.Trim().ToUpperInvariant()
            }

            $queue = [System.Collections.Generic.List[string]]::new()
            $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($entry in $targets) {
                if ($dedup.Add($entry)) {
                    $queue.Add($entry) | Out-Null
                }
            }

            Write-NCGraphBatchNotice -Count ($queue.Count) -Noun 'user(s)'

            $counter = 0
            for ($offset = 0; $offset -lt $queue.Count; $offset += 20) {
                $chunk = @($queue[$offset..([Math]::Min($offset + 20, $queue.Count) - 1)])
                $Percentage = Get-NCProgressPercent -Current $counter -Total $queue.Count
                Write-Progress -Activity "Reading usage location" -Status "$counter of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $chunk -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation') -FailedIdentifier $failedUsers

                foreach ($upn in $chunk) {
                    $counter++
                    $user = $resolvedUsers[$upn.Trim()]
                    if (-not $user -or -not $user.id) {
                        if (-not $failedUsers.Contains($upn.Trim())) {
                            Write-NCMessage "Unable to resolve user recipient for $upn" -Level ERROR
                        }
                        continue
                    }

                    [pscustomobject]@{
                        UserPrincipalName              = $user.userPrincipalName
                        DisplayName                    = $user.displayName
                        UsageLocation                  = $user.usageLocation
                        ConfiguredDefaultUsageLocation = $defaultUsageLocation
                        MatchesConfiguredDefault       = (& $normalizeUsageLocation $user.usageLocation) -eq (& $normalizeUsageLocation $defaultUsageLocation)
                    }
                }
            }
        }
        finally {
            Write-Progress -Activity "Processing users" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Move-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Moves all licenses from one user to another.
    .DESCRIPTION
        Reads source user licenses (including disabled service plans), assigns them to the destination user,
        and then removes them from the source. Uses Microsoft Graph and the cached license catalog for friendly names.
    .PARAMETER SourceUserPrincipalName
        UserPrincipalName or object ID of the source user.
    .PARAMETER DestinationUserPrincipalName
        UserPrincipalName or object ID of the destination user.
    .EXAMPLE
        Move-MsolAccountSku -SourceUserPrincipalName user1@contoso.com -DestinationUserPrincipalName user2@contoso.com
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias('Source', 'From')]
        [string]$SourceUserPrincipalName,
        [Parameter(Mandatory = $true, Position = 1)]
        [Alias('Destination', 'To')]
        [string]$DestinationUserPrincipalName
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $resolvedUsers = Resolve-NCGraphUserBatch -Identifier @($SourceUserPrincipalName, $DestinationUserPrincipalName) -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation') -FailedIdentifier $failedUsers

        $sourceUser = $resolvedUsers[$SourceUserPrincipalName.Trim()]
        if (-not $sourceUser -or -not $sourceUser.id) {
            if (-not $failedUsers.Contains($SourceUserPrincipalName.Trim())) {
                Write-NCMessage "Unable to resolve source user recipient for $SourceUserPrincipalName" -Level ERROR
            }
            return
        }

        $destinationUser = $resolvedUsers[$DestinationUserPrincipalName.Trim()]
        if (-not $destinationUser -or -not $destinationUser.id) {
            if (-not $failedUsers.Contains($DestinationUserPrincipalName.Trim())) {
                Write-NCMessage "Unable to resolve destination user recipient for $DestinationUserPrincipalName" -Level ERROR
            }
            return
        }

        if ([string]$sourceUser.id -eq [string]$destinationUser.id) {
            Write-NCMessage "Source and destination users are the same. Aborting." -Level ERROR
            return
        }

        $defaultUsageLocation = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) {
            [string]$NCVars.UsageLocation
        }
        else { 'US' }

        $normalizeUsageLocation = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return $value.Trim().ToUpperInvariant()
        }

        $currentUsage = & $normalizeUsageLocation $destinationUser.usageLocation
        $desiredUsage = & $normalizeUsageLocation $defaultUsageLocation

        if ($desiredUsage -and $desiredUsage -ne $currentUsage) {
            $targetUsage = $defaultUsageLocation
            try {
                Update-MgUser -UserId $destinationUser.id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                Write-Verbose "Usage location set to $targetUsage for $($destinationUser.userPrincipalName)."
            }
            catch {
                Write-NCMessage "Unable to set usage location ($targetUsage) for $($destinationUser.userPrincipalName): $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        try {
            $licenseCatalog = Get-LicenseCatalog
        }
        catch {
            Write-NCMessage "License catalog unavailable: $($_.Exception.Message)" -Level WARNING
            $licenseCatalog = $null
        }

        $licenseLookup = $null
        $customLookup = $null
        if ($licenseCatalog) {
            if ($licenseCatalog.PSObject.Properties['Lookup']) {
                $licenseLookup = $licenseCatalog.Lookup
            }
            if ($licenseCatalog.PSObject.Properties['CustomLookup']) {
                $customLookup = $licenseCatalog.CustomLookup
            }
        }
        $maxAttempts = 3

        $licenseDetailRequests = @(
            @{ Id = 'l0'; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$sourceUser.id))/licenseDetails" }
            @{ Id = 'l1'; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$destinationUser.id))/licenseDetails" }
        )
        $licenseDetailResults = @(Invoke-NCGraphBatchCollection -Requests $licenseDetailRequests -Activity 'Reading user licenses')

        if (-not $licenseDetailResults[0].Success) {
            Write-NCMessage "Failed to retrieve licenses for $($sourceUser.userPrincipalName) after $maxAttempts attempts." -Level ERROR
            Write-Verbose $licenseDetailResults[0].ErrorMessage
            return
        }
        $sourceLicenses = if ($null -ne $licenseDetailResults[0].Items) { @($licenseDetailResults[0].Items) } else { @() }

        if ($sourceLicenses.Count -eq 0) {
            Write-NCMessage "Source user $($sourceUser.userPrincipalName) has no licenses to move." -Level WARNING
            return
        }

        if (-not $licenseDetailResults[1].Success) {
            Write-NCMessage "Unable to read destination licenses for $($destinationUser.userPrincipalName): $($licenseDetailResults[1].ErrorMessage)" -Level ERROR
            return
        }
        $destinationLicenses = if ($null -ne $licenseDetailResults[1].Items) { @($licenseDetailResults[1].Items) } else { @() }

        $destinationSkuIds = @()
        if ($destinationLicenses) {
            foreach ($sku in $destinationLicenses.SkuId) {
                $parsedSku = [guid]::Empty
                if ([guid]::TryParse([string]$sku, [ref]$parsedSku)) {
                    $destinationSkuIds += $parsedSku
                }
                elseif (-not [string]::IsNullOrWhiteSpace([string]$sku)) {
                    Write-Verbose "Skipping invalid destination SkuId '$sku'."
                }
            }
        }

        try {
            $tenantSkus = Invoke-NCRetry -Action {
                Get-MgSubscribedSku -All -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve tenant licenses" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve tenant licenses, attempt $currentAttempt of $currentMax." -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Unable to retrieve tenant licenses after $maxAttempts attempts." -Level ERROR
            return
        }

        $availabilityBySkuId = @{}
        foreach ($sku in $tenantSkus) {
            $skuIdString = [string]$sku.SkuId
            if ([string]::IsNullOrWhiteSpace($skuIdString)) { continue }
            $prepaidUnits = $sku.PrepaidUnits
            $enabledUnits = if ($prepaidUnits) { [int]$prepaidUnits.Enabled } else { 0 }
            $consumedUnits = if ($sku.ConsumedUnits -is [int]) { [int]$sku.ConsumedUnits } else { [int]0 }
            $availabilityBySkuId[$skuIdString.ToUpperInvariant()] = [Math]::Max($enabledUnits - $consumedUnits, 0)
        }

        $addLicenses = @()
        $removeSkuIds = @()
        $skippedInvalid = @()
        $namesNoAvailability = @()

        foreach ($lic in $sourceLicenses) {
            $skuIdString = [string]$lic.SkuId

            if ([string]::IsNullOrWhiteSpace($skuIdString)) {
                $skippedInvalid += "empty SkuId ($($lic.SkuPartNumber))"
                continue
            }

            $parsedGuid = [guid]::Empty
            if (-not [guid]::TryParse($skuIdString, [ref]$parsedGuid)) {
                $skippedInvalid += "invalid SkuId '$skuIdString' ($($lic.SkuPartNumber))"
                continue
            }

            if ($destinationSkuIds -contains $parsedGuid) {
                $removeSkuIds += $parsedGuid
                Write-Verbose "Destination already has $($lic.SkuPartNumber); skipping add."
                continue
            }

            $available = $availabilityBySkuId[$skuIdString.ToUpperInvariant()]
            if ($null -eq $available) { $available = 0 }
            if ($available -le 0) {
                Write-NCMessage ("No available units for license {0} (available: {1}); leaving it on the source." -f $lic.SkuPartNumber, $available) -Level WARNING
                $namesNoAvailability += $lic.SkuPartNumber
                continue
            }

            $removeSkuIds += $parsedGuid

            $validatedDisabled = @()
            if ($lic.DisabledPlans) {
                foreach ($plan in $lic.DisabledPlans) {
                    $planString = [string]$plan
                    if ([guid]::TryParse($planString, [ref]([guid]::Empty))) {
                        $validatedDisabled += $plan
                    }
                    elseif (-not [string]::IsNullOrWhiteSpace($planString)) {
                        Write-Verbose "Skipping invalid disabled plan '$planString' for $($lic.SkuPartNumber)."
                    }
                }
            }

            $addLicenses += @{
                SkuId         = $parsedGuid
                DisabledPlans = $validatedDisabled
            }
        }

        if ($addLicenses.Count -eq 0 -and $removeSkuIds.Count -eq 0) {
            Write-NCMessage "Nothing to move between $($sourceUser.userPrincipalName) and $($destinationUser.userPrincipalName)." -Level WARNING
            return
        }

        if ($skippedInvalid.Count -gt 0) {
            Write-NCMessage ("Skipped licenses with invalid IDs: {0}" -f ($skippedInvalid -join '; ')) -Level WARNING
        }

        if ($namesNoAvailability.Count -gt 0) {
            Write-NCMessage ("Left on source (no available units on destination tenant): {0}" -f (($namesNoAvailability | Select-Object -Unique) -join ', ')) -Level WARNING
        }

        $licenseNames = $sourceLicenses | ForEach-Object {
            $matchSource = $null
            $name = if ($licenseLookup) {
                Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $_.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
            }
            if ($name) { $name } else { $_.SkuPartNumber }
        }
        $actionSummary = "Move licenses ($($licenseNames -join ', ')) from $($sourceUser.userPrincipalName) to $($destinationUser.userPrincipalName)"

        if (-not $PSCmdlet.ShouldProcess($destinationUser.userPrincipalName, $actionSummary)) {
            return
        }

        if ($addLicenses.Count -gt 0) {
            try {
                Invoke-NCRetry -Action {
                    Set-MgUserLicense -UserId $destinationUser.id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
                } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "assign licenses to $($destinationUser.userPrincipalName)" -OnError {
                    param($attempt, $max, $err)
                    $currentAttempt = if ($attempt) { $attempt } else { '?' }
                    $currentMax = if ($max) { $max } else { $maxAttempts }
                    Write-NCMessage ("Failed to assign licenses to {0}, attempt {1} of {2}. {3}" -f $destinationUser.userPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
                } | Out-Null
                Write-NCMessage "Assigned licenses to $($destinationUser.userPrincipalName)." -Level SUCCESS
            }
            catch {
                Write-NCMessage "License assignment to $($destinationUser.userPrincipalName) failed. Aborting removal from source. $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        if ($removeSkuIds.Count -gt 0) {
            try {
                Invoke-NCRetry -Action {
                    Set-MgUserLicense -UserId $sourceUser.id -AddLicenses @() -RemoveLicenses ($removeSkuIds | Select-Object -Unique) -ErrorAction Stop
                } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "remove licenses from $($sourceUser.userPrincipalName)" -OnError {
                    param($attempt, $max, $err)
                    $currentAttempt = if ($attempt) { $attempt } else { '?' }
                    $currentMax = if ($max) { $max } else { $maxAttempts }
                    Write-NCMessage ("Failed to remove licenses from {0}, attempt {1} of {2}. {3}" -f $sourceUser.userPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
                } | Out-Null
                Write-NCMessage "Removed licenses from $($sourceUser.userPrincipalName)." -Level SUCCESS
            }
            catch {
                Write-NCMessage "Failed to remove licenses from $($sourceUser.userPrincipalName): $($_.Exception.Message)" -Level ERROR
            }
        }
    }
    finally {
        Restore-ProgressAndInfoPreferences
    }
}

function Remove-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Removes licenses from a user by friendly name or SKU identifier.
    .DESCRIPTION
        Resolves provided license names using the cached license catalog and the user's assigned licenses,
        then removes them from the target user. Accepts friendly product names, SKU part numbers, or SKU IDs.
    .PARAMETER UserPrincipalName
        Target user UPN or object ID.
    .PARAMETER License
        One or more license identifiers: friendly name (as resolved by the catalog), SKU part number, or SKU ID (GUID).
    .PARAMETER ForceLicenseCatalogRefresh
        Force a refresh of the cached license catalog before resolving friendly names.
    .PARAMETER ShowErrorDetails
        Kept for compatibility. Error messages always include the Microsoft Graph error detail, so this switch has no effect.
    .EXAMPLE
        Remove-UserMsolAccountSku -UserPrincipalName user@contoso.com -License "Microsoft 365 E3"
    .EXAMPLE
        Remove-UserMsolAccountSku -UserPrincipalName user@contoso.com -License "ENTERPRISEPACK","VISIOCLIENT"
    .EXAMPLE
        Remove-UserMsolAccountSku user@contoso.com -License "Exchange Online (Plan 2)"
    .EXAMPLE
        Remove-UserMsolAccountSku -UserPrincipalName user@contoso.com -License "18181a46-0d4e-45cd-891e-60aabd171b4e"
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'ByLicense')]
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'All')]
        [Alias('User', 'UPN')]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true, ParameterSetName = 'ByLicense')]
        [string[]]$License,
        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [switch]$All,
        [Parameter(ParameterSetName = 'ByLicense')]
        [Parameter(ParameterSetName = 'All')]
        [switch]$ForceLicenseCatalogRefresh,
        [switch]$ShowErrorDetails
    )

    begin {
        Set-ProgressAndInfoPreferences

        $graphConnected = Test-MgGraphConnection
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $queue = [System.Collections.Generic.List[object]]::new()
        $state = @{ Started = $false }
        $maxAttempts = 3

        $licenseLookup = $null
        $customLookup = $null
        if ($graphConnected) {
            try {
                $licenseCatalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent
            }
            catch {
                Write-NCMessage $_ -Level WARNING
                $licenseCatalog = $null
            }

            if ($licenseCatalog) {
                if ($licenseCatalog.PSObject.Properties['Lookup']) { $licenseLookup = $licenseCatalog.Lookup }
                if ($licenseCatalog.PSObject.Properties['CustomLookup']) { $customLookup = $licenseCatalog.CustomLookup }
            }
        }

        $normalizeString = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return ($value.Trim().ToUpperInvariant())
        }

        # Resolves, reads licenses for and updates the queued users in Graph batches.
        $flush = {
            $entries = @($queue)
            $queue.Clear()
            if ($entries.Count -eq 0) { return }

            if (-not $state.Started) {
                $state.Started = $true
                Write-NCGraphBatchNotice -Count ($entries.Count) -Noun 'users' -Streaming
            }

            # (a) Resolve every queued user.
            $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $entries -Property @('id', 'userPrincipalName', 'displayName') -FailedIdentifier $failedUsers

            $targets = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $entries) {
                $user = $resolvedUsers[$entry.Trim()]
                if (-not $user -or -not $user.id) {
                    if (-not $failedUsers.Contains($entry.Trim())) {
                        Write-NCMessage "Unable to resolve user recipient for $entry" -Level ERROR
                    }
                    continue
                }
                $targets.Add([pscustomobject]@{ Id = [string]$user.id; UserPrincipalName = [string]$user.userPrincipalName })
            }
            if ($targets.Count -eq 0) { return }

            # (b) Read assigned licenses for every resolved user.
            $detailRequests = @(for ($i = 0; $i -lt $targets.Count; $i++) {
                    @{ Id = "l$i"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString($targets[$i].Id))/licenseDetails" }
                })
            $detailLookup = @{}
            foreach ($result in @(Invoke-NCGraphBatchCollection -Requests $detailRequests -Activity 'Reading user licenses')) {
                $detailLookup[$result.Id] = $result
            }

            # (c) Work out what to remove per user and ask for confirmation.
            $approved = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $targets.Count; $i++) {
                $user = $targets[$i]
                $detail = $detailLookup["l$i"]
                if (-not $detail.Success) {
                    Write-NCMessage "Unable to retrieve licenses for $($user.UserPrincipalName) after $maxAttempts attempts." -Level ERROR
                    continue
                }

                $assignedLicenses = @($detail.Items | ForEach-Object {
                        [pscustomobject]@{ SkuId = [string]$_.skuId; SkuPartNumber = [string]$_.skuPartNumber }
                    })

                if ($assignedLicenses.Count -eq 0) {
                    Write-NCMessage "User $($user.UserPrincipalName) has no licenses to remove." -Level WARNING
                    continue
                }

                $licenseNames = @()
                $removeLicenseIds = @()

                if ($PSCmdlet.ParameterSetName -eq 'All') {
                    foreach ($lic in $assignedLicenses) {
                        if (-not $lic.SkuId) { continue }
                        $removeLicenseIds += $lic.SkuId

                        $matchSource = $null
                        $display = $null
                        if ($licenseLookup) {
                            $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $lic.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
                        }
                        $licenseNames += if ($display) { $display } else { $lic.SkuPartNumber }
                    }

                    $removeLicenseIds = @($removeLicenseIds | Where-Object { $_ } | Select-Object -Unique)
                    $licenseNames = @($licenseNames | Where-Object { $_ } | Select-Object -Unique)

                    if ($removeLicenseIds.Count -eq 0) {
                        Write-NCMessage "No licenses to remove for $($user.UserPrincipalName)." -Level WARNING
                        continue
                    }
                }
                else {
                    $resolved = @()
                    $unmatched = @()
                    $inputLicenses = $License | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique

                    foreach ($entry in $inputLicenses) {
                        $target = & $normalizeString $entry
                        $match = $null

                        foreach ($lic in $assignedLicenses) {
                            $skuIdString = [string]$lic.SkuId
                            $skuPart = & $normalizeString $lic.SkuPartNumber

                            $matchSource = $null
                            $display = $null
                            if ($licenseLookup) {
                                $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $lic.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
                            }
                            $displayNormalized = if ($display) { & $normalizeString $display } else { $null }

                            if ($target -eq $skuPart -or $target -eq ($skuIdString.ToUpperInvariant()) -or ($displayNormalized -and $target -eq $displayNormalized)) {
                                $match = @{
                                    SkuId         = $lic.SkuId
                                    SkuPartNumber = $lic.SkuPartNumber
                                    Name          = if ($display) { $display } else { $lic.SkuPartNumber }
                                }
                                break
                            }
                        }

                        if ($match) {
                            $resolved += $match
                        }
                        else {
                            $unmatched += $entry
                        }
                    }

                    if ($unmatched.Count -gt 0) {
                        Write-NCMessage ("Unable to resolve license(s) for removal: {0}" -f ($unmatched -join ', ')) -Level ERROR
                        continue
                    }

                    $removeLicenses = @($resolved | Group-Object SkuId | ForEach-Object {
                            $_.Group | Select-Object -First 1
                        })

                    if ($removeLicenses.Count -eq 0) {
                        Write-NCMessage "No licenses matched for removal." -Level ERROR
                        continue
                    }

                    $licenseNames = @($removeLicenses | ForEach-Object { $_.Name } | Select-Object -Unique)
                    $removeLicenseIds = @($removeLicenses | ForEach-Object { [string]$_.SkuId })
                }

                $summary = if ($PSCmdlet.ParameterSetName -eq 'All') {
                    "Remove all license(s): {0} from {1}" -f ($licenseNames -join ', '), $user.UserPrincipalName
                }
                else {
                    "Remove license(s): {0} from {1}" -f ($licenseNames -join ', '), $user.UserPrincipalName
                }

                if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, $summary)) {
                    continue
                }

                $approved.Add([pscustomobject]@{
                        Id               = $user.Id
                        UserPrincipalName = $user.UserPrincipalName
                        LicenseNames     = $licenseNames
                        RemoveLicenseIds = $removeLicenseIds
                    })
            }
            if ($approved.Count -eq 0) { return }

            # (d) Remove the licenses.
            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{
                        Id     = "m$i"
                        Method = 'POST'
                        Url    = "/users/$([uri]::EscapeDataString($approved[$i].Id))/assignLicense"
                        Body   = @{ addLicenses = @(); removeLicenses = @($approved[$i].RemoveLicenseIds) }
                    }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity 'Removing licenses')
            for ($i = 0; $i -lt $approved.Count; $i++) {
                $response = $responses[$i]
                if ($response.Success) {
                    Write-NCMessage ("Removed license(s) from {0}: {1}" -f $approved[$i].UserPrincipalName, ($approved[$i].LicenseNames -join ', ')) -Level SUCCESS
                }
                else {
                    Write-NCMessage "License removal failed for $($approved[$i].UserPrincipalName): $($response.ErrorMessage)" -Level ERROR
                }
            }
        }
    }

    process {
        if (-not $graphConnected) {
            return
        }

        $queue.Add($UserPrincipalName)
        if ($queue.Count -ge 20) {
            & $flush
        }
    }

    end {
        try {
            if ($graphConnected) {
                & $flush
            }
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Set-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Adds and removes user licenses together in one Microsoft Graph request.
    .DESCRIPTION
        Accepts friendly product names, SKU part numbers, or SKU IDs. Collects pipeline records,
        merges them by resolved user ID, and validates the complete change before writing.
        Each user is updated independently; changes across users are not transactional.
        Existing service-plan selections on retained licenses are preserved.
    .PARAMETER UserPrincipalName
        Target user UPN or object ID. Accepts pipeline values and properties.
    .PARAMETER AddLicense
        Licenses to add. Accepts pipeline properties.
    .PARAMETER RemoveLicense
        Licenses to remove. Accepts pipeline properties. At least one operation is required.
    .PARAMETER ForceLicenseCatalogRefresh
        Force a refresh of the cached license catalog before resolving friendly names.
    .PARAMETER ShowErrorDetails
        Kept for compatibility. Graph errors always include details.
    .PARAMETER QuietNoChanges
        Suppress per-user messages when the requested licenses are already in the desired state.
    .EXAMPLE
        Set-UserMsolAccountSku user@contoso.com -AddLicense 'NEWPACK' -RemoveLicense 'ENTERPRISEPACK' -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN')]
        [string]$UserPrincipalName,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$AddLicense,
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$RemoveLicense,
        [switch]$ForceLicenseCatalogRefresh,
        [switch]$ShowErrorDetails,
        [switch]$QuietNoChanges
    )

    begin {
        Set-ProgressAndInfoPreferences
        $graphConnected = Test-MgGraphConnection -Scopes @('User.Read.All', 'User.ReadWrite.All', 'LicenseAssignment.Read.All', 'LicenseAssignment.ReadWrite.All')
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }
        $queue = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if ($graphConnected) {
            # Snapshot the per-record arrays; do not write until all aliases and rows have been merged.
            $queue.Add([pscustomobject]@{
                    Identifier = $UserPrincipalName
                    Add = @($AddLicense | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    Remove = @($RemoveLicense)
                })
        }
    }

    end {
        $batchResults = @{}
        $shouldSummarize = $false
        $recordResult = {
            param([string]$Key, [string]$UserPrincipalName, [string]$Status, [string]$Reason)
            $batchResults[$Key] = [pscustomobject]@{
                UserPrincipalName = $UserPrincipalName
                Status = $Status
                Reason = $Reason
            }
        }
        try {
            if (-not $graphConnected -or $queue.Count -eq 0) { return }
            $shouldSummarize = $true
            Write-NCGraphBatchNotice -Count $queue.Count -Noun 'users'
            $licenseLookup = $null
            $customLookup = $null
            try {
                $catalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent
                if ($catalog) {
                    $licenseLookup = $catalog.Lookup
                    $customLookup = $catalog.CustomLookup
                }
            }
            catch { Write-NCMessage $_ -Level WARNING }

            $tenantSkus = @()
            if (@($queue | Where-Object { $_.Add.Count -gt 0 }).Count -gt 0) {
                try {
                    $tenantSkus = @(Invoke-NCRetry -Action { Get-MgSubscribedSku -All -ErrorAction Stop } -MaxAttempts 3 -DelaySeconds 5 -OperationDescription 'retrieve tenant licenses' -OnError {
                            param($attempt, $max, $err)
                            Write-NCMessage "Failed to retrieve tenant licenses, attempt $attempt of $max." -Level ERROR
                        })
                }
                catch { Write-NCMessage 'Unable to retrieve tenant licenses after 3 attempts.' -Level ERROR }
            }

            $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolvedUsers = Resolve-NCGraphUserBatch -Identifier @($queue.Identifier) -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation', 'licenseAssignmentStates') -FailedIdentifier $failedUsers
            $byId = @{}
            $targets = [System.Collections.Generic.List[object]]::new()
            foreach ($row in $queue) {
                $identifier = $row.Identifier.Trim()
                $user = $resolvedUsers[$identifier]
                if (-not $user -or -not $user.id) {
                    if (-not $failedUsers.Contains($identifier)) { Write-NCMessage "Unable to resolve user recipient for $identifier" -Level ERROR }
                    & $recordResult "unresolved:$identifier" $identifier 'Failed' 'Unable to resolve user.'
                    continue
                }
                $id = [string]$user.id
                if (-not $byId.ContainsKey($id)) {
                    $target = [pscustomobject]@{
                        Id = $id; UserPrincipalName = [string]$user.userPrincipalName
                        AssignmentStates = $user.licenseAssignmentStates
                        UsageLocation = [string]$user.usageLocation
                        Add = [System.Collections.Generic.List[string]]::new()
                        Remove = [System.Collections.Generic.List[string]]::new()
                        EffectiveAdds = @()
                        EffectiveRemoves = @()
                        AlreadyAssignedAdds = @()
                        AlreadyAbsentRemovals = @()
                        FailureReasons = [System.Collections.Generic.List[string]]::new()
                        Invalid = $false
                    }
                    $byId[$id] = $target
                    $targets.Add($target)
                    & $recordResult "user:$id" $target.UserPrincipalName 'Pending' ''
                }
                $target = $byId[$id]
                $adds = @($row.Add | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $removes = @($row.Remove | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($adds.Count + $removes.Count -eq 0) {
                    Write-NCMessage "No license changes requested for $identifier." -Level ERROR
                    $target.Invalid = $true
                    $target.FailureReasons.Add('No license changes were specified.')
                }
                foreach ($entry in $adds) { $target.Add.Add($entry.Trim()) }
                foreach ($entry in $removes) { $target.Remove.Add($entry.Trim()) }
            }
            if ($targets.Count -eq 0) { return }

            $detailRequests = @(for ($i = 0; $i -lt $targets.Count; $i++) {
                    @{ Id = "l$i"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString($targets[$i].Id))/licenseDetails" }
                })
            $details = @{}
            foreach ($result in @(Invoke-NCGraphBatchCollection -Requests $detailRequests -Activity 'Reading user licenses')) { $details[$result.Id] = $result }

            $requests = [System.Collections.Generic.List[object]]::new()
            $approved = @{}
            $tenantSkuLookupAttempted = $tenantSkus.Count -gt 0
            for ($i = 0; $i -lt $targets.Count; $i++) {
                $target = $targets[$i]
                $detail = $details["l$i"]
                if (-not $detail.Success) {
                    Write-NCMessage "Unable to retrieve licenses for $($target.UserPrincipalName): $($detail.ErrorMessage)" -Level ERROR
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' ([string]$detail.ErrorMessage)
                    continue
                }
                if ($target.Invalid) {
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' ($target.FailureReasons -join '; ')
                    continue
                }
                # licenseDetails includes group-inherited licenses as well as direct assignments.
                # Require the explicit assignment source before deciding whether an addition is redundant.
                $assignmentStateComplete = $target.AssignmentStates -is [array]
                $assignedSkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $directSkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($state in $target.AssignmentStates) {
                    $assignment = [pscustomobject]$state
                    if (-not $assignment -or [string]::IsNullOrWhiteSpace([string]$assignment.skuId) -or
                        -not $assignment.PSObject.Properties['assignedByGroup'] -or
                        -not $assignment.PSObject.Properties['disabledPlans'] -or
                        $assignment.disabledPlans -isnot [array]) {
                        $assignmentStateComplete = $false
                        continue
                    }
                    $null = $assignedSkuIds.Add([string]$assignment.skuId)
                    if ($null -eq $assignment.assignedByGroup) { $null = $directSkuIds.Add([string]$assignment.skuId) }
                }
                foreach ($license in $detail.Items) {
                    if (-not $assignedSkuIds.Contains([string]$license.skuId)) { $assignmentStateComplete = $false }
                }
                if (-not $assignmentStateComplete) {
                    Write-NCMessage "Unable to read complete license assignment state for $($target.UserPrincipalName)." -Level ERROR
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' 'Unable to read complete license assignment state.'
                    continue
                }
                $resolved = @{ Add = @{}; Remove = @{}; AlreadyAbsentRemove = @{} }
                $presentSkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($license in $detail.Items) { $null = $presentSkuIds.Add([string]$license.skuId) }
                foreach ($operation in @('Add', 'Remove')) {
                    $candidates = if ($operation -eq 'Add') { $tenantSkus } else { @($detail.Items) }
                    foreach ($entry in $target.$operation) {
                        $match = $null
                        foreach ($sku in $candidates) {
                            $display = $null
                            $source = $null
                            if ($licenseLookup) {
                                $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $sku.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$source)
                            }
                            if ($entry -ieq [string]$sku.SkuId -or $entry -ieq [string]$sku.SkuPartNumber -or ($display -and $entry -ieq $display)) {
                                $match = $sku
                                break
                            }
                        }
                        if (-not $match -and $operation -eq 'Remove') {
                            if (-not $tenantSkuLookupAttempted) {
                                $tenantSkuLookupAttempted = $true
                                try {
                                    $tenantSkus = @(Invoke-NCRetry -Action { Get-MgSubscribedSku -All -ErrorAction Stop } -MaxAttempts 3 -DelaySeconds 5 -OperationDescription 'retrieve tenant licenses' -OnError {
                                            param($attempt, $max, $err)
                                            Write-NCMessage "Failed to retrieve tenant licenses, attempt $attempt of $max." -Level ERROR
                                        })
                                }
                                catch { Write-NCMessage 'Unable to retrieve tenant licenses after 3 attempts.' -Level ERROR }
                            }
                            foreach ($sku in $tenantSkus) {
                                $display = $null
                                $source = $null
                                if ($licenseLookup) {
                                    $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $sku.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$source)
                                }
                                if ($entry -ieq [string]$sku.SkuId -or $entry -ieq [string]$sku.SkuPartNumber -or ($display -and $entry -ieq $display)) {
                                    $match = $sku
                                    break
                                }
                            }
                        }
                        if (-not $match -or -not $match.SkuId) {
                            Write-NCMessage "Unable to resolve license '$entry' for $operation on $($target.UserPrincipalName)." -Level ERROR
                            $target.Invalid = $true
                            $target.FailureReasons.Add("Unable to resolve license '$entry' for $operation.")
                        }
                        elseif ($operation -eq 'Remove' -and -not $presentSkuIds.Contains([string]$match.SkuId)) {
                            $resolved.AlreadyAbsentRemove[[string]$match.SkuId] = $entry
                        }
                        else { $resolved[$operation][[string]$match.SkuId] = $entry }
                    }
                }
                $overlap = @($resolved.Add.Keys | Where-Object { $resolved.Remove.ContainsKey($_) })
                if ($overlap.Count -gt 0) {
                    Write-NCMessage "Contradictory license additions and removals for $($target.UserPrincipalName): $($overlap -join ', ')" -Level ERROR
                    $target.Invalid = $true
                    $target.FailureReasons.Add("Contradictory license additions and removals: $($overlap -join ', ').")
                }
                if ($target.Invalid) {
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' ($target.FailureReasons -join '; ')
                    continue
                }

                # Omit only existing direct assignments to preserve their disabled service plans.
                # A group-inherited SKU still needs the direct addition the caller requested.
                $addIds = @($resolved.Add.Keys | Where-Object { -not $directSkuIds.Contains($_) })
                $removeIds = @($resolved.Remove.Keys)
                $target.EffectiveAdds = @($addIds | ForEach-Object { $resolved.Add[$_] })
                $target.EffectiveRemoves = @($removeIds | ForEach-Object { $resolved.Remove[$_] })
                $target.AlreadyAssignedAdds = @($resolved.Add.Keys | Where-Object { $directSkuIds.Contains($_) } | ForEach-Object { $resolved.Add[$_] })
                $target.AlreadyAbsentRemovals = @($resolved.AlreadyAbsentRemove.Values)
                $inheritedRemovals = @($removeIds | Where-Object { -not $directSkuIds.Contains($_) })
                if ($inheritedRemovals.Count -gt 0) {
                    $inheritedNames = @($inheritedRemovals | ForEach-Object { $resolved.Remove[$_] })
                    Write-NCMessage ("Cannot remove inherited-only license(s) for {0}: {1}. Remove the license assignment from its group." -f $target.UserPrincipalName, ($inheritedNames -join ', ')) -Level ERROR
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' ("Cannot remove inherited-only license(s): {0}." -f ($inheritedNames -join ', '))
                    continue
                }
                $targetUsage = $null
                if ($addIds.Count -gt 0 -and [string]::IsNullOrWhiteSpace($target.UsageLocation)) {
                    $targetUsage = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) { [string]$NCVars.UsageLocation } else { 'US' }
                }
                if ($addIds.Count + $removeIds.Count -eq 0) {
                    $stateSummary = @()
                    if ($target.AlreadyAssignedAdds.Count -gt 0) { $stateSummary += "already assigned: $($target.AlreadyAssignedAdds -join ', ')" }
                    if ($target.AlreadyAbsentRemovals.Count -gt 0) { $stateSummary += "already absent: $($target.AlreadyAbsentRemovals -join ', ')" }
                    if (-not $QuietNoChanges) {
                        Write-NCMessage ("No license changes needed for {0}; {1}." -f $target.UserPrincipalName, ($stateSummary -join '; ')) -Level SUCCESS
                    }
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Unchanged' ($stateSummary -join '; ')
                    continue
                }
                $summary = "Add license(s): {0}; remove license(s): {1}" -f ($target.EffectiveAdds -join ', '), ($target.EffectiveRemoves -join ', ')
                if ($target.AlreadyAssignedAdds.Count -gt 0) { $summary += "; already assigned: $($target.AlreadyAssignedAdds -join ', ')" }
                if ($target.AlreadyAbsentRemovals.Count -gt 0) { $summary += "; already absent: $($target.AlreadyAbsentRemovals -join ', ')" }
                if ($targetUsage) { $summary = "Set usage location to {0}; {1}" -f $targetUsage, $summary }
                if (-not $PSCmdlet.ShouldProcess($target.UserPrincipalName, $summary)) {
                    $reason = if ($WhatIfPreference) { 'Preview only (-WhatIf).' } else { 'Not approved.' }
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Skipped' $reason
                    continue
                }
                $requestId = "m$i"
                $approved[$requestId] = $target
                if ($targetUsage) {
                    $patchId = "u$i"
                    $approved[$patchId] = $target
                    $requests.Add(@{
                            Id = $patchId; Method = 'PATCH'; Url = "/users/$([uri]::EscapeDataString($target.Id))"
                            Body = @{ usageLocation = $targetUsage }
                        })
                }
                $requests.Add(@{
                        Id = $requestId; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString($target.Id))/assignLicense"
                        Body = @{
                            addLicenses = @(foreach ($id in $addIds) { @{ skuId = $id; disabledPlans = @() } })
                            removeLicenses = @($removeIds)
                        }
                    })
            }
            if ($requests.Count -eq 0) { return }
            $patches = @($requests | Where-Object { $_.Method -eq 'PATCH' })
            $readyUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($licenseRequest in @($requests | Where-Object { $_.Method -eq 'POST' })) {
                $null = $readyUsers.Add([string]$licenseRequest.Id.Substring(1))
            }
            if ($patches.Count -gt 0) {
                foreach ($response in @(Invoke-NCGraphBatch -Requests $patches -Activity 'Setting usage location')) {
                    $target = $approved[$response.Id]
                    $key = [string]$response.Id.Substring(1)
                    if (-not $response.Success) {
                        $null = $readyUsers.Remove($key)
                        Write-NCMessage "Unable to set usage location for $($target.UserPrincipalName): $($response.ErrorMessage)" -Level ERROR
                        & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' ("Unable to set usage location: {0}" -f $response.ErrorMessage)
                    }
                }
            }
            $licenseRequests = @($requests | Where-Object { $_.Method -eq 'POST' -and $readyUsers.Contains($_.Id.Substring(1)) })
            foreach ($response in @(Invoke-NCGraphBatch -Requests $licenseRequests -Activity 'Changing user licenses')) {
                $target = $approved[$response.Id]
                if ($response.Success) {
                    $changeSummary = @()
                    if ($target.EffectiveAdds.Count -gt 0) { $changeSummary += "added: $($target.EffectiveAdds -join ', ')" }
                    if ($target.EffectiveRemoves.Count -gt 0) { $changeSummary += "removed: $($target.EffectiveRemoves -join ', ')" }
                    if ($target.AlreadyAssignedAdds.Count -gt 0) { $changeSummary += "already assigned: $($target.AlreadyAssignedAdds -join ', ')" }
                    if ($target.AlreadyAbsentRemovals.Count -gt 0) { $changeSummary += "already absent: $($target.AlreadyAbsentRemovals -join ', ')" }
                    Write-NCMessage ("Changed licenses for {0} ({1})." -f $target.UserPrincipalName, ($changeSummary -join '; ')) -Level SUCCESS
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Succeeded' ($changeSummary -join '; ')
                }
                else {
                    Write-NCMessage "License change failed for $($target.UserPrincipalName): $($response.ErrorMessage)" -Level ERROR
                    & $recordResult "user:$($target.Id)" $target.UserPrincipalName 'Failed' ([string]$response.ErrorMessage)
                }
            }
        }
        finally {
            if ($shouldSummarize) {
                foreach ($key in @($batchResults.Keys)) {
                    if ($batchResults[$key].Status -eq 'Pending') {
                        & $recordResult $key $batchResults[$key].UserPrincipalName 'Failed' 'Operation did not complete.'
                    }
                }
                $results = @($batchResults.Values)
                $succeededCount = @($results | Where-Object Status -eq 'Succeeded').Count
                $failedCount = @($results | Where-Object Status -eq 'Failed').Count
                $unchangedCount = @($results | Where-Object Status -eq 'Unchanged').Count
                $skippedCount = @($results | Where-Object Status -eq 'Skipped').Count
                Write-NCMessage ("License batch totals — changed successfully: {0}; failed: {1}; unchanged: {2}; skipped/not executed: {3}." -f $succeededCount, $failedCount, $unchangedCount, $skippedCount) -Level INFO
                if ($failedCount -gt 0) {
                    $failedDetails = @($results | Where-Object Status -eq 'Failed' | ForEach-Object { "{0} ({1})" -f $_.UserPrincipalName, $_.Reason }) -join '; '
                    Write-NCMessage "Failed users: $failedDetails" -Level ERROR
                }
            }
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Swap-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Swaps explicitly selected direct licenses between two users without requiring spare tenant seats.
    .DESCRIPTION
        Validates both users and their selected direct license assignments, removes the selected licenses
        from both users, and only then assigns each user's licenses to the other. If a phase fails, it
        attempts to restore the original assignments and reports any incomplete recovery.
        Use -AllLicenses to select all direct licenses automatically; -ExcludeLicense leaves the named
        license assigned to its current user. Explicit license selections remain available otherwise.
    .PARAMETER FirstUserPrincipalName
        UPN or object ID of the first user.
    .PARAMETER FirstUserLicense
        One or more licenses currently assigned directly to the first user.
    .PARAMETER SecondUserPrincipalName
        UPN or object ID of the second user.
    .PARAMETER SecondUserLicense
        One or more licenses currently assigned directly to the second user.
    .PARAMETER AllLicenses
        Automatically swap all direct licenses currently assigned to both users.
    .PARAMETER ExcludeLicense
        In -AllLicenses mode, leave these license(s) assigned to their current user.
    .EXAMPLE
        Swap-UserMsolAccountSku -FirstUserPrincipalName user1@contoso.com -FirstUserLicense ENTERPRISEPACK -SecondUserPrincipalName user2@contoso.com -SecondUserLicense SPE_E3
    .EXAMPLE
        Swap-UserMsolAccountSku -FirstUserPrincipalName user1@contoso.com -SecondUserPrincipalName user2@contoso.com -AllLicenses -ExcludeLicense 'FABRIC_FREE'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$FirstUserPrincipalName,
        [Parameter(Position = 1)]
        [string[]]$FirstUserLicense,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$SecondUserPrincipalName,
        [Parameter(Position = 3)]
        [string[]]$SecondUserLicense,
        [switch]$AllLicenses,
        [string[]]$ExcludeLicense,
        [switch]$ForceLicenseCatalogRefresh
    )

    Set-ProgressAndInfoPreferences
    try {
        $graphConnected = Test-MgGraphConnection -Scopes @('User.Read.All', 'User.ReadWrite.All', 'LicenseAssignment.Read.All', 'LicenseAssignment.ReadWrite.All')
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $firstSelections = @($FirstUserLicense | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
        $secondSelections = @($SecondUserLicense | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
        $excludeSelections = @($ExcludeLicense | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
        if ($AllLicenses -and ($firstSelections.Count -gt 0 -or $secondSelections.Count -gt 0)) {
            Write-NCMessage 'Use either -AllLicenses or explicit license selections, not both.' -Level ERROR
            return
        }
        if (-not $AllLicenses -and $excludeSelections.Count -gt 0) {
            Write-NCMessage '-ExcludeLicense is available only with -AllLicenses.' -Level ERROR
            return
        }
        if (-not $AllLicenses -and ($firstSelections.Count -eq 0 -or $secondSelections.Count -eq 0)) {
            Write-NCMessage 'Specify at least one license for each user to swap.' -Level ERROR
            return
        }

        $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $resolvedUsers = Resolve-NCGraphUserBatch -Identifier @($FirstUserPrincipalName, $SecondUserPrincipalName) -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation', 'licenseAssignmentStates') -FailedIdentifier $failedUsers -Activity 'Resolving swap users'
        $firstUser = $resolvedUsers[$FirstUserPrincipalName.Trim()]
        $secondUser = $resolvedUsers[$SecondUserPrincipalName.Trim()]
        if (-not $firstUser -or -not $firstUser.id) {
            if (-not $failedUsers.Contains($FirstUserPrincipalName.Trim())) { Write-NCMessage "Unable to resolve first user '$FirstUserPrincipalName'." -Level ERROR }
            return
        }
        if (-not $secondUser -or -not $secondUser.id) {
            if (-not $failedUsers.Contains($SecondUserPrincipalName.Trim())) { Write-NCMessage "Unable to resolve second user '$SecondUserPrincipalName'." -Level ERROR }
            return
        }
        if ([string]$firstUser.id -eq [string]$secondUser.id) {
            Write-NCMessage 'Both identifiers resolve to the same user; no swap was performed.' -Level ERROR
            return
        }

        $catalog = $null
        try { $catalog = Get-LicenseCatalog -IncludeMetadata -ForceRefresh:$ForceLicenseCatalogRefresh.IsPresent }
        catch { Write-NCMessage "License catalog unavailable: $($_.Exception.Message)" -Level WARNING }
        $licenseLookup = if ($catalog -and $catalog.PSObject.Properties['Lookup']) { $catalog.Lookup } else { $null }
        $customLookup = if ($catalog -and $catalog.PSObject.Properties['CustomLookup']) { $catalog.CustomLookup } else { $null }

        $detailRequests = @(
            @{ Id = 's0'; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))/licenseDetails" }
            @{ Id = 's1'; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))/licenseDetails" }
        )
        $detailResults = @(Invoke-NCGraphBatchCollection -Requests $detailRequests -Activity 'Reading swap user licenses')
        if ($detailResults.Count -lt 2 -or -not $detailResults[0].Success -or -not $detailResults[1].Success) {
            $failedDetail = @($detailResults | Where-Object { -not $_.Success } | ForEach-Object { $_.ErrorMessage }) -join '; '
            Write-NCMessage "Unable to read both users' license details: $failedDetail" -Level ERROR
            return
        }

        if ($AllLicenses) {
            $excludedSkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $allUserLicenses = @($detailResults[0].Items) + @($detailResults[1].Items)
            foreach ($entry in $excludeSelections) {
                $match = $null
                foreach ($license in $allUserLicenses) {
                    $display = $null
                    $source = $null
                    if ($licenseLookup) { $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $license.skuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$source) }
                    if ($entry -ieq [string]$license.skuId -or $entry -ieq [string]$license.skuPartNumber -or ($display -and $entry -ieq $display)) {
                        $match = $license
                        break
                    }
                }
                if (-not $match) {
                    Write-NCMessage "Excluded license '$entry' is not assigned to either user." -Level ERROR
                    return
                }
                $null = $excludedSkuIds.Add([string]$match.skuId)
            }
            $firstSelections = @($detailResults[0].Items | Where-Object { -not $excludedSkuIds.Contains([string]$_.skuId) } | ForEach-Object { [string]$_.skuId })
            $secondSelections = @($detailResults[1].Items | Where-Object { -not $excludedSkuIds.Contains([string]$_.skuId) } | ForEach-Object { [string]$_.skuId })
            if ($firstSelections.Count -eq 0 -or $secondSelections.Count -eq 0) {
                Write-NCMessage 'After exclusions, each user must have at least one license to swap.' -Level ERROR
                return
            }
        }

        $resolveSelections = {
            param($User, $Items, [string[]]$Selections)
            $states = @{}
            $stateComplete = $User.licenseAssignmentStates -is [array]
            foreach ($state in $User.licenseAssignmentStates) {
                $assignment = [pscustomobject]$state
                if (-not $assignment -or [string]::IsNullOrWhiteSpace([string]$assignment.skuId) -or
                    -not $assignment.PSObject.Properties['assignedByGroup'] -or
                    -not $assignment.PSObject.Properties['disabledPlans'] -or
                    $assignment.disabledPlans -isnot [array]) {
                    $stateComplete = $false
                    continue
                }
                $skuKey = [string]$assignment.skuId
                if (-not $states.ContainsKey($skuKey)) {
                    $states[$skuKey] = [pscustomobject]@{ Direct = $false; Group = $false; DisabledPlans = @() }
                }
                if ($null -eq $assignment.assignedByGroup) {
                    $states[$skuKey].Direct = $true
                    $states[$skuKey].DisabledPlans = @($assignment.disabledPlans)
                }
                else { $states[$skuKey].Group = $true }
            }
            foreach ($license in $Items) {
                if (-not $states.ContainsKey([string]$license.skuId)) { $stateComplete = $false }
            }
            if (-not $stateComplete) {
                Write-NCMessage "Unable to read complete license assignment state for $($User.userPrincipalName)." -Level ERROR
                return $null
            }

            $resolved = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $Selections) {
                $match = $null
                foreach ($license in $Items) {
                    $display = $null
                    $source = $null
                    if ($licenseLookup) { $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $license.skuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$source) }
                    if ($entry -ieq [string]$license.skuId -or $entry -ieq [string]$license.skuPartNumber -or ($display -and $entry -ieq $display)) {
                        $match = $license
                        break
                    }
                }
                if (-not $match) {
                    Write-NCMessage "License '$entry' is not currently assigned to $($User.userPrincipalName)." -Level ERROR
                    return $null
                }
                $state = $states[[string]$match.skuId]
                if ($state.Group) {
                    Write-NCMessage "License '$entry' is assigned to $($User.userPrincipalName) through a group and cannot be moved directly." -Level ERROR
                    return $null
                }
                if (-not $state.Direct) {
                    Write-NCMessage "License '$entry' is not a direct assignment on $($User.userPrincipalName)." -Level ERROR
                    return $null
                }
                $resolved.Add([pscustomobject]@{
                        SkuId = [string]$match.skuId
                        Name = if ($display) { [string]$display } else { [string]$match.skuPartNumber }
                        DisabledPlans = @($state.DisabledPlans)
                    })
            }
            return ,@($resolved)
        }

        $firstLicenses = & $resolveSelections $firstUser @($detailResults[0].Items) $firstSelections
        if ($null -eq $firstLicenses) { return }
        $secondLicenses = & $resolveSelections $secondUser @($detailResults[1].Items) $secondSelections
        if ($null -eq $secondLicenses) { return }

        $firstSkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $secondSkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($license in $firstLicenses) { $null = $firstSkuIds.Add([string]$license.SkuId) }
        foreach ($license in $secondLicenses) { $null = $secondSkuIds.Add([string]$license.SkuId) }
        $firstAlreadyOnSecond = @($firstLicenses | Where-Object { $secondSkuIds.Contains([string]$_.SkuId) })
        $secondAlreadyOnFirst = @($secondLicenses | Where-Object { $firstSkuIds.Contains([string]$_.SkuId) })
        if ($firstAlreadyOnSecond.Count -gt 0 -or $secondAlreadyOnFirst.Count -gt 0) {
            if (-not $AllLicenses) {
                $overlapNames = @($firstAlreadyOnSecond + $secondAlreadyOnFirst | ForEach-Object { $_.Name } | Select-Object -Unique)
                Write-NCMessage ("The selected licenses overlap between users ({0}); no swap was performed." -f ($overlapNames -join ', ')) -Level ERROR
                return
            }
            $firstLicenses = @($firstLicenses | Where-Object { -not $secondSkuIds.Contains([string]$_.SkuId) })
            $secondLicenses = @($secondLicenses | Where-Object { -not $firstSkuIds.Contains([string]$_.SkuId) })
            if ($firstLicenses.Count -eq 0 -and $secondLicenses.Count -eq 0) {
                Write-NCMessage 'Both users already have the same direct license set; no swap is needed.' -Level SUCCESS
                return
            }
            if ($firstLicenses.Count -eq 0 -or $secondLicenses.Count -eq 0) {
                Write-NCMessage 'After retaining licenses shared by both users, each user must have at least one license to swap.' -Level ERROR
                return
            }
        }
        $firstCurrentlyAssignedAtSecond = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $secondCurrentlyAssignedAtFirst = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($license in $detailResults[1].Items) { $null = $firstCurrentlyAssignedAtSecond.Add([string]$license.skuId) }
        foreach ($license in $detailResults[0].Items) { $null = $secondCurrentlyAssignedAtFirst.Add([string]$license.skuId) }
        $firstSkuAlreadyHeldBySecond = @($firstLicenses | Where-Object { $firstCurrentlyAssignedAtSecond.Contains([string]$_.SkuId) })
        $secondSkuAlreadyHeldByFirst = @($secondLicenses | Where-Object { $secondCurrentlyAssignedAtFirst.Contains([string]$_.SkuId) })
        if ($firstSkuAlreadyHeldBySecond.Count -gt 0 -or $secondSkuAlreadyHeldByFirst.Count -gt 0) {
            $alreadyHeldNames = @($firstSkuAlreadyHeldBySecond + $secondSkuAlreadyHeldByFirst | ForEach-Object { $_.Name } | Select-Object -Unique)
            Write-NCMessage ("The destination already has selected license(s) ({0}); no swap was performed." -f ($alreadyHeldNames -join ', ')) -Level ERROR
            return
        }

        $summary = "Remove {0} from {1} and {2} from {3}, then assign them reciprocally" -f ($firstLicenses.Name -join ', '), $firstUser.userPrincipalName, ($secondLicenses.Name -join ', '), $secondUser.userPrincipalName
        if (-not $PSCmdlet.ShouldProcess("$($firstUser.userPrincipalName) <-> $($secondUser.userPrincipalName)", $summary)) {
            Write-NCMessage "License swap skipped for $($firstUser.userPrincipalName) and $($secondUser.userPrincipalName)." -Level INFO
            return
        }

        $defaultUsageLocation = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) { [string]$NCVars.UsageLocation } else { 'US' }
        $usagePatches = [System.Collections.Generic.List[object]]::new()
        $patchTargets = @{}
        if ($firstLicenses.Count -gt 0 -and [string]::IsNullOrWhiteSpace([string]$secondUser.usageLocation)) {
            $usagePatches.Add(@{ Id = 'u0'; Method = 'PATCH'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))"; Body = @{ usageLocation = $defaultUsageLocation } })
            $patchTargets['u0'] = $secondUser
        }
        if ($secondLicenses.Count -gt 0 -and [string]::IsNullOrWhiteSpace([string]$firstUser.usageLocation)) {
            $usagePatches.Add(@{ Id = 'u1'; Method = 'PATCH'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))"; Body = @{ usageLocation = $defaultUsageLocation } })
            $patchTargets['u1'] = $firstUser
        }
        if ($usagePatches.Count -gt 0) {
            $patchResults = @(Invoke-NCGraphBatch -Requests @($usagePatches) -Activity 'Preparing license swap users')
            $patchFailed = @($patchResults | Where-Object { -not $_.Success })
            if ($patchFailed.Count -gt 0) {
                foreach ($failure in $patchFailed) {
                    Write-NCMessage "Unable to set usage location for $($patchTargets[$failure.Id].userPrincipalName): $($failure.ErrorMessage). No licenses were moved." -Level ERROR
                }
                return
            }
        }

        $removeRequests = @(
            @{ Id = 'r0'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))/assignLicense"; Body = @{ addLicenses = @(); removeLicenses = @($firstLicenses.SkuId) } }
            @{ Id = 'r1'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))/assignLicense"; Body = @{ addLicenses = @(); removeLicenses = @($secondLicenses.SkuId) } }
        )
        $removeResults = @(Invoke-NCGraphBatch -Requests $removeRequests -Activity 'Removing licenses before swap')
        $removeById = @{}
        foreach ($result in $removeResults) { $removeById[$result.Id] = $result }
        $removeFailed = @($removeRequests | Where-Object { -not $removeById[$_.Id].Success })
        if ($removeFailed.Count -gt 0) {
            foreach ($request in $removeFailed) { Write-NCMessage "License removal failed for $($request.Url): $($removeById[$request.Id].ErrorMessage)" -Level ERROR }
            $restoreRequests = [System.Collections.Generic.List[object]]::new()
            if ($removeById['r0'].Success) { $restoreRequests.Add(@{ Id = 'restore0'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))/assignLicense"; Body = @{ addLicenses = @($firstLicenses | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans) } }); removeLicenses = @() } }) }
            if ($removeById['r1'].Success) { $restoreRequests.Add(@{ Id = 'restore1'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))/assignLicense"; Body = @{ addLicenses = @($secondLicenses | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans) } }); removeLicenses = @() } }) }
            if ($restoreRequests.Count -gt 0) {
                foreach ($result in @(Invoke-NCGraphBatch -Requests @($restoreRequests) -Activity 'Restoring licenses after failed removal')) {
                    $user = if ($result.Id -eq 'restore0') { $firstUser } else { $secondUser }
                    if ($result.Success) { Write-NCMessage "Restored original licenses for $($user.userPrincipalName)." -Level WARNING }
                    else { Write-NCMessage "Failed to restore original licenses for $($user.userPrincipalName): $($result.ErrorMessage)" -Level ERROR }
                }
            }
            Write-NCMessage "License swap aborted for $($firstUser.userPrincipalName) and $($secondUser.userPrincipalName); reciprocal assignments were not attempted." -Level ERROR
            return
        }

        $addRequests = @(
            @{ Id = 'a0'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))/assignLicense"; Body = @{ addLicenses = @($secondLicenses | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans) } }); removeLicenses = @() } }
            @{ Id = 'a1'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))/assignLicense"; Body = @{ addLicenses = @($firstLicenses | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans) } }); removeLicenses = @() } }
        )
        $addResults = @(Invoke-NCGraphBatch -Requests $addRequests -Activity 'Assigning swapped licenses')
        $addById = @{}
        foreach ($result in $addResults) { $addById[$result.Id] = $result }
        $addFailed = @($addRequests | Where-Object { -not $addById[$_.Id].Success })
        if ($addFailed.Count -eq 0) {
            Write-NCMessage ("License swap completed: {0} received {1}; {2} received {3}." -f $firstUser.userPrincipalName, ($secondLicenses.Name -join ', '), $secondUser.userPrincipalName, ($firstLicenses.Name -join ', ')) -Level SUCCESS
            return
        }

        foreach ($request in $addFailed) { Write-NCMessage "License assignment failed for $($request.Url): $($addById[$request.Id].ErrorMessage)" -Level ERROR }
        $rollbackRemovals = [System.Collections.Generic.List[object]]::new()
        if ($addById['a0'].Success) { $rollbackRemovals.Add(@{ Id = 'undo0'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))/assignLicense"; Body = @{ addLicenses = @(); removeLicenses = @($secondLicenses.SkuId) } }) }
        if ($addById['a1'].Success) { $rollbackRemovals.Add(@{ Id = 'undo1'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))/assignLicense"; Body = @{ addLicenses = @(); removeLicenses = @($firstLicenses.SkuId) } }) }
        if ($rollbackRemovals.Count -gt 0) {
            foreach ($result in @(Invoke-NCGraphBatch -Requests @($rollbackRemovals) -Activity 'Removing partial swap assignments')) {
                $user = if ($result.Id -eq 'undo0') { $firstUser } else { $secondUser }
                if (-not $result.Success) { Write-NCMessage "Failed to remove partial swap assignment from $($user.userPrincipalName): $($result.ErrorMessage)" -Level ERROR }
            }
        }
        $restoreOriginals = @(
            @{ Id = 'return0'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$firstUser.id))/assignLicense"; Body = @{ addLicenses = @($firstLicenses | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans) } }); removeLicenses = @() } }
            @{ Id = 'return1'; Method = 'POST'; Url = "/users/$([uri]::EscapeDataString([string]$secondUser.id))/assignLicense"; Body = @{ addLicenses = @($secondLicenses | ForEach-Object { @{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans) } }); removeLicenses = @() } }
        )
        foreach ($result in @(Invoke-NCGraphBatch -Requests $restoreOriginals -Activity 'Restoring original swap licenses')) {
            $user = if ($result.Id -eq 'return0') { $firstUser } else { $secondUser }
            if ($result.Success) { Write-NCMessage "Restored original licenses for $($user.userPrincipalName)." -Level WARNING }
            else { Write-NCMessage "Failed to restore original licenses for $($user.userPrincipalName): $($result.ErrorMessage)" -Level ERROR }
        }
        Write-NCMessage "License swap failed for $($firstUser.userPrincipalName) and $($secondUser.userPrincipalName); restoration was attempted." -Level ERROR
    }
    finally { Restore-ProgressAndInfoPreferences }
}

function Set-UserUsageLocation {
    <#
    .SYNOPSIS
        Updates usage location for one or more users.
    .DESCRIPTION
        Resolves users from pipeline or explicit input, connects to Microsoft Graph, and updates
        their UsageLocation value. If -UsageLocation is omitted, Nebula.Core uses the configured
        UsageLocation default (or US when no override exists).
    .PARAMETER UserPrincipalName
        User principal name, object ID, or short identifier. Accepts pipeline input.
    .PARAMETER UsageLocation
        Two-letter country code to set. When omitted, the configured default is used.
    .PARAMETER PassThru
        Emit the processed users as objects.
    .EXAMPLE
        Set-UserUsageLocation -UserPrincipalName user@contoso.com -UsageLocation IT
    .EXAMPLE
        'user1@contoso.com','user2@contoso.com' | Set-UserUsageLocation -UsageLocation DE
    .EXAMPLE
        Get-MgUser -Filter "endsWith(userPrincipalName,'@contoso.com')" | Set-UserUsageLocation -UsageLocation FR
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [string]$UsageLocation,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($entry in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                $targets.Add($entry.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if ($targets.Count -eq 0) {
                Write-NCMessage "No user principal names provided." -Level WARNING
                return
            }

            $GraphConnection = Test-MgGraphConnection -Scopes @('Directory.ReadWrite.All') -EnsureExchangeOnline:$false
            if (-not $GraphConnection) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            $defaultUsageLocation = if (-not [string]::IsNullOrWhiteSpace($UsageLocation)) {
                $UsageLocation.Trim()
            }
            elseif (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) {
                [string]$NCVars.UsageLocation
            }
            else {
                'US'
            }

            $targetUsage = $defaultUsageLocation.Trim().ToUpperInvariant()
            if ($targetUsage.Length -ne 2) {
                Write-NCMessage "UsageLocation must be a two-letter country code." -Level ERROR
                return
            }

            $normalizeUsageLocation = {
                param($value)
                if ([string]::IsNullOrWhiteSpace($value)) { return $null }
                return $value.Trim().ToUpperInvariant()
            }

            $queue = [System.Collections.Generic.List[string]]::new()
            $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($entry in $targets) {
                if ($dedup.Add($entry)) {
                    $queue.Add($entry) | Out-Null
                }
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $updatedCount = 0
            $skippedCount = 0
            $counter = 0

            Write-NCGraphBatchNotice -Count ($queue.Count) -Noun 'user(s)'

            for ($offset = 0; $offset -lt $queue.Count; $offset += 20) {
                $chunk = @($queue[$offset..([Math]::Min($offset + 20, $queue.Count) - 1)])
                $Percentage = Get-NCProgressPercent -Current $counter -Total $queue.Count
                Write-Progress -Activity "Updating usage location" -Status "$counter of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                # (a) Resolve the whole chunk.
                $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $resolvedUsers = Resolve-NCGraphUserBatch -Identifier $chunk -Property @('id', 'userPrincipalName', 'displayName', 'usageLocation') -FailedIdentifier $failedUsers

                # (b) Decide per user: skip or confirm.
                $approved = [System.Collections.Generic.List[object]]::new()
                $slots = [System.Collections.Generic.List[object]]::new()
                foreach ($upn in $chunk) {
                    $counter++
                    $user = $resolvedUsers[$upn.Trim()]
                    if (-not $user -or -not $user.id) {
                        if (-not $failedUsers.Contains($upn.Trim())) {
                            Write-NCMessage "Unable to resolve user recipient for $upn" -Level ERROR
                        }
                        continue
                    }

                    $currentUsage = & $normalizeUsageLocation $user.usageLocation
                    if ($currentUsage -eq $targetUsage) {
                        $skippedCount++
                        if ($PassThru.IsPresent) {
                            $slots.Add([pscustomobject]@{
                                    UserPrincipalName     = $user.userPrincipalName
                                    DisplayName           = $user.displayName
                                    PreviousUsageLocation = $user.usageLocation
                                    UsageLocation         = $user.usageLocation
                                    Action                = 'Skipped'
                                }) | Out-Null
                        }
                        Write-Verbose "Usage location already set to $targetUsage for $($user.userPrincipalName)."
                        continue
                    }

                    if (-not $PSCmdlet.ShouldProcess($user.userPrincipalName, "Set usage location to $targetUsage")) {
                        continue
                    }

                    $approved.Add([pscustomobject]@{ User = $user; Slot = $slots.Count }) | Out-Null
                    $slots.Add($null) | Out-Null
                }

                # (c) Update the approved users in one batch.
                if ($approved.Count -gt 0) {
                    $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                            @{
                                Id     = "m$i"
                                Method = 'PATCH'
                                Url    = "/users/$([uri]::EscapeDataString([string]$approved[$i].User.id))"
                                Body   = @{ usageLocation = $targetUsage }
                            }
                        })
                    $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity 'Updating usage location')
                    for ($i = 0; $i -lt $approved.Count; $i++) {
                        $user = $approved[$i].User
                        $response = $responses[$i]
                        if ($response.Success) {
                            $updatedCount++
                            if ($PassThru.IsPresent) {
                                $slots[$approved[$i].Slot] = [pscustomobject]@{
                                    UserPrincipalName     = $user.userPrincipalName
                                    DisplayName           = $user.displayName
                                    PreviousUsageLocation = $user.usageLocation
                                    UsageLocation         = $targetUsage
                                    Action                = 'Updated'
                                }
                            }
                            Write-Verbose "Usage location set to $targetUsage for $($user.userPrincipalName)."
                        }
                        else {
                            Write-NCMessage "Unable to set usage location ($targetUsage) for $($user.userPrincipalName): $($response.ErrorMessage)" -Level ERROR
                        }
                    }
                }

                foreach ($slot in $slots) {
                    if ($null -ne $slot) { $results.Add($slot) | Out-Null }
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($updatedCount -gt 0) {
                $summary = "Usage location updated for {0} user(s) to {1}." -f $updatedCount, $targetUsage
                if ($skippedCount -gt 0) {
                    $summary = $summary + (" {0} user(s) already had that value." -f $skippedCount)
                }
                Write-NCMessage $summary -Level SUCCESS
            }
            elseif ($skippedCount -gt 0) {
                Write-NCMessage ("Usage location already set to {0} for {1} user(s)." -f $targetUsage, $skippedCount) -Level INFO
            }
            else {
                Write-NCMessage "No usage location changes were applied." -Level WARNING
            }
        }
        finally {
            Write-Progress -Activity "Processing users" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Update-LicenseCatalog {
    <#
    .SYNOPSIS
        Forces an immediate refresh of the cached license catalog.
    .DESCRIPTION
        Downloads the latest catalog from GitHub, updates the local cache, and returns the resulting
        object so callers can inspect the data if needed.
    .EXAMPLE
        Update-LicenseCatalog
    #>
    [CmdletBinding()]
    param()

    try {
        $catalog = Get-LicenseCatalog -ForceRefresh -IncludeMetadata
        if ($catalog.LastCommitUtc) {
            $timestamp = Format-NCDateTime -Value $catalog.LastCommitUtc -AsLocalTime
            Write-NCMessage "Primary license catalog refreshed. Last commit: $timestamp" -Level SUCCESS
        }
        else {
            Write-NCMessage "Primary license catalog refreshed." -Level SUCCESS
        }

        if ($catalog.CustomLookup) {
            if ($catalog.CustomLastCommitUtc) {
                $customStamp = Format-NCDateTime -Value $catalog.CustomLastCommitUtc -AsLocalTime
                Write-NCMessage "Custom license catalog refreshed. Last commit: $customStamp" -Level INFO
            }
            else {
                Write-NCMessage "Custom license catalog refreshed." -Level INFO
            }
        }

        return $catalog
    }
    catch {
        Write-NCMessage "Unable to refresh license catalog. $($_.Exception.Message)" -Level ERROR
        throw
    }
}


