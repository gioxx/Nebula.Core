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
        Include exception details in error messages.
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
    }

    process {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $UserPrincipalName -PreferGraphIdentity
        if (-not $resolvedPrincipal) {
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        }

        try {
            $user = Get-MgUser -UserId $resolvedPrincipal -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
        }
        catch {
            $detail = if ($ShowErrorDetails.IsPresent) { ": $($_.Exception.Message)" } else { "." }
            Write-NCMessage ("User {0} not found or query failed{1}" -f $UserPrincipalName, $detail) -Level ERROR
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

        $normalizedCurrentUsage = & $normalizeUsageLocation $user.UsageLocation
        $normalizedTargetUsage = & $normalizeUsageLocation $defaultUsageLocation
        $targetUsage = if ($normalizedTargetUsage -and $normalizedTargetUsage -ne $normalizedCurrentUsage) { $defaultUsageLocation } else { $null }

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

        $maxAttempts = 3
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

        if (-not $tenantSkus -or $tenantSkus.Count -eq 0) {
            Write-NCMessage "No tenant licenses available to assign." -Level WARNING
            return
        }

        $normalizeString = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return ($value.Trim().ToUpperInvariant())
        }

        $resolved = @()
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
            return
        }

        $uniqueAdds = $resolved | Group-Object SkuId | ForEach-Object {
            $_.Group | Select-Object -First 1
        }

        $addLicenses = @()
        $assignableItems = @()
        $namesNoAvailability = @()
        foreach ($item in $uniqueAdds) {
            $available = $item.Available
            if ($available -le 0) {
                Write-NCMessage ("No available units for license {0} ({1}) (available: {2})" -f $item.Name, $item.SkuPartNumber, $available) -Level WARNING
                $namesNoAvailability += $item.Name
                continue
            }
            $assignableItems += $item
            $addLicenses += @{
                SkuId         = $item.SkuId
                DisabledPlans = @()
            }
        }

        if ($addLicenses.Count -eq 0) {
            $requestedList = ($resolved | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '
            Write-NCMessage ("No licenses to assign: none available. Requested: {0}" -f $requestedList) -Level ERROR
            return
        }

        $assignableList = ($assignableItems | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '
        $summary = if ($targetUsage) {
            "Set usage location to {0} and assign license(s): {1} to {2}" -f $targetUsage, $assignableList, $user.UserPrincipalName
        }
        else {
            "Assign license(s): {0} to {1}" -f $assignableList, $user.UserPrincipalName
        }

        if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, $summary)) {
            return
        }

        if ($targetUsage) {
            try {
                Update-MgUser -UserId $user.Id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                $user.UsageLocation = $targetUsage
                Write-Verbose "Usage location set to $targetUsage for $($user.UserPrincipalName)."
            }
            catch {
                Write-NCMessage "Unable to set usage location ($targetUsage) for $($user.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        try {
            Invoke-NCRetry -Action {
                Set-MgUserLicense -UserId $user.Id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "assign licenses to $($user.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage ("Failed to assign licenses to {0}, attempt {1} of {2}. {3}" -f $user.UserPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
            } | Out-Null
            Write-NCMessage ("Assigned license(s) to {0}: {1}" -f $user.UserPrincipalName, $assignableList) -Level SUCCESS
            if ($namesNoAvailability.Count -gt 0) {
                Write-NCMessage ("Skipped license(s) with no available units: {0}" -f (($namesNoAvailability | Select-Object -Unique) -join ', ')) -Level WARNING
            }
        }
        catch {
            Write-NCMessage "License assignment failed for $($user.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
    }
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

            foreach ($upn in $queue) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $queue.Count
                Write-Progress -Activity "Processing $upn" -Status "$counter of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $upn -PreferGraphIdentity
                if (-not $resolvedPrincipal) {
                    Write-NCMessage "Unable to resolve user recipient for $upn" -Level ERROR
                    continue
                }

                try {
                    $user = Get-MgUser -UserId $resolvedPrincipal -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to retrieve user '$upn'. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                $currentUsage = & $normalizeUsageLocation $user.UsageLocation
                if ($currentUsage -eq $targetUsage) {
                    $skippedCount++
                    if ($PassThru.IsPresent) {
                        $results.Add([pscustomobject]@{
                                UserPrincipalName     = $user.UserPrincipalName
                                DisplayName           = $user.DisplayName
                                PreviousUsageLocation = $user.UsageLocation
                                UsageLocation         = $user.UsageLocation
                                Action                = 'Skipped'
                            }) | Out-Null
                    }
                    Write-Verbose "Usage location already set to $targetUsage for $($user.UserPrincipalName)."
                    continue
                }

                if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, "Set usage location to $targetUsage")) {
                    continue
                }

                try {
                    Update-MgUser -UserId $user.Id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                    $updatedCount++
                    if ($PassThru.IsPresent) {
                        $results.Add([pscustomobject]@{
                                UserPrincipalName     = $user.UserPrincipalName
                                DisplayName           = $user.DisplayName
                                PreviousUsageLocation = $user.UsageLocation
                                UsageLocation         = $targetUsage
                                Action                = 'Updated'
                            }) | Out-Null
                    }
                    Write-Verbose "Usage location set to $targetUsage for $($user.UserPrincipalName)."
                }
                catch {
                    Write-NCMessage "Unable to set usage location ($targetUsage) for $($user.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
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

            $counter = 0
            foreach ($upn in $queue) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $queue.Count
                Write-Progress -Activity "Processing $upn" -Status "$counter of $($queue.Count) - $Percentage%" -PercentComplete $Percentage

                $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $upn -PreferGraphIdentity
                if (-not $resolvedPrincipal) {
                    Write-NCMessage "Unable to resolve user recipient for $upn" -Level ERROR
                    continue
                }

                try {
                    $user = Get-MgUser -UserId $resolvedPrincipal -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to retrieve user '$upn'. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                [pscustomobject]@{
                    UserPrincipalName                = $user.UserPrincipalName
                    DisplayName                      = $user.DisplayName
                    UsageLocation                    = $user.UsageLocation
                    ConfiguredDefaultUsageLocation   = $defaultUsageLocation
                    MatchesConfiguredDefault         = (& $normalizeUsageLocation $user.UsageLocation) -eq (& $normalizeUsageLocation $defaultUsageLocation)
                }
            }
        }
        finally {
            Write-Progress -Activity "Processing users" -Completed
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

        $resolvedSource = Find-UserRecipient -UserPrincipalName $SourceUserPrincipalName -PreferGraphIdentity
        $resolvedDestination = Find-UserRecipient -UserPrincipalName $DestinationUserPrincipalName -PreferGraphIdentity

        try {
            if ($resolvedSource) {
                $sourceUser = Get-MgUser -UserId $resolvedSource -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
            }
            else {
                $sourceUser = Get-MgUser -UserId $SourceUserPrincipalName -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
                $resolvedSource = $sourceUser.Id
            }
        }
        catch {
            Write-NCMessage "Unable to retrieve source user $($SourceUserPrincipalName): $($_.Exception.Message)" -Level ERROR
            return
        }

        try {
            if ($resolvedDestination) {
                $destinationUser = Get-MgUser -UserId $resolvedDestination -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
            }
            else {
                $destinationUser = Get-MgUser -UserId $DestinationUserPrincipalName -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
                $resolvedDestination = $destinationUser.Id
            }
        }
        catch {
            Write-NCMessage "Unable to retrieve destination user $($DestinationUserPrincipalName): $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($resolvedSource -eq $resolvedDestination) {
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

        $currentUsage = & $normalizeUsageLocation $destinationUser.UsageLocation
        $desiredUsage = & $normalizeUsageLocation $defaultUsageLocation

        if ($desiredUsage -and $desiredUsage -ne $currentUsage) {
            $targetUsage = $defaultUsageLocation
            try {
                Update-MgUser -UserId $destinationUser.Id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                $destinationUser.UsageLocation = $targetUsage
                Write-Verbose "Usage location set to $targetUsage for $($destinationUser.UserPrincipalName)."
            }
            catch {
                Write-NCMessage "Unable to set usage location ($targetUsage) for $($destinationUser.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
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

        try {
            $sourceLicenses = Invoke-NCRetry -Action {
                Get-MgUserLicenseDetail -UserId $sourceUser.Id -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve licenses for $($sourceUser.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve licenses for $($sourceUser.UserPrincipalName), attempt $currentAttempt of $currentMax." -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Failed to retrieve licenses for $($sourceUser.UserPrincipalName) after $maxAttempts attempts." -Level ERROR
            return
        }

        if (-not $sourceLicenses -or $sourceLicenses.Count -eq 0) {
            Write-NCMessage "Source user $($sourceUser.UserPrincipalName) has no licenses to copy." -Level WARNING
            return
        }

        try {
            $destinationLicenses = Get-MgUserLicenseDetail -UserId $destinationUser.Id -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to read destination licenses for $($destinationUser.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
            return
        }

        $destinationSkuIds = if ($destinationLicenses) { $destinationLicenses.SkuId } else { @() }
        $addLicenses = @()
        $licenseNames = @()
        $skippedInvalid = @()

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
            Write-NCMessage "Nothing to copy: destination already has all licenses from $($sourceUser.UserPrincipalName)." -Level WARNING
            return
        }

        $uniqueNames = $licenseNames | Select-Object -Unique
        $actionSummary = "Copy licenses ($($uniqueNames -join ', ')) from $($sourceUser.UserPrincipalName) to $($destinationUser.UserPrincipalName)"

        if ($skippedInvalid.Count -gt 0) {
            Write-NCMessage ("Skipped licenses with invalid IDs: {0}" -f ($skippedInvalid -join '; ')) -Level WARNING
        }

        if (-not $PSCmdlet.ShouldProcess($destinationUser.UserPrincipalName, $actionSummary)) {
            return
        }

        try {
            Invoke-NCRetry -Action {
                Set-MgUserLicense -UserId $destinationUser.Id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "assign licenses to $($destinationUser.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage ("Failed to assign licenses to {0}, attempt {1} of {2}. {3}" -f $destinationUser.UserPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
            } | Out-Null
            Write-NCMessage "Copied licenses to $($destinationUser.UserPrincipalName): $($uniqueNames -join ', ')." -Level SUCCESS
        }
        catch {
            Write-NCMessage "License copy to $($destinationUser.UserPrincipalName) failed: $($_.Exception.Message)" -Level ERROR
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
    .PARAMETER ShowErrorDetails
        Include exception details in error messages.
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

        foreach ($User in $Users) {
            $processedCount++
            $Percentage = Get-NCProgressPercent -Current $processedCount -Total $totalUsers
            Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$processedCount out of $totalUsers - $Percentage%" -PercentComplete $Percentage

            if ($processedUsers.Contains($User.UserPrincipalName)) {
                Write-NCMessage "Skipping $($User.UserPrincipalName), already processed." -Level WARNING
                continue
            }

            $processedSinceFlush++

            try {
                $GraphLicense = Invoke-NCRetry -Action {
                    Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve licenses for $($User.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName), attempt $currentAttempt of $currentMax" -Level ERROR
            }
        }
            catch {
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
                $userRows = $userRows | ForEach-Object {
                    $_ | Add-Member -NotePropertyName MatchedLicenses -NotePropertyValue ($matchedNames -join ', ') -PassThru
                }
            }
            elseif ($License) {
                $userRows = $userRows | ForEach-Object {
                    $_ | Add-Member -NotePropertyName MatchedLicenses -NotePropertyValue $null -PassThru
                }
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
                                -Property Id,UserPrincipalName,DisplayName,Mail,ProxyAddresses `
                                -ErrorAction Stop |
                                Where-Object { & $matchesDomain $_ } |
                                Select-Object -First $sampleUserLimit
                        }
                        else {
                            Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($sku.SkuId))" `
                                -Top $sampleUserLimit `
                                -ConsistencyLevel eventual `
                                -Property Id,UserPrincipalName,DisplayName `
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
                    Name          = $sku.Name
                    SkuPartNumber = $sku.SkuPartNumber
                    SkuId         = $sku.SkuId
                    Total         = $sku.Total
                    TotalCount    = $sku.TotalCount
                    Consumed      = $sku.Consumed
                    Available     = $sku.Available
                    Enabled       = $sku.Enabled
                    Suspended     = $sku.Suspended
                    Warning       = $sku.Warning
                    Source        = $sku.Source
                    SampleUsers   = $sampleUserList
                    SampleUsersText = if ($sampleUserList.Count -gt 0) { $sampleUserList -join [Environment]::NewLine } else { $null }
                }
            }
        }

        if ($GridView.IsPresent) {
            $summaryRows = $sorted | Select-Object Name, SkuPartNumber, Total, Consumed, Available
            $summaryRows | Out-GridView -Title "M365 Tenant Licenses"

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
                    $sampleRows | Out-GridView -Title "M365 Tenant License Sample Users"
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
        $clipboardLines = @()
        $clipboardHasContent = $false
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
                $licenseCatalog.LastCommitUtc.ToLocalTime().ToString($NCVars.DateTimeString_Full)
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
    }

    process {
        if (-not $initSucceeded) { return }
        
        $inputUserPrincipalName = $UserPrincipalName
        $resolvedGraphIdentity = Find-UserRecipient -UserPrincipalName $inputUserPrincipalName -PreferGraphIdentity
        if (-not $resolvedGraphIdentity) {
            Write-NCMessage "Unable to resolve user recipient for $inputUserPrincipalName" -Level ERROR
            return
        }

        try {
            $User = Get-MgUser -UserId $resolvedGraphIdentity -Property Id, DisplayName, UserPrincipalName, Mail -ErrorAction Stop
        }
        catch {
            $detail = if ($ShowErrorDetails.IsPresent) { ": $($_.Exception.Message)" } else { "." }
            Write-NCMessage ("User {0} not found or query failed{1}" -f $inputUserPrincipalName, $detail) -Level ERROR
            return
        }

        Write-NCMessage ("`nProcessing user: {0} <{1}>{2}`n" -f $User.DisplayName, $User.UserPrincipalName, $catalogInfo) -Level SUCCESS

        try {
            $GraphLicense = Invoke-NCRetry -Action {
                Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve licenses for $($User.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName), attempt $currentAttempt of $currentMax" -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName) after $maxAttempts attempts." -Level ERROR
            return
        }

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
                    $clipboardLines += ($quoted -join ",")
                    $clipboardHasContent = $true
                }
            }
        }
        else {
            Write-Verbose "No licenses assigned to this user."
            Write-NCMessage ("No licenses assigned to user {0}." -f $User.UserPrincipalName) -Level WARNING
        }
    }
    end {
        if ($Clipboard.IsPresent -and $clipboardHasContent) {
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

        $resolvedSource = Find-UserRecipient -UserPrincipalName $SourceUserPrincipalName -PreferGraphIdentity
        $resolvedDestination = Find-UserRecipient -UserPrincipalName $DestinationUserPrincipalName -PreferGraphIdentity

        if (-not $resolvedSource) {
            Write-NCMessage "Unable to resolve source user recipient for $SourceUserPrincipalName" -Level ERROR
            return
        }
        if (-not $resolvedDestination) {
            Write-NCMessage "Unable to resolve destination user recipient for $DestinationUserPrincipalName" -Level ERROR
            return
        }
        if ($resolvedSource -eq $resolvedDestination) {
            Write-NCMessage "Source and destination users are the same. Aborting." -Level ERROR
            return
        }

        try {
            $sourceUser = Get-MgUser -UserId $resolvedSource -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
            $destinationUser = Get-MgUser -UserId $resolvedDestination -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to retrieve users: $($_.Exception.Message)" -Level ERROR
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

        $currentUsage = & $normalizeUsageLocation $destinationUser.UsageLocation
        $desiredUsage = & $normalizeUsageLocation $defaultUsageLocation

        if ($desiredUsage -and $desiredUsage -ne $currentUsage) {
            $targetUsage = $defaultUsageLocation
            try {
                Update-MgUser -UserId $destinationUser.Id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                $destinationUser.UsageLocation = $targetUsage
                Write-Verbose "Usage location set to $targetUsage for $($destinationUser.UserPrincipalName)."
            }
            catch {
                Write-NCMessage "Unable to set usage location ($targetUsage) for $($destinationUser.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
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

        try {
            $sourceLicenses = Invoke-NCRetry -Action {
                Get-MgUserLicenseDetail -UserId $sourceUser.Id -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve licenses for $($sourceUser.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve licenses for $($sourceUser.UserPrincipalName), attempt $currentAttempt of $currentMax." -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Failed to retrieve licenses for $($sourceUser.UserPrincipalName) after $maxAttempts attempts." -Level ERROR
            return
        }

        if (-not $sourceLicenses -or $sourceLicenses.Count -eq 0) {
            Write-NCMessage "Source user $($sourceUser.UserPrincipalName) has no licenses to move." -Level WARNING
            return
        }

        try {
            $destinationLicenses = Get-MgUserLicenseDetail -UserId $destinationUser.Id -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to read destination licenses for $($destinationUser.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
            return
        }

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

        $addLicenses = @()
        $removeSkuIds = @()
        $skippedInvalid = @()

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

            $removeSkuIds += $parsedGuid
            if ($destinationSkuIds -contains $parsedGuid) {
                Write-Verbose "Destination already has $($lic.SkuPartNumber); skipping add."
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
        }

        if ($addLicenses.Count -eq 0 -and $removeSkuIds.Count -eq 0) {
            Write-NCMessage "Nothing to move between $($sourceUser.UserPrincipalName) and $($destinationUser.UserPrincipalName)." -Level WARNING
            return
        }

        if ($skippedInvalid.Count -gt 0) {
            Write-NCMessage ("Skipped licenses with invalid IDs: {0}" -f ($skippedInvalid -join '; ')) -Level WARNING
        }

        $licenseNames = $sourceLicenses | ForEach-Object {
            $matchSource = $null
            $name = if ($licenseLookup) {
                Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $_.SkuPartNumber -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
            }
            if ($name) { $name } else { $_.SkuPartNumber }
        }
        $actionSummary = "Move licenses ($($licenseNames -join ', ')) from $($sourceUser.UserPrincipalName) to $($destinationUser.UserPrincipalName)"

        if (-not $PSCmdlet.ShouldProcess($destinationUser.UserPrincipalName, $actionSummary)) {
            return
        }

        if ($addLicenses.Count -gt 0) {
            try {
                Invoke-NCRetry -Action {
                    Set-MgUserLicense -UserId $destinationUser.Id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
                } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "assign licenses to $($destinationUser.UserPrincipalName)" -OnError {
                    param($attempt, $max, $err)
                    $currentAttempt = if ($attempt) { $attempt } else { '?' }
                    $currentMax = if ($max) { $max } else { $maxAttempts }
                    Write-NCMessage ("Failed to assign licenses to {0}, attempt {1} of {2}. {3}" -f $destinationUser.UserPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
                } | Out-Null
                Write-NCMessage "Assigned licenses to $($destinationUser.UserPrincipalName)." -Level SUCCESS
            }
            catch {
                Write-NCMessage "License assignment to $($destinationUser.UserPrincipalName) failed. Aborting removal from source. $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        if ($removeSkuIds.Count -gt 0) {
            try {
                Invoke-NCRetry -Action {
                    Set-MgUserLicense -UserId $sourceUser.Id -AddLicenses @() -RemoveLicenses ($removeSkuIds | Select-Object -Unique) -ErrorAction Stop
                } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "remove licenses from $($sourceUser.UserPrincipalName)" -OnError {
                    param($attempt, $max, $err)
                    $currentAttempt = if ($attempt) { $attempt } else { '?' }
                    $currentMax = if ($max) { $max } else { $maxAttempts }
                    Write-NCMessage ("Failed to remove licenses from {0}, attempt {1} of {2}. {3}" -f $sourceUser.UserPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
                } | Out-Null
                Write-NCMessage "Removed licenses from $($sourceUser.UserPrincipalName)." -Level SUCCESS
            }
            catch {
                Write-NCMessage "Failed to remove licenses from $($sourceUser.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
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
    }

    process {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $UserPrincipalName -PreferGraphIdentity
        if (-not $resolvedPrincipal) {
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        }

        try {
            $user = Get-MgUser -UserId $resolvedPrincipal -ErrorAction Stop
        }
        catch {
            $detail = if ($ShowErrorDetails.IsPresent) { ": $($_.Exception.Message)" } else { "." }
            Write-NCMessage ("User {0} not found or query failed{1}" -f $UserPrincipalName, $detail) -Level ERROR
            return
        }

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

        $maxAttempts = 3
        try {
            $assignedLicenses = Invoke-NCRetry -Action {
                Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve licenses for $($user.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage "Failed to retrieve licenses for $($user.UserPrincipalName), attempt $currentAttempt of $currentMax." -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Unable to retrieve licenses for $($user.UserPrincipalName) after $maxAttempts attempts." -Level ERROR
            return
        }

        if (-not $assignedLicenses -or $assignedLicenses.Count -eq 0) {
            Write-NCMessage "User $($user.UserPrincipalName) has no licenses to remove." -Level WARNING
            return
        }

        $normalizeString = {
            param($value)
            if ([string]::IsNullOrWhiteSpace($value)) { return $null }
            return ($value.Trim().ToUpperInvariant())
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

            $removeLicenseIds = $removeLicenseIds | Where-Object { $_ } | Select-Object -Unique
            $licenseNames = $licenseNames | Where-Object { $_ } | Select-Object -Unique

            if ($removeLicenseIds.Count -eq 0) {
                Write-NCMessage "No licenses to remove for $($user.UserPrincipalName)." -Level WARNING
                return
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
                return
            }

            $removeLicenses = $resolved | Group-Object SkuId | ForEach-Object {
                $_.Group | Select-Object -First 1
            }

            if ($removeLicenses.Count -eq 0) {
                Write-NCMessage "No licenses matched for removal." -Level ERROR
                return
            }

            $licenseNames = $removeLicenses | ForEach-Object { $_.Name } | Select-Object -Unique
            $removeLicenseIds = $removeLicenses.SkuId
        }

        $summary = if ($PSCmdlet.ParameterSetName -eq 'All') {
            "Remove all license(s): {0} from {1}" -f ($licenseNames -join ', '), $user.UserPrincipalName
        }
        else {
            "Remove license(s): {0} from {1}" -f ($licenseNames -join ', '), $user.UserPrincipalName
        }

        if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, $summary)) {
            return
        }

        try {
            Invoke-NCRetry -Action {
                Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses $removeLicenseIds -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "remove licenses from $($user.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                $currentAttempt = if ($attempt) { $attempt } else { '?' }
                $currentMax = if ($max) { $max } else { $maxAttempts }
                Write-NCMessage ("Failed to remove licenses from {0}, attempt {1} of {2}. {3}" -f $user.UserPrincipalName, $currentAttempt, $currentMax, $err.Exception.Message) -Level ERROR
            } | Out-Null
            Write-NCMessage ("Removed license(s) from {0}: {1}" -f $user.UserPrincipalName, ($licenseNames -join ', ')) -Level SUCCESS
        }
        catch {
            Write-NCMessage "License removal failed for $($user.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
        }
    }

    end {
        Restore-ProgressAndInfoPreferences
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
            $timestamp = $catalog.LastCommitUtc.ToLocalTime().ToString($NCVars.DateTimeString_Full)
            Write-NCMessage "Primary license catalog refreshed. Last commit: $timestamp" -Level SUCCESS
        }
        else {
            Write-NCMessage "Primary license catalog refreshed." -Level SUCCESS
        }

        if ($catalog.CustomLookup) {
            if ($catalog.CustomLastCommitUtc) {
                $customStamp = $catalog.CustomLastCommitUtc.ToLocalTime().ToString($NCVars.DateTimeString_Full)
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


