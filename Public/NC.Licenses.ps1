#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Licenses =============================================================================================================================

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
    .EXAMPLE
        Add-MsolAccountSku -UserPrincipalName user@contoso.com -License "Microsoft 365 E3"
    .EXAMPLE
        Add-MsolAccountSku -UserPrincipalName user@contoso.com -License "ENTERPRISEPACK","VISIOCLIENT"
    .EXAMPLE
        Add-MsolAccountSku -UserPrincipalName user@contoso.com -License "18181a46-0d4e-45cd-891e-60aabd171b4e"
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [Alias('User', 'UPN')]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string[]]$License,
        [switch]$ForceLicenseCatalogRefresh
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $UserPrincipalName
        if (-not $resolvedPrincipal) {
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        }

        try {
            $user = Get-MgUser -UserId $resolvedPrincipal -Property Id,UserPrincipalName,DisplayName,UsageLocation -ErrorAction Stop
        }
        catch {
            Write-NCMessage "User $UserPrincipalName not found or query failed: $($_.Exception.Message)" -Level ERROR
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
                    $match = @{
                        SkuId         = $sku.SkuId
                        SkuPartNumber = $sku.SkuPartNumber
                        Name          = if ($display) { $display } else { $sku.SkuPartNumber }
                        Available     = ($sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning + $sku.PrepaidUnits.Suspended) - $sku.ConsumedUnits
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
        $namesNoAvailability = @()
        foreach ($item in $uniqueAdds) {
            $available = $item.Available
            if ($available -le 0) {
                Write-NCMessage ("No available units for license {0} ({1}) (available: {2})" -f $item.Name, $item.SkuPartNumber, $available) -Level WARNING
                $namesNoAvailability += $item.Name
                continue
            }
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

        $summary = if ($targetUsage) {
            "Set usage location to {0} and assign license(s): {1} to {2}" -f $targetUsage, (($resolved | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '), $user.UserPrincipalName
        }
        else {
            "Assign license(s): {0} to {1}" -f (($resolved | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '), $user.UserPrincipalName
        }

        if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, $summary)) {
            return
        }

        if ($targetUsage) {
            try {
                Update-MgUser -UserId $user.Id -UsageLocation $targetUsage -ErrorAction Stop | Out-Null
                $user.UsageLocation = $targetUsage
                Write-NCMessage "Usage location set to $targetUsage for $($user.UserPrincipalName)." -Level VERBOSE
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
            Write-NCMessage ("Assigned license(s) to {0}: {1}" -f $user.UserPrincipalName, (($resolved | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', ')) -Level SUCCESS
        }
        catch {
            Write-NCMessage "License assignment failed for $($user.UserPrincipalName): $($_.Exception.Message)" -Level ERROR
        }
    }
    finally {
        Restore-ProgressAndInfoPreferences
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
        [Parameter(Mandatory = $true)]
        [Alias('Source', 'From')]
        [string]$SourceUserPrincipalName,
        [Parameter(Mandatory = $true)]
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

        $resolvedSource = Find-UserRecipient -UserPrincipalName $SourceUserPrincipalName
        $resolvedDestination = Find-UserRecipient -UserPrincipalName $DestinationUserPrincipalName

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
                Write-NCMessage "Usage location set to $targetUsage for $($destinationUser.UserPrincipalName)." -Level VERBOSE
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
                Write-NCMessage "Destination already has $($lic.SkuPartNumber); skipping add." -Level VERBOSE
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
                        Write-NCMessage "Skipping invalid disabled plan '$planString' for $($lic.SkuPartNumber)." -Level VERBOSE
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
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .EXAMPLE
        Export-MsolAccountSku
    .EXAMPLE
        Export-MsolAccountSku -CsvFolder "C:\Temp"
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $True, HelpMessage = "Folder where export CSV file (e.g. C:\Temp)")]
        [string]$CSVFolder,
        [switch]$ForceLicenseCatalogRefresh
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
        $arr_MsolAccountSku = @()
        $ProcessedCount = 0
        $maxAttempts = 3
        $resolvedViaCustom = @{}
        $unknownSkuTracker = @{}

        $CSV = New-File("$($folder)\$((Get-Date -Format $($NCVars.DateTimeString_CSV)).ToString())_M365-User-License-Report.csv")
        if (Test-Path $CSV) {
            $ProcessedUsers = Import-CSV $CSV | Select-Object -ExpandProperty UserPrincipalName
        }
        else {
            $ProcessedUsers = @()
        }

        try {
            $Users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable totalUsers -All -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Failed to retrieve users with assigned licenses: $_" -Level ERROR
            return
        }

        foreach ($User in $Users) {
            $ProcessedCount++
            $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
            Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete

            if ($ProcessedUsers -contains $User.UserPrincipalName) {
                Write-NCMessage "Skipping $($User.UserPrincipalName), already processed." -Level WARNING
                continue
            }

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
                continue
            }

            if ($null -ne $GraphLicense) {
                foreach ($licenseSku in $GraphLicense.SkuPartNumber) {
                    $matchSource = $null
                    $productName = Get-LicenseDisplayName -Lookup $licenseLookup `
                        -SkuPartNumber $licenseSku `
                        -FallbackLookup $customLookup `
                        -MatchSource ([ref]$matchSource)

                    if (-not $productName) {
                        Write-Verbose "Unknown license: $licenseSku for $($User.UserPrincipalName)"
                        if ($unknownSkuTracker.ContainsKey($licenseSku)) {
                            $unknownSkuTracker[$licenseSku]++
                        }
                        else {
                            $unknownSkuTracker[$licenseSku] = 1
                        }
                        $productName = $licenseSku
                    }
                    elseif ($matchSource -and $matchSource -ne 'Primary') {
                        if ($resolvedViaCustom.ContainsKey($licenseSku)) {
                            $resolvedViaCustom[$licenseSku]++
                        }
                        else {
                            $resolvedViaCustom[$licenseSku] = 1
                        }
                    }

                    $arr_MsolAccountSku += [pscustomobject]@{
                        DisplayName        = $User.DisplayName
                        UserPrincipalName  = $User.UserPrincipalName
                        PrimarySmtpAddress = $User.Mail
                        Licenses           = $productName
                    }
                }
            }

            if ($ProcessedCount % 50 -eq 0) {
                Write-NCMessage "Processed $ProcessedCount out of $totalUsers, saving partial results ..." -Level VERBOSE
                $arr_MsolAccountSku | Export-CSV $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding) -Append
            }
        }

        $arr_MsolAccountSku | Export-CSV $CSV -NoTypeInformation -Delimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding)

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
        returns counts for total, consumed, available, suspended, and warning seats.
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .PARAMETER AsTable
        Display the result as a formatted table instead of returning objects.
    .PARAMETER GridView
        Show the result in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-TenantMsolAccountSku
    .EXAMPLE
        Get-TenantMsolAccountSku -AsTable
    #>
    [CmdletBinding()]
    param(
        [switch]$ForceLicenseCatalogRefresh,
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
            $total = $enabled + $suspended + $warning
            $consumed = if ($sku.ConsumedUnits -is [int]) { [int]$sku.ConsumedUnits } else { [int]0 }
            $available = if ($total -gt 0) { [Math]::Max($total - $consumed, 0) } else { $null }
            $nameSource = if ($matchSource) { $matchSource } elseif ($display) { 'Primary' } else { 'Unknown' }

            [pscustomobject][ordered]@{
                Name          = if ($display) { $display } else { $sku.SkuPartNumber }
                SkuPartNumber = $sku.SkuPartNumber
                SkuId         = $sku.SkuId
                Total         = $total
                Consumed      = $consumed
                Available     = $available
                Enabled       = $enabled
                Suspended     = $suspended
                Warning       = $warning
                Source        = $nameSource
            }
        }

        $sorted = $results | Sort-Object Name

        if ($GridView.IsPresent) {
            $sorted | Out-GridView -Title "M365 Tenant Licenses"
        }
        elseif ($AsTable.IsPresent) {
            $limited = $sorted | Select-Object @{
                    Name       = 'Name'
                    Expression = { Format-OutputString -Value $_.Name -MaxLength $NCVars.MaxFieldLength }
                }, SkuPartNumber, Total, Consumed, Available
            Show-Table -Rows $limited -AsTable
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
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .EXAMPLE
        Get-UserMsolAccountSku -UserPrincipalName "user@contoso.com"
    .EXAMPLE
        Get-UserMsolAccountSku -UserPrincipalName "user@contoso.com" -Clipboard
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "User Principal Name (e.g. user@contoso.com)")]
        [Alias('User', 'UPN')]
        [string] $UserPrincipalName,
        [switch] $Clipboard,
        [switch] $ForceLicenseCatalogRefresh
    )

    begin {
        Set-ProgressAndInfoPreferences
        $initSucceeded = $true
        $clipboardLines = @()

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
        }
    }

    process {
        if (-not $initSucceeded) { return }

        $inputUserPrincipalName = $UserPrincipalName
        $resolvedRecipient = Find-UserRecipient -UserPrincipalName $inputUserPrincipalName
        if (-not $resolvedRecipient) {
            Write-NCMessage "Unable to resolve user recipient for $inputUserPrincipalName" -Level ERROR
            return
        }

        try {
            $User = Get-MgUser -UserId $resolvedRecipient -ErrorAction Stop
        }
        catch {
            Write-NCMessage "User $inputUserPrincipalName not found or query failed: $_" -Level ERROR
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
            }

            if ($Clipboard.IsPresent) {
                $normalized = $licensesForClipboard | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                $quoted = $normalized | ForEach-Object { "`"$($_.Replace('"', '\"'))`"" }
                $clipboardLines += ($quoted -join ",")
            }
        }
        else {
            Write-NCMessage "No licenses assigned to this user." -Level VERBOSE
            if ($Clipboard.IsPresent) {
                $clipboardLines += ''
            }
        }
    }
    end {
        if ($Clipboard.IsPresent) {
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
        [Parameter(Mandatory = $true)]
        [Alias('Source', 'From')]
        [string]$SourceUserPrincipalName,
        [Parameter(Mandatory = $true)]
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

        $resolvedSource = Find-UserRecipient -UserPrincipalName $SourceUserPrincipalName
        $resolvedDestination = Find-UserRecipient -UserPrincipalName $DestinationUserPrincipalName

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
                Write-NCMessage "Usage location set to $targetUsage for $($destinationUser.UserPrincipalName)." -Level VERBOSE
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
                    Write-NCMessage "Skipping invalid destination SkuId '$sku'." -Level VERBOSE
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
                Write-NCMessage "Destination already has $($lic.SkuPartNumber); skipping add." -Level VERBOSE
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
                        Write-NCMessage "Skipping invalid disabled plan '$planString' for $($lic.SkuPartNumber)." -Level VERBOSE
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
        Remove-UserMsolAccountSku -UserPrincipalName user@contoso.com -License "18181a46-0d4e-45cd-891e-60aabd171b4e"
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByLicense')]
        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [Alias('User', 'UPN')]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true, ParameterSetName = 'ByLicense')]
        [string[]]$License,
        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [switch]$All,
        [Parameter(ParameterSetName = 'ByLicense')]
        [Parameter(ParameterSetName = 'All')]
        [switch]$ForceLicenseCatalogRefresh
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $UserPrincipalName
        if (-not $resolvedPrincipal) {
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        }

        try {
            $user = Get-MgUser -UserId $resolvedPrincipal -ErrorAction Stop
        }
        catch {
            Write-NCMessage "User $UserPrincipalName not found or query failed: $($_.Exception.Message)" -Level ERROR
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

            $removeLicenseIds = $removeLicenses.SkuId
            $licenseNames = $removeLicenses | ForEach-Object { $_.Name } | Select-Object -Unique
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
    finally {
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
