#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Licenses =============================================================================================================================

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
            Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
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
                    Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName), attempt $attempt of $max" -Level ERROR
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

function Get-UserMsolAccountSku {
    <#
    .SYNOPSIS
        Shows licenses assigned to a specific user.
    .DESCRIPTION
        Downloads the license catalog, fetches the target user via Microsoft Graph, and prints each
        assigned SKU with the mapped product name (when available).
    .PARAMETER UserPrincipalName
        Target user UPN or object ID.
    .PARAMETER ForceLicenseCatalogRefresh
        Force a fresh download of the cached license catalog before processing.
    .EXAMPLE
        Get-UserMsolAccountSku -UserPrincipalName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "User Principal Name (e.g. user@contoso.com)")]
        [string] $UserPrincipalName,
        [switch] $ForceLicenseCatalogRefresh
    )

    Set-ProgressAndInfoPreferences
    try {
        $GraphConnection = Test-MgGraphConnection
        if (-not $GraphConnection) {
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

        $resolvedRecipient = Find-UserRecipient -UserPrincipalName $UserPrincipalName
        if (-not $resolvedRecipient) {
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        } else {
            $UserPrincipalName = $resolvedRecipient
        }

        try {
            $User = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        }
        catch {
            Write-NCMessage "User $UserPrincipalName not found or query failed: $_" -Level ERROR
            return
        }

        $catalogSource = $licenseCatalog.Source
        $catalogUpdated = if ($licenseCatalog.LastCommitUtc) {
            $licenseCatalog.LastCommitUtc.ToLocalTime().ToString($NCVars.DateTimeString_Full)
        } else { $null }
        $catalogInfo = if ($catalogSource -or $catalogUpdated) {
            $parts = @()
            if ($catalogSource) { $parts += $catalogSource }
            if ($catalogUpdated) { $parts += "last updated: $catalogUpdated" }
            " (source: {0})" -f ($parts -join ', ')
        }
        else { '' }

        Write-NCMessage ("`nProcessing user: {0} <{1}>{2}`n" -f $User.DisplayName, $User.UserPrincipalName, $catalogInfo) -Level SUCCESS

        try {
            $GraphLicense = Invoke-NCRetry -Action {
                Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "retrieve licenses for $($User.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName), attempt $attempt of $max" -Level ERROR
            }
        }
        catch {
            Write-NCMessage "Failed to retrieve licenses for $($User.UserPrincipalName) after $maxAttempts attempts." -Level ERROR
            return
        }

        if ($GraphLicense -and $GraphLicense.Count -gt 0) {
            foreach ($lic in $GraphLicense) {
                $skuPart = $lic.SkuPartNumber
                $skuId = $lic.SkuId
                $matchSource = $null
                $display = Get-LicenseDisplayName -Lookup $licenseLookup -SkuPartNumber $skuPart -FallbackLookup $customLookup -MatchSource ([ref]$matchSource)
                if ($display) {
                    $suffix = if ($matchSource -and $matchSource -ne 'Primary') { ' (custom)' } else { '' }
                    Write-NCMessage ("  - {0}{2} ({1})" -f $display, $skuId, $suffix) -Level INFO
                }
                else {
                    Write-Verbose ("  - Unknown license: {0} ({1})" -f $skuPart, $skuId)
                    Write-NCMessage ("  - {0} ({1})" -f $skuPart, $skuId) -Level WARNING
                }
            }
        }
        else {
            Write-NCMessage "No licenses assigned to this user." -Level VERBOSE
        }
    }
    finally {
        Add-EmptyLine
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
            Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
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
                Write-NCMessage "Failed to retrieve tenant licenses, attempt $attempt of $max." -Level ERROR
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

function Move-MsolAccountSku {
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
            Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
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
            $sourceUser = Get-MgUser -UserId $resolvedSource -ErrorAction Stop
            $destinationUser = Get-MgUser -UserId $resolvedDestination -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to retrieve users: $($_.Exception.Message)" -Level ERROR
            return
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
                Write-NCMessage "Failed to retrieve licenses for $($sourceUser.UserPrincipalName), attempt $attempt of $max." -Level ERROR
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

        $destinationSkuIds = if ($destinationLicenses) { $destinationLicenses.SkuId } else { @() }
        $addLicenses = @()
        $removeSkuIds = @()

        foreach ($lic in $sourceLicenses) {
            $removeSkuIds += $lic.SkuId
            if ($destinationSkuIds -contains $lic.SkuId) {
                Write-NCMessage "Destination already has $($lic.SkuPartNumber); skipping add." -Level VERBOSE
                continue
            }

            $addLicenses += @{
                SkuId         = $lic.SkuId
                DisabledPlans = if ($lic.DisabledPlans) { $lic.DisabledPlans } else { @() }
            }
        }

        if ($addLicenses.Count -eq 0 -and $removeSkuIds.Count -eq 0) {
            Write-NCMessage "Nothing to move between $($sourceUser.UserPrincipalName) and $($destinationUser.UserPrincipalName)." -Level WARNING
            return
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
                    Write-NCMessage "Failed to assign licenses to $($destinationUser.UserPrincipalName), attempt $attempt of $max." -Level ERROR
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
                    Set-MgUserLicense -UserId $sourceUser.Id -AddLicenses @() -RemoveLicenses $removeSkuIds -ErrorAction Stop
                } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "remove licenses from $($sourceUser.UserPrincipalName)" -OnError {
                    param($attempt, $max, $err)
                    Write-NCMessage "Failed to remove licenses from $($sourceUser.UserPrincipalName), attempt $attempt of $max." -Level ERROR
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

function Add-MsolAccountSku {
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
            Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
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

        $defaultUsageLocation = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('UsageLocation') -and $NCVars.UsageLocation) {
            [string]$NCVars.UsageLocation
        }
        else { 'US' }

        if ([string]::IsNullOrWhiteSpace($user.UsageLocation)) {
            $targetUsage = $defaultUsageLocation
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
                Write-NCMessage "Failed to retrieve tenant licenses, attempt $attempt of $max." -Level ERROR
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

        $summary = "Assign license(s): {0} to {1}" -f (($resolved | ForEach-Object { $_.Name } | Select-Object -Unique) -join ', '), $user.UserPrincipalName
        if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, $summary)) {
            return
        }

        try {
            Invoke-NCRetry -Action {
                Set-MgUserLicense -UserId $user.Id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
            } -MaxAttempts $maxAttempts -DelaySeconds 5 -OperationDescription "assign licenses to $($user.UserPrincipalName)" -OnError {
                param($attempt, $max, $err)
                Write-NCMessage ("Failed to assign licenses to {0}, attempt {1} of {2}. {3}" -f $user.UserPrincipalName, $attempt, $max, $err.Exception.Message) -Level ERROR
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
