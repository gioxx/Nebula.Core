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
                $arr_MsolAccountSku | Export-CSV $CSV -NoTypeInformation -DefaultLimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding) -Append
            }
        }

        $arr_MsolAccountSku | Export-CSV $CSV -NoTypeInformation -DefaultLimiter $($NCVars.CSV_DefaultLimiter) -Encoding $($NCVars.CSV_Encoding)

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
        [switch]$ForceLicenseCatalogRefresh
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

        Write-NCMessage ("`nProcessing user: {0} <{1}>{2}`n" -f $User.DisplayName, $User.UserPrincipalName, $catalogInfo) -Level VERBOSE

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
                    Write-NCMessage ("  - {0}{2} ({1})" -f $display, $skuId, $suffix) -Level SUCCESS
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
