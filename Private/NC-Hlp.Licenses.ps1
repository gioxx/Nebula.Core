#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Licenses's Utilities =======================================================================================================

function Get-LicenseCacheInfo {
    <#
    .SYNOPSIS
        Returns license catalog cache paths.
    .DESCRIPTION
        Resolves (and optionally creates) the local directory used to persist the downloaded
        license catalog JSON files and related metadata.
    .PARAMETER Ensure
        Create the cache directory if it does not exist.
    .PARAMETER CacheFileName
        Name of the cached JSON file (defaults to M365_licenses.json).
    #>
    [CmdletBinding()]
    param(
        [switch]$Ensure,
        [string]$CacheFileName = 'M365_licenses.json'
    )

    $defaultRoot = if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('LicenseCacheDirectory') -and $NCVars.LicenseCacheDirectory) {
        [string]$NCVars.LicenseCacheDirectory
    }
    else {
        Join-Path $env:USERPROFILE '.NebulaCore\Cache'
    }

    if ($Ensure -and -not (Test-Path -LiteralPath $defaultRoot)) {
        New-Item -ItemType Directory -Path $defaultRoot -Force | Out-Null
    }

    $metaFileName = '{0}.meta.json' -f ([System.IO.Path]::GetFileNameWithoutExtension($CacheFileName))

    return @{
        Directory = $defaultRoot
        DataPath  = Join-Path $defaultRoot $CacheFileName
        MetaPath  = Join-Path $defaultRoot $metaFileName
    }
}

function Get-NormalizedLicenseKey {
    <#
    .SYNOPSIS
        Normalizes SKU identifiers for dictionary lookups.
    .DESCRIPTION
        Returns $null for blank strings; otherwise uppercases and replaces whitespace, dots and dashes with underscores.
    .PARAMETER Value
        SKU string to normalize.
    #>
    [CmdletBinding()]
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    return (($Value -replace '[-\.\s]', '_').ToUpperInvariant())
}

function New-LicenseLookup {
    [CmdletBinding()]
    param([object[]]$Items)

    $lookup = @{}
    if (-not $Items) { return $lookup }

    foreach ($item in $Items) {
        $key = Get-NormalizedLicenseKey -Value $item.String_Id
        if (-not $key) { continue }

        $names = @()
        if ($null -ne $item.Product_Display_Name) {
            if ($item.Product_Display_Name -is [System.Collections.IEnumerable] -and -not ($item.Product_Display_Name -is [string])) {
                foreach ($name in $item.Product_Display_Name) {
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        $names += $name
                    }
                }
            }
            else {
                $names += $item.Product_Display_Name
            }
        }

        $names = $names | Where-Object { $_ -and ($_ -ne '') } | Select-Object -Unique
        if ($names.Count -gt 0) {
            $lookup[$key] = ($names -join ' / ')
        }
    }

    return $lookup
}

function Get-LicenseSourceData {
    <#
    .SYNOPSIS
        Retrieves (and caches) license data from a specific source.
    .DESCRIPTION
        Handles cache validation, optional metadata retrieval, download, and persistence for a JSON feed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CacheFileName,
        [Parameter(Mandatory)]
        [string]$FileUrl,
        [string]$ApiUrl,
        [switch]$ForceRefresh,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 5,
        [int]$CacheDays = 7
    )

    if ([string]::IsNullOrWhiteSpace($FileUrl)) {
        return $null
    }

    $cacheInfo = Get-LicenseCacheInfo -Ensure -CacheFileName $CacheFileName
    $cacheFile = $cacheInfo.DataPath
    $metaFile  = $cacheInfo.MetaPath

    $ttl = [TimeSpan]::FromDays([Math]::Max(1, $CacheDays))
    $nowUtc = (Get-Date).ToUniversalTime()

    $meta = $null
    if (Test-Path -LiteralPath $metaFile) {
        try {
            $meta = Get-Content -LiteralPath $metaFile -Raw | ConvertFrom-Json
        }
        catch {
            $meta = $null
        }
    }

    $tryParseUtc = {
        param($value)
        if (-not $value) { return $null }
        try { return [DateTime]::Parse($value, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal) }
        catch { return $null }
    }

    $currentCommitUtc = if ($meta) { & $tryParseUtc $meta.LastCommitUtc }
    $lastCheckedUtc = if ($meta) { & $tryParseUtc $meta.LastCheckedUtc }
    if (-not $lastCheckedUtc -and (Test-Path -LiteralPath $cacheFile)) {
        $lastCheckedUtc = (Get-Item -LiteralPath $cacheFile).LastWriteTimeUtc
    }

    $needDownload = $ForceRefresh.IsPresent -or -not (Test-Path -LiteralPath $cacheFile)
    if (-not $needDownload -and $lastCheckedUtc) {
        if ($nowUtc - $lastCheckedUtc -ge $ttl) {
            $needDownload = $true
        }
    }

    $remoteCommitUtc = $null
    if ($ApiUrl -and ($needDownload -or $ForceRefresh.IsPresent)) {
        try {
            $response = Invoke-NCRetry -Action {
                Invoke-RestMethod -Uri $ApiUrl -Headers @{ 'User-Agent' = 'Nebula.Core' } -ErrorAction Stop
            } -MaxAttempts $MaxAttempts -DelaySeconds $DelaySeconds -OperationDescription "retrieve license metadata" -OnError {
                param($attempt, $max, $err)
                Write-NCMessage "Failed to retrieve license metadata, attempt $attempt of $max." -Level WARNING
            }

            if ($response -and $response[0]) {
                $lastCommitDate = $response[0].commit.committer.date
                try {
                    $remoteCommitUtc = [DateTime]::Parse($lastCommitDate, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal)
                }
                catch {
                    $remoteCommitUtc = $null
                }
            }
        }
        catch {
            if ($needDownload -and -not (Test-Path -LiteralPath $cacheFile)) {
                throw "Unable to contact GitHub to retrieve license metadata or catalog. $($_.Exception.Message)"
            }
        }
    }

    if ($needDownload -and -not $ForceRefresh.IsPresent -and $remoteCommitUtc -and $currentCommitUtc) {
        if ($remoteCommitUtc -le $currentCommitUtc) {
            $needDownload = $false
        }
    }

    $licenseItems = $null
    $source = 'Cache'

    if (-not $needDownload) {
        try {
            $licenseItems = Get-Content -LiteralPath $cacheFile -Raw | ConvertFrom-Json
        }
        catch {
            Write-NCMessage "Unable to read cached license catalog ($CacheFileName). Attempting to download a fresh copy ..." -Level WARNING
            $needDownload = $true
        }
    }

    if ($needDownload) {
        try {
            $licenseItems = Invoke-NCRetry -Action {
                Invoke-RestMethod -Method Get -Uri $FileUrl -ErrorAction Stop
            } -MaxAttempts $MaxAttempts -DelaySeconds $DelaySeconds -OperationDescription "download license file" -OnError {
                param($attempt, $max, $err)
                Write-NCMessage "Failed downloading license file ($CacheFileName), attempt $attempt of $max." -Level ERROR
            }

            $licenseItems | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $cacheFile -Encoding UTF8
            $source = 'Remote'
            Write-NCMessage "License file downloaded and cached at $cacheFile." -Level VERBOSE

            if (-not $remoteCommitUtc) {
                $remoteCommitUtc = $nowUtc
            }
            $currentCommitUtc = $remoteCommitUtc
        }
        catch {
            throw "Downloading license file failed after $MaxAttempts attempts."
        }
    }

    if (-not $currentCommitUtc -and $remoteCommitUtc) {
        $currentCommitUtc = $remoteCommitUtc
    }

    $metaObject = [ordered]@{
        LastCheckedUtc = $nowUtc.ToString('o')
        LastCommitUtc  = if ($currentCommitUtc) { $currentCommitUtc.ToString('o') } else { $null }
    }
    try {
        $metaObject | ConvertTo-Json | Set-Content -LiteralPath $metaFile -Encoding UTF8
    }
    catch {
        Write-NCMessage "Unable to update license cache metadata for $($CacheFileName): $($_.Exception.Message)" -Level WARNING
    }

    return [pscustomobject]@{
        Items         = $licenseItems
        Source        = $source
        CachePath     = $cacheFile
        LastCommitUtc = $currentCommitUtc
    }
}

function Get-LicenseCatalog {
    <#
    .SYNOPSIS
        Retrieves (and caches) the license catalog JSON (with custom fallback).
    .DESCRIPTION
        Loads the primary catalog from a local cache when possible, refreshes it from GitHub when forced
        or when stale, and optionally loads a custom catalog to resolve missing SKUs.
    .PARAMETER IncludeMetadata
        Adds last commit/update details to the output (and logs them).
    .PARAMETER ForceRefresh
        Forces a re-download of the catalog(s) regardless of cache age.
    .PARAMETER MaxAttempts
        Maximum number of retries while calling the remote endpoints.
    .PARAMETER DelaySeconds
        Delay between retries.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeMetadata,
        [switch]$ForceRefresh,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 5
    )

    $cacheDays = 7
    if (($NCVars -is [System.Collections.IDictionary]) -and $NCVars.Contains('LicenseCacheDays') -and $NCVars.LicenseCacheDays) {
        [void][int]::TryParse([string]$NCVars.LicenseCacheDays, [ref]$cacheDays)
        if ($cacheDays -lt 1) { $cacheDays = 1 }
    }

    $primarySource = $script:NCLicenseSources.Primary
    if (-not $primarySource) {
        throw "Primary license source configuration missing."
    }

    $primaryData = Get-LicenseSourceData -CacheFileName $primarySource.CacheFileName `
        -FileUrl $primarySource.FileUrl `
        -ApiUrl $primarySource.ApiUrl `
        -ForceRefresh:$ForceRefresh.IsPresent `
        -MaxAttempts $MaxAttempts `
        -DelaySeconds $DelaySeconds `
        -CacheDays $cacheDays

    if (-not $primaryData -or -not $primaryData.Items) {
        throw "Unable to load the primary license catalog."
    }

    $primaryLookup = New-LicenseLookup -Items $primaryData.Items

    $customLookup = $null
    $customData = $null
    $customSource = $script:NCLicenseSources.Custom
    if ($customSource -and $customSource.FileUrl) {
        $customData = Get-LicenseSourceData -CacheFileName $customSource.CacheFileName `
            -FileUrl $customSource.FileUrl `
            -ApiUrl $customSource.ApiUrl `
            -ForceRefresh:$ForceRefresh.IsPresent `
            -MaxAttempts $MaxAttempts `
            -DelaySeconds $DelaySeconds `
            -CacheDays $cacheDays

        if ($customData -and $customData.Items) {
            $customLookup = New-LicenseLookup -Items $customData.Items
        }
    }

    if ($IncludeMetadata.IsPresent -and $primaryData.LastCommitUtc) {
        Write-Verbose "License catalog last updated: $($primaryData.LastCommitUtc.ToLocalTime().ToString($NCVars.DateTimeString_Full)) (source: $primaryData.Source)"
    }
    if ($IncludeMetadata.IsPresent -and $customData -and $customData.LastCommitUtc) {
        Write-Verbose "Custom license catalog last updated: $($customData.LastCommitUtc.ToLocalTime().ToString($NCVars.DateTimeString_Full)) (source: $customData.Source)"
    }

    return [pscustomobject]@{
        Items               = $primaryData.Items
        Lookup              = $primaryLookup
        LastCommitUtc       = $primaryData.LastCommitUtc
        Source              = $primaryData.Source
        CachePath           = $primaryData.CachePath
        CustomLookup        = $customLookup
        CustomLastCommitUtc = $customData?.LastCommitUtc
        CustomSource        = $customData?.Source
        CustomCachePath     = $customData?.CachePath
    }
}

function Get-LicenseDisplayName {
    <#
    .SYNOPSIS
        Resolves a SKU part number into a friendly display name.
    .DESCRIPTION
        Uses the lookup dictionary built by Get-LicenseCatalog to find mapped names, with optional fallback lookup.
    .PARAMETER Lookup
        Hashtable that maps normalized SKU keys to friendly product names.
    .PARAMETER SkuPartNumber
        SKU string to look up.
    .PARAMETER FallbackLookup
        Secondary lookup (e.g. custom catalog) used when the primary lookup does not contain the SKU.
    .PARAMETER MatchSource
        [ref] string that receives the source used: 'Primary', 'Fallback', or $null when unresolved.
    .PARAMETER FallbackSourceLabel
        Label assigned to the fallback lookup when MatchSource is requested (defaults to 'Custom').
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Lookup,
        [Parameter(Mandatory)]
        [string]$SkuPartNumber,
        [hashtable]$FallbackLookup,
        [ref]$MatchSource,
        [string]$FallbackSourceLabel = 'Custom'
    )

    $key = Get-NormalizedLicenseKey -Value $SkuPartNumber
    if (-not $key) {
        if ($PSBoundParameters.ContainsKey('MatchSource')) { $MatchSource.Value = $null }
        return $null
    }

    if ($Lookup -and $Lookup.ContainsKey($key)) {
        if ($PSBoundParameters.ContainsKey('MatchSource')) { $MatchSource.Value = 'Primary' }
        return $Lookup[$key]
    }

    if ($FallbackLookup -and $FallbackLookup.ContainsKey($key)) {
        if ($PSBoundParameters.ContainsKey('MatchSource')) { $MatchSource.Value = $FallbackSourceLabel }
        return $FallbackLookup[$key]
    }

    if ($PSBoundParameters.ContainsKey('MatchSource')) { $MatchSource.Value = $null }
    return $null
}
