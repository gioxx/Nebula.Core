#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Intune helpers =======================================================================================================================

function Get-IntuneProfileAssignmentsByGroup {
    <#
    .SYNOPSIS
        Shows where an Entra group is used in Intune assignments.
    .DESCRIPTION
        Searches Intune device configurations, settings catalog policies, and apps for assignments that target
        the specified Entra group. Supports lookup by group name or group ID, with optional filtering by profile
        name or profile ID. Can also include parent groups that contain the requested group as a member.
    .PARAMETER GroupName
        Target Entra group display name. Accepts pipeline input.
    .PARAMETER GroupId
        Target Entra group object ID. Use this instead of GroupName.
    .PARAMETER ProfileName
        Optional profile or app display name filter.
    .PARAMETER ProfileId
        Optional filter for a specific Intune object ID.
    .PARAMETER IncludeNestedGroups
        Also match parent groups that include the requested Entra group.
    .PARAMETER GridView
        Show additional details in Out-GridView.
    .PARAMETER Diagnostic
        Include diagnostic columns in the returned objects.
    .EXAMPLE
        Get-IntuneProfileAssignmentsByGroup -GroupName "Windows 11 Pilot"
    .EXAMPLE
        Get-IntuneProfileAssignmentsByGroup -GroupId "00000000-0000-0000-0000-000000000000"
    .EXAMPLE
        "Windows 11 Pilot" | Get-IntuneProfileAssignmentsByGroup -GridView
    .EXAMPLE
        Get-IntuneProfileAssignmentsByGroup -GroupName "Intune - Reception" -IncludeNestedGroups
    .EXAMPLE
        Get-IntuneProfileAssignmentsByGroup -GroupName "Intune - Reception" -ProfileName "Zoom Workplace" -Diagnostic
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Group', 'DisplayName', 'Name', 'Identity')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
        [string]$GroupId,

        [string]$ProfileName,
        [string]$ProfileId,
        [switch]$IncludeNestedGroups,
        [switch]$GridView,
        [switch]$Diagnostic
    )

    process {
        Invoke-NCIntuneGroupUsageCore -ParameterSetName $PSCmdlet.ParameterSetName -GroupName $GroupName -GroupId $GroupId -ProfileName $ProfileName -ProfileId $ProfileId -IncludeNestedGroups:$IncludeNestedGroups -GridView:$GridView -Diagnostic:$Diagnostic
    }
}

function Search-IntuneProfileLocation {
    <#
    .SYNOPSIS
        Finds where an Intune profile lives across multiple Microsoft Graph surfaces.
    .DESCRIPTION
        Connects to Microsoft Graph and searches a curated set of Intune endpoints for profile names
        matching the provided text. Use this command to identify the correct source before querying
        assignments or extending support for new profile families.
    .PARAMETER SearchText
        Profile name text to search for.
    .PARAMETER Exact
        Match the profile name exactly instead of using a contains search.
    .PARAMETER GridView
        Show the results in Out-GridView instead of returning objects.
    .EXAMPLE
        Search-IntuneProfileLocation -SearchText "iOS - Wi-Fi M-Smartphone"
    .EXAMPLE
        Search-IntuneProfileLocation -SearchText "Wi-Fi" -GridView
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'DisplayName', 'ProfileName', 'Query')]
        [string]$SearchText,
        [switch]$Exact,
        [switch]$GridView
    )

    begin {
        $graphConnected = $null
    }

    process {
        if ($null -eq $graphConnected) {
            $graphConnected = Test-MgGraphConnection -Scopes @('DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All', 'Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
            if (-not $graphConnected) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
                Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
                return
            }
        }

        if ([string]::IsNullOrWhiteSpace($SearchText)) {
            Write-NCMessage "SearchText cannot be empty." -Level WARNING
            return
        }

        $endpoints = @(
            @{ Source = 'deviceConfigurations'; Uris = @('v1.0/deviceManagement/deviceConfigurations?$top=100') },
            @{ Source = 'betaDeviceConfigurations'; Uris = @('beta/deviceManagement/deviceConfigurations?$top=100') },
            @{ Source = 'configurationPolicies'; Uris = @('v1.0/deviceManagement/configurationPolicies?$top=100', 'beta/deviceManagement/configurationPolicies?$top=100') },
            @{ Source = 'groupPolicyConfigurations'; Uris = @('beta/deviceManagement/groupPolicyConfigurations?$top=100') },
            @{ Source = 'resourceAccessProfiles'; Uris = @('beta/deviceManagement/resourceAccessProfiles?$top=100') },
            @{ Source = 'deviceCompliancePolicies'; Uris = @('v1.0/deviceManagement/deviceCompliancePolicies?$top=100') },
            @{ Source = 'deviceEnrollmentConfigurations'; Uris = @('v1.0/deviceManagement/deviceEnrollmentConfigurations?$top=100') },
            @{ Source = 'deviceHealthScripts'; Uris = @('beta/deviceManagement/deviceHealthScripts?$top=100') },
            @{ Source = 'deviceManagementScripts'; Uris = @('beta/deviceManagement/deviceManagementScripts?$top=100') },
            @{ Source = 'deviceShellScripts'; Uris = @('beta/deviceManagement/deviceShellScripts?$top=100') }
        )

        $results = [System.Collections.Generic.List[object]]::new()
        $normalizedSearch = $SearchText.Trim()

        foreach ($endpoint in $endpoints) {
            $items = @()
            $queried = $false
            $lastError = $null

            foreach ($endpointUri in $endpoint.Uris) {
                try {
                    $items = @(Invoke-NCGraphCollectionRequest -Uri $endpointUri)
                    $queried = $true
                    break
                }
                catch {
                    $lastError = $_.Exception.Message
                }
            }

            if (-not $queried) {
                if ($lastError) {
                    Write-NCMessage "Unable to query $($endpoint.Source): $lastError" -Level WARNING
                }
                continue
            }

            foreach ($item in $items) {
                $itemName = Get-NCIntuneItemName -Item $item
                if ([string]::IsNullOrWhiteSpace($itemName)) {
                    continue
                }

                $isMatch = if ($Exact.IsPresent) {
                    $itemName -eq $normalizedSearch
                }
                else {
                    $itemName -like "*$normalizedSearch*"
                }

                if (-not $isMatch) {
                    continue
                }

                $results.Add([pscustomobject][ordered]@{
                        'Profile Name' = $itemName
                        'Source'       = $endpoint.Source
                        'Profile Id'   = Get-NCIntuneItemId -Item $item
                        'Profile Type' = Get-NCIntuneItemODataType -Item $item
                    }) | Out-Null
            }
        }

        Add-EmptyLine
        Write-Verbose "Intune profiles found for '$normalizedSearch': $($results.Count)"

        if ($results.Count -eq 0) {
            Write-NCMessage "No Intune profiles found for '$normalizedSearch' in the currently scanned endpoints." -Level WARNING
            return
        }

        $sorted = $results | Sort-Object 'Profile Name', 'Source' -Unique
        if ($GridView.IsPresent) {
            $sorted | Out-GridView -Title "Intune Profile Search - $normalizedSearch"
        }
        else {
            $sorted
        }
    }
}

function Export-IntuneAppInventory {
    <#
    .SYNOPSIS
        Reports Intune-managed devices that have matching applications installed.
    .DESCRIPTION
        Connects to Microsoft Graph, scans managed devices for detected apps, and can optionally
        enrich the report with deployed app device status information. The output is report-friendly
        and can also be exported to CSV and/or JSON.
    .PARAMETER ApplicationName
        Application name or wildcard pattern to match. Accepts pipeline input.
    .PARAMETER MinimumVersion
        Minimum application version to keep in the report.
    .PARAMETER FilterByType
        Optional app type filter when deployed app data is included.
    .PARAMETER FilterByPlatform
        Optional device platform filter.
    .PARAMETER LastInventory
        Include the managed device last Intune sync date in the report.
    .PARAMETER OnlySuccessfulInstalls
        When deployed app data is included, keep only successful installs.
    .PARAMETER IncludeDeployedApps
        Also query deployed app device statuses in addition to detected apps.
    .PARAMETER MaxDevices
        Maximum number of devices to process. Use 0 for all devices.
    .PARAMETER OutputCsvPath
        Optional CSV export path.
    .PARAMETER OutputJsonPath
        Optional JSON export path.
    .PARAMETER BatchSize
        Number of rows to flush at a time when writing CSV output.
    .PARAMETER Resume
        Resume CSV export from the latest matching CSV or from -CsvPath.
    .PARAMETER CsvPath
        Explicit CSV file to resume or export to. When omitted, a default file is used.
    .PARAMETER MaxConsecutiveErrors
        Stop after this many consecutive device-level failures.
    .PARAMETER PivotSummary
        Print a per-app summary after the report is built.
    .EXAMPLE
        Export-IntuneAppInventory -ApplicationName "TeamViewer"
    .EXAMPLE
        Export-IntuneAppInventory -ApplicationName "Microsoft*" -IncludeDeployedApps -FilterByType Win32 -OutputCsvPath "apps.csv"
    .EXAMPLE
        Export-IntuneAppInventory -ApplicationName "*java*" -FilterByPlatform Windows -LastInventory
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('SearchText', 'Name', 'DisplayName', 'Query', 'AppName')]
        [string]$ApplicationName,

        [string]$MinimumVersion,

        [ValidateSet('Win32', 'Store', 'LOB', 'Web', 'iOS', 'Android', 'macOS', 'All')]
        [string]$FilterByType = 'All',

        [ValidateSet('Windows', 'iOS', 'Android', 'macOS', 'All')]
        [string]$FilterByPlatform = 'All',

        [switch]$OnlySuccessfulInstalls,
        [switch]$IncludeDeployedApps,
        [switch]$LastInventory,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$MaxDevices = 0,

        [string]$OutputCsvPath,
        [string]$OutputJsonPath,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25,
        [switch]$Resume,
        [string]$CsvPath,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5,
        [switch]$PivotSummary
    )

    begin {
        $graphConnected = $null
    }

    process {
        try {
            if ($null -eq $graphConnected) {
                $graphConnected = Test-MgGraphConnection -Scopes @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementApps.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
                if (-not $graphConnected) {
                    Add-EmptyLine
                    Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                    return
                }

                if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
                    Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
                    return
                }
            }

            if ([string]::IsNullOrWhiteSpace($ApplicationName)) {
                Write-NCMessage "ApplicationName cannot be empty." -Level WARNING
                return
            }

            $normalizeText = {
                param($value)
                return Get-NormalizedText -Value $value
            }
            $buildRowKey = {
                param($row)

                $app = & $normalizeText $row.AppName
                $device = & $normalizeText $row.DeviceId
                $source = & $normalizeText $row.Source
                $installState = & $normalizeText $row.InstallState
                $version = & $normalizeText $row.Version
                $publisher = & $normalizeText $row.Publisher
                return "{0}|{1}|{2}|{3}|{4}|{5}" -f $app, $device, $source, $installState, $version, $publisher
            }
            $existingRowKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $exportCsv = $false
            $csvPathResolved = $null
            $processedSinceFlush = 0
            $consecutiveErrors = 0
            $aborted = $false

            Write-NCMessage "Starting app inventory reporting ..." -Level INFO

            # Pull devices
            $devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,userPrincipalName,lastSyncDateTime"
            if ($MaxDevices -gt 0) {
                $devicesUri += "&`$top=$MaxDevices"
            }
            $devices = @(Invoke-NCGraphAllPagesCore -Uri $devicesUri)
            $lastInventoryCache = @{}
            if ($LastInventory) {
                foreach ($device in $devices) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$device.lastSyncDateTime)) {
                        $lastInventoryCache[[string]$device.id] = Format-NCDateTime -Value $device.lastSyncDateTime -AsLocalTime
                    }
                }
            }
            if ($FilterByPlatform -ne "All") {
                $devices = @($devices | Where-Object { $_.operatingSystem -like "$FilterByPlatform*" })
            }
            Write-NCMessage "Managed devices retrieved: $($devices.Count)" -Level INFO

            $lastInventoryFailed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            # Devices already read in v1.0 (even with a blank last-sync date): never requested again
            $lastInventoryChecked = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            # Reads missing last-sync dates for the given devices in Graph batches (v1.0, 20 per request).
            $loadLastInventory = {
                param([string[]]$DeviceIds)

                if (-not $LastInventory) {
                    return
                }

                $missing = @($DeviceIds | Where-Object {
                        -not [string]::IsNullOrWhiteSpace($_) -and
                        -not ($lastInventoryCache.ContainsKey($_) -and -not [string]::IsNullOrWhiteSpace([string]$lastInventoryCache[$_])) -and
                        -not $lastInventoryFailed.Contains($_) -and
                        -not $lastInventoryChecked.Contains($_)
                    } | Select-Object -Unique)
                if ($missing.Count -eq 0) {
                    return
                }

                $detailRequests = @(for ($i = 0; $i -lt $missing.Count; $i++) {
                        @{ Id = "l$i"; Method = 'GET'; Url = "/deviceManagement/managedDevices/$([uri]::EscapeDataString($missing[$i]))?`$select=lastSyncDateTime" }
                    })
                $detailResponses = @(Invoke-NCGraphBatch -Requests $detailRequests -ApiVersion 'v1.0' -Activity 'Reading last inventory dates')
                for ($i = 0; $i -lt $missing.Count; $i++) {
                    $detailResponse = $detailResponses[$i]
                    $detailError = $detailResponse.ErrorMessage
                    if ($detailResponse.Success) {
                        try {
                            $lastInventoryValue = $detailResponse.Body.lastSyncDateTime
                            if (-not [string]::IsNullOrWhiteSpace([string]$lastInventoryValue)) {
                                $lastInventoryCache[$missing[$i]] = Format-NCDateTime -Value $lastInventoryValue -AsLocalTime
                            }
                            $null = $lastInventoryChecked.Add($missing[$i])
                        }
                        catch {
                            $detailError = $_.Exception.Message
                        }
                    }
                    if ($detailError) {
                        $null = $lastInventoryFailed.Add($missing[$i])
                        Write-NCMessage "Unable to read last inventory date for device $($missing[$i]): $detailError" -Level WARNING
                    }
                }
            }

            function Get-NCIntuneManagedDeviceLastInventory {
                param(
                    [Parameter(Mandatory = $true)]
                    [string]$DeviceId
                )

                if (-not $LastInventory) {
                    return $null
                }

                if ($lastInventoryCache.ContainsKey($DeviceId) -and -not [string]::IsNullOrWhiteSpace([string]$lastInventoryCache[$DeviceId])) {
                    return [string]$lastInventoryCache[$DeviceId]
                }

                return $null
            }

            # Build app --> device mapping from Detected Apps
            $appDeviceMap = @{}
            $processed = 0

            $batchNoticeWritten = $false
            if ($devices.Count -gt 0) {
                $batchNoticeWritten = Write-NCGraphBatchNotice -Count ($devices.Count) -Noun 'device(s)' -PassThru
            }

            for ($offset = 0; $offset -lt $devices.Count -and -not $aborted; $offset += 20) {
                $deviceChunk = @($devices[$offset..([Math]::Min($offset + 20, $devices.Count) - 1)])

                $Percentage = Get-NCProgressPercent -Current $processed -Total $devices.Count
                Write-Progress -Activity "Reading Detected Apps" -Status "$($deviceChunk[0].deviceName) - $processed / $($devices.Count) devices - $Percentage%" -PercentComplete $Percentage

                $appRequests = @(for ($i = 0; $i -lt $deviceChunk.Count; $i++) {
                        @{ Id = "d$i"; Method = 'GET'; Url = "/deviceManagement/managedDevices/$([uri]::EscapeDataString([string]$deviceChunk[$i].id))?`$expand=detectedApps" }
                    })
                $appResponses = @(Invoke-NCGraphBatch -Requests $appRequests -ApiVersion 'beta' -Activity 'Reading detected apps')

                # Matching apps per device (after name and version filters), then last-sync dates for the matched devices in one go
                $matchedApps = @{}
                for ($i = 0; $i -lt $deviceChunk.Count; $i++) {
                    if (-not $appResponses[$i].Success) { continue }
                    $matchedApps[$i] = @(foreach ($app in ($appResponses[$i].Body.detectedApps | Where-Object { $_.displayName -like $ApplicationName })) {
                            if ($MinimumVersion -and $app.version) {
                                if (-not (Test-NCIntuneVersionAtLeast -CurrentVersion $app.version -MinimumVersion $MinimumVersion)) {
                                    continue
                                }
                            }
                            $app
                        })
                }
                if ($LastInventory) {
                    & $loadLastInventory -DeviceIds @(for ($i = 0; $i -lt $deviceChunk.Count; $i++) {
                            if ($matchedApps.ContainsKey($i) -and $matchedApps[$i].Count -gt 0) { [string]$deviceChunk[$i].id }
                        })
                }

                for ($i = 0; $i -lt $deviceChunk.Count; $i++) {
                    $device = $deviceChunk[$i]
                    $processed++
                    $Percentage = Get-NCProgressPercent -Current $processed -Total $devices.Count

                    if (-not $appResponses[$i].Success) {
                        Write-NCMessage "Error reading apps for $($device.deviceName): $($appResponses[$i].ErrorMessage)" -Level WARNING
                        $consecutiveErrors++
                        if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                            $aborted = $true
                            break
                        }
                        continue
                    }

                    foreach ($app in $matchedApps[$i]) {
                        Write-Progress -Activity "Reading Detected Apps" -Status "$($device.deviceName) / $($app.displayName) - $processed / $($devices.Count) devices - $Percentage%" -PercentComplete $Percentage

                        $key = $app.displayName
                        if (-not $appDeviceMap.ContainsKey($key)) {
                            $appDeviceMap[$key] = [ordered]@{ Devices = @(); Versions = @{}; Publishers = @{} }
                        }

                        $deviceRow = [ordered]@{
                            DeviceId   = $device.id
                            DeviceName = $device.deviceName
                            Platform   = $device.operatingSystem
                            User       = $device.userPrincipalName
                            Version    = $app.version
                            Publisher  = $app.publisher
                            Source     = "DetectedApps"
                        }
                        if ($LastInventory) {
                            $deviceRow.LastInventory = Get-NCIntuneManagedDeviceLastInventory -DeviceId $device.id
                        }
                        $appDeviceMap[$key].Devices += $deviceRow

                        if ($app.version) {
                            $appDeviceMap[$key].Versions[$app.version] = ($appDeviceMap[$key].Versions[$app.version] + 1)
                        }

                        if ($app.publisher) {
                            $appDeviceMap[$key].Publishers[$app.publisher] = ($appDeviceMap[$key].Publishers[$app.publisher] + 1)
                        }
                    }
                }
            }
            Write-Progress -Activity "Reading Detected Apps" -Completed

            # Optionally incorporate deployment statuses (broadens coverage)
            if ($IncludeDeployedApps) {
                Write-NCMessage "Including deployed apps device status ..." -Level INFO
                $appsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"
                $allApps = @(Invoke-NCGraphAllPagesCore -Uri $appsUri)

                $deployedCandidates = @(foreach ($app in $allApps | Where-Object { $_.displayName -like $ApplicationName }) {
                        $appType = Get-NCIntuneAppTypeFromODataType -ODataType $app.'@odata.type'
                        if ($FilterByType -ne "All" -and $appType -ne $FilterByType) {
                            continue
                        }
                        [pscustomobject]@{ App = $app; AppType = $appType }
                    })

                if ($deployedCandidates.Count -gt 0 -and -not $batchNoticeWritten) {
                    $batchNoticeWritten = Write-NCGraphBatchNotice -Count ($deployedCandidates.Count) -Noun 'app(s)' -PassThru
                }

                # Deployment statuses in Graph batches (beta, 20 apps per request), in app order
                $deployedStatuses = [System.Collections.Generic.List[object]]::new()
                for ($offset = 0; $offset -lt $deployedCandidates.Count; $offset += 20) {
                    $candidateChunk = @($deployedCandidates[$offset..([Math]::Min($offset + 20, $deployedCandidates.Count) - 1)])
                    $statusRequests = @(for ($i = 0; $i -lt $candidateChunk.Count; $i++) {
                            @{ Id = "s$i"; Method = 'GET'; Url = "/deviceAppManagement/mobileApps/$([uri]::EscapeDataString([string]$candidateChunk[$i].App.id))/deviceStatuses" }
                        })
                    $statusResponses = @(Invoke-NCGraphBatchCollection -Requests $statusRequests -ApiVersion 'beta' -Activity 'Reading app deployment status')

                    $chunkEntries = @(for ($i = 0; $i -lt $candidateChunk.Count; $i++) {
                            $chunkStatuses = @()
                            if ($statusResponses[$i].Success) {
                                $chunkStatuses = @($statusResponses[$i].Items)
                            }
                            else {
                                Write-NCMessage "Error fetching data: $($statusResponses[$i].ErrorMessage)" -Level WARNING
                            }
                            [pscustomobject]@{ App = $candidateChunk[$i].App; AppType = $candidateChunk[$i].AppType; Statuses = $chunkStatuses }
                        })

                    if ($LastInventory) {
                        $statusDeviceIds = @(foreach ($entry in $chunkEntries) {
                                $entry.Statuses | Where-Object { -not ($OnlySuccessfulInstalls -and $_.installState -ne "installed") } | ForEach-Object { [string]$_.deviceId } | Where-Object { $devices.id -contains $_ }
                            })
                        & $loadLastInventory -DeviceIds $statusDeviceIds
                    }

                    foreach ($entry in $chunkEntries) {
                        $deployedStatuses.Add($entry) | Out-Null
                    }
                }

                foreach ($entry in $deployedStatuses) {
                    $app = $entry.App
                    $appType = $entry.AppType
                    $statuses = $entry.Statuses

                    foreach ($s in $statuses) {
                        if ($OnlySuccessfulInstalls -and $s.installState -ne "installed") {
                            continue
                        }

                        $d = $devices | Where-Object { $_.id -eq $s.deviceId } | Select-Object -First 1
                        if (-not $d) {
                            continue
                        }

                        $key = $app.displayName
                        if (-not $appDeviceMap.ContainsKey($key)) {
                            $appDeviceMap[$key] = [ordered]@{ Devices = @(); Versions = @{}; Publishers = @{} }
                        }

                        # Avoid duplicates for the same device/app when DetectedApps already included it
                        $exists = $appDeviceMap[$key].Devices | Where-Object { $_.DeviceId -eq $d.id }
                        if (-not $exists) {
                            $deviceRow = [ordered]@{
                                DeviceId     = $d.id
                                DeviceName   = $d.deviceName
                                Platform     = $d.operatingSystem
                                User         = $d.userPrincipalName
                                Version      = $null
                                Publisher    = $null
                                InstallState = $s.installState
                                AppType      = $appType
                                Source       = "DeploymentStatus"
                            }
                            if ($LastInventory) {
                                $deviceRow.LastInventory = Get-NCIntuneManagedDeviceLastInventory -DeviceId $d.id
                            }
                            $appDeviceMap[$key].Devices += $deviceRow
                        }
                    }
                }
            }

            # Optional app type filter when only Detected Apps were used
            if (-not $IncludeDeployedApps -and $FilterByType -ne "All") {
                Write-Verbose "FilterByType applies only when -IncludeDeployedApps is used. Skipping type filter for Detected Apps only."
            }

            # Build flat rows
            $rows = @()
            foreach ($appName in $appDeviceMap.Keys) {
                foreach ($dev in $appDeviceMap[$appName].Devices) {
                    $rows += [pscustomobject]@{
                        AppName      = $appName
                        Version      = $dev.Version
                        Publisher    = $dev.Publisher
                        AppType      = $dev.AppType
                        DeviceName   = $dev.DeviceName
                        DeviceId     = $dev.DeviceId
                        Platform     = $dev.Platform
                        LastInventory = $dev.LastInventory
                        User         = $dev.User
                        InstallState = $dev.InstallState
                        Source       = $dev.Source
                    }
                }
            }

            if ($OutputCsvPath -or $Resume.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvPath)) {
                if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                    $csvPathResolved = $CsvPath
                }
                elseif (-not [string]::IsNullOrWhiteSpace($OutputCsvPath)) {
                    $csvPathResolved = $OutputCsvPath
                }
                elseif ($Resume) {
                    $existingCsv = Get-ChildItem -Path (Get-Location).Path -File -Filter "*_M365-IntuneAppInventory-Report.csv" |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
                    if ($existingCsv) {
                        $csvPathResolved = $existingCsv.FullName
                    }
                    else {
                        $csvPathResolved = New-File (Join-Path -Path (Get-Location).Path -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-IntuneAppInventory-Report.csv")
                    }
                }

                if ($Resume -and (Test-Path -LiteralPath $csvPathResolved)) {
                    try {
                        foreach ($row in (Import-CSV -LiteralPath $csvPathResolved -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                            $null = $existingRowKeys.Add((& $buildRowKey $row))
                        }
                        Write-NCMessage ("Resuming Intune app inventory export from {0}; {1} row(s) already recorded." -f $csvPathResolved, $existingRowKeys.Count) -Level INFO
                    }
                    catch {
                        Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csvPathResolved, $_.Exception.Message) -Level WARNING
                        $existingRowKeys.Clear()
                    }
                }
                elseif ($Resume -and -not (Test-Path -LiteralPath $csvPathResolved)) {
                    Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csvPathResolved) -Level INFO
                }

                $exportCsv = $true
                Write-NCMessage ("Intune app inventory CSV export will flush every {0} row(s). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
                Write-NCMessage "Saving report to $csvPathResolved" -Level DEBUG
            }

            if (-not $rows -or $rows.Count -eq 0) {
                if ($exportCsv -and (Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    Write-NCMessage "No new Intune app inventory rows found. Existing CSV at $csvPathResolved already contains the requested rows." -Level INFO
                }
                else {
                    Write-NCMessage "No applications matched '$ApplicationName' with the provided filters." -Level WARNING
                }
                return
            }

            $displayColumns = @('AppName', 'Version', 'DeviceName')
            $exportColumns = @('AppName', 'Version', 'Publisher', 'DeviceName', 'DeviceId')
            if ($IncludeDeployedApps) {
                $exportColumns += 'AppType'
            }
            if ($FilterByPlatform -eq 'All') {
                $displayColumns += 'Platform'
                $exportColumns += 'Platform'
            }
            if ($LastInventory) {
                $displayColumns += 'LastInventory'
                $exportColumns += 'LastInventory'
            }
            $displayColumns += @('User', 'Source')
            $exportColumns += 'User'
            if ($IncludeDeployedApps) {
                $exportColumns += 'InstallState'
            }
            $exportColumns += 'Source'
            $reportRows = @($rows | Select-Object -Property $exportColumns)

            # Console table output
            $reportRows | Sort-Object AppName, DeviceName | Select-Object -Property $displayColumns | Format-Table -AutoSize

            # Optional exports
            if ($exportCsv) {
                try {
                    $sortedRows = $reportRows | Sort-Object AppName, DeviceName, DeviceId, Source
                    $exportRows = if ($Resume) { @($sortedRows | Where-Object { -not $existingRowKeys.Contains((& $buildRowKey $_)) }) } else { @($sortedRows) }

                    if ($exportRows.Count -eq 0) {
                        Write-NCMessage "No new CSV rows needed to be written to $csvPathResolved." -Level INFO
                    }
                    else {
                        $chunkSize = if ($BatchSize -gt 0) { $BatchSize } else { $exportRows.Count }
                        for ($offset = 0; $offset -lt $exportRows.Count; $offset += $chunkSize) {
                            $chunk = $exportRows[$offset..([Math]::Min($offset + $chunkSize - 1, $exportRows.Count - 1))]
                            if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                                $chunk | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding UTF8 -Delimiter $NCVars.CSV_DefaultLimiter -Append
                            }
                            else {
                                $chunk | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding UTF8 -Delimiter $NCVars.CSV_DefaultLimiter
                            }
                        }
                        if ($aborted) {
                            Write-NCMessage "Intune app inventory CSV export stopped early. Partial data kept at $csvPathResolved." -Level ERROR
                        }
                        else {
                            Write-NCMessage "CSV exported to $csvPathResolved" -Level SUCCESS
                        }
                    }
                }
                catch {
                    Write-NCMessage "Failed to export CSV: $($_.Exception.Message)" -Level WARNING
                }
            }

            if ($OutputJsonPath) {
                try {
                    $reportRows | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputJsonPath -Encoding UTF8
                    Write-NCMessage "JSON exported to $OutputJsonPath" -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Failed to export JSON: $($_.Exception.Message)" -Level WARNING
                }
            }

            # Optional per-app pivot/summary
            if ($PivotSummary) {
                Write-NCMessage "`n=== SUMMARY BY APPLICATION ===" -Level INFO
                $reportRows | Group-Object AppName | Sort-Object Count -Descending | ForEach-Object {
                    $app = $_.Name
                    $count = $_.Count
                    $versions = ($_.Group | Where-Object Version | Group-Object Version | Sort-Object Count -Descending | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Count }) -join ", "
                    $publishers = ($_.Group | Where-Object Publisher | Group-Object Publisher | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Count }) -join ", "
                    "- {0}: {1} devices`n    Versions: {2}`n    Top publishers: {3}" -f $app, $count, ($(if ($versions) { $versions } else { 'n/a' })), ($(if ($publishers) { $publishers } else { 'n/a' }))
                }
            }

            Write-NCMessage "App inventory report completed successfully." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Script execution failed: $($_.Exception.Message)" -Level ERROR
            exit 1
        }
    }
}

function Get-IntuneAppPresence {
    <#
    .SYNOPSIS
        Checks whether a single Intune-managed device has a matching app installed.
    .DESCRIPTION
        Queries one managed device and returns a single summary row with the match result.
        Use this when you want a quick yes/no check without the full inventory report.
    .PARAMETER DeviceName
        Intune managed device name to inspect.
    .PARAMETER ApplicationName
        Application name or wildcard pattern to match.
    .PARAMETER MinimumVersion
        Minimum application version to consider a match.
    .EXAMPLE
        Get-IntuneAppPresence -DeviceName "UE5CG30740PT" -ApplicationName "*java*"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DeviceName,

        [Parameter(Mandatory = $true)]
        [Alias('SearchText', 'Name', 'DisplayName', 'Query', 'AppName')]
        [string]$ApplicationName,

        [string]$MinimumVersion
    )

    try {
        $graphConnected = Test-MgGraphConnection -Scopes @('DeviceManagementManagedDevices.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
            return
        }

        $escapedDeviceName = $DeviceName.Replace("'", "''")
        $devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$escapedDeviceName'&`$select=id,deviceName,operatingSystem,userPrincipalName,lastSyncDateTime"
        $devices = @(Invoke-MgGraphRequest -Uri $devicesUri -Method GET -ErrorAction Stop).value
        $device = $devices | Select-Object -First 1

        if (-not $device) {
            return [pscustomobject]@{
                DeviceName    = $DeviceName
                DeviceId      = $null
                AppName       = $ApplicationName
                Present       = $false
                Version       = $null
                Publisher     = $null
                LastInventory = $null
                Source        = 'ManagedDevice'
                Status        = 'DeviceNotFound'
            }
        }

        $deviceAppsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($device.id)?`$expand=detectedApps"
        $deviceWithApps = Invoke-MgGraphRequest -Uri $deviceAppsUri -Method GET -ErrorAction Stop
        $matches = @($deviceWithApps.detectedApps | Where-Object { $_.displayName -like $ApplicationName })

        if ($MinimumVersion) {
            $matches = @($matches | Where-Object {
                $_.version -and (Test-NCIntuneVersionAtLeast -CurrentVersion $_.version -MinimumVersion $MinimumVersion)
            })
        }

        $match = $matches | Select-Object -First 1
        $present = [bool]$match

        $lastInventoryValue = Format-NCDateTime -Value $device.lastSyncDateTime -AsLocalTime

        return [pscustomobject]@{
            DeviceName    = $device.deviceName
            DeviceId      = $device.id
            AppName       = if ($match) { $match.displayName } else { $ApplicationName }
            Present       = $present
            Version       = if ($match) { $match.version } else { $null }
            Publisher     = if ($match) { $match.publisher } else { $null }
            LastInventory = $lastInventoryValue
            Source        = 'DetectedApps'
            Status        = if ($present) { 'Found' } else { 'NotFound' }
        }
    }
    catch {
        Write-NCMessage "Get-IntuneAppPresence failed: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function New-IntuneAppBasedGroup {
    <#
    .SYNOPSIS
        Creates Entra groups based on apps installed on Intune-managed devices.
    .DESCRIPTION
        Queries Intune-managed devices, discovers matching apps through detected apps and deployed
        app status data, and creates or updates Entra security groups populated with the matching
        Entra device objects. Use this for dynamic device targeting based on installed software.
    .PARAMETER ApplicationName
        Application name or wildcard pattern to match.
    .PARAMETER GroupName
        Explicit full group name to use instead of generating one from prefix and suffix.
        When supplied, all matching devices are collected into a single group target.
    .PARAMETER GroupPrefix
        Prefix applied to generated group names.
    .PARAMETER GroupSuffix
        Suffix applied to generated group names.
    .PARAMETER UpdateExisting
        Update matching groups instead of skipping them when they already exist.
    .PARAMETER MinimumVersion
        Minimum application version to keep in the result set.
    .PARAMETER FilterByType
        Optional app type filter for deployment-status coverage.
    .PARAMETER FilterByPlatform
        Optional device platform filter.
    .PARAMETER OnlySuccessfulInstalls
        When deployment data is used, keep only successful installs.
    .PARAMETER DryRun
        Preview changes without creating or updating groups.
    .PARAMETER MaxDevices
        Maximum number of devices to process. Use 0 for all devices.
    .EXAMPLE
        New-IntuneAppBasedGroup -ApplicationName "TeamViewer"
    .EXAMPLE
        New-IntuneAppBasedGroup -ApplicationName "TeamViewer" -GroupName "Devices - TeamViewer"
    .EXAMPLE
        New-IntuneAppBasedGroup -ApplicationName "Microsoft*" -GroupPrefix "SW-" -GroupSuffix "-Installed"
    .EXAMPLE
        New-IntuneAppBasedGroup -ApplicationName "Chrome" -MinimumVersion "120.0" -UpdateExisting
    .EXAMPLE
        New-IntuneAppBasedGroup -ApplicationName "*" -FilterByType Win32 -DryRun
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('SearchText', 'Name', 'DisplayName', 'Query', 'AppName')]
        [string]$ApplicationName,

        [string]$GroupName,
        [string]$GroupPrefix = 'Devices-With-',
        [string]$GroupSuffix = '',
        [switch]$UpdateExisting,
        [string]$MinimumVersion,

        [ValidateSet('Win32', 'Store', 'LOB', 'Web', 'iOS', 'Android', 'macOS', 'All')]
        [string]$FilterByType = 'All',

        [ValidateSet('Windows', 'iOS', 'Android', 'macOS', 'All')]
        [string]$FilterByPlatform = 'All',

        [switch]$OnlySuccessfulInstalls,
        [switch]$DryRun,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$MaxDevices = 0
    )

    begin {
        $graphConnected = $null
    }

    process {
        try {
            if ($null -eq $graphConnected) {
                $graphConnected = Test-MgGraphConnection -Scopes @(
                    'DeviceManagementManagedDevices.Read.All',
                    'DeviceManagementApps.Read.All',
                    'Group.ReadWrite.All',
                    'Directory.Read.All'
                ) -EnsureExchangeOnline:$false
                if (-not $graphConnected) {
                    Add-EmptyLine
                    Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                    return
                }

                if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
                    Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
                    return
                }
            }

            if ([string]::IsNullOrWhiteSpace($ApplicationName)) {
                Write-NCMessage "ApplicationName cannot be empty." -Level WARNING
                return
            }

            $useExplicitGroupName = -not [string]::IsNullOrWhiteSpace($GroupName)
            if ($useExplicitGroupName -and ($PSBoundParameters.ContainsKey('GroupPrefix') -or $PSBoundParameters.ContainsKey('GroupSuffix'))) {
                Write-Verbose 'GroupName was supplied; GroupPrefix and GroupSuffix will be ignored.'
            }

            Write-NCMessage "Starting app-based group creation process ..." -Level INFO

            Write-NCMessage "Retrieving managed devices ..." -Level INFO
            $devicesUri = 'https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,userPrincipalName,azureADDeviceId,azureActiveDirectoryDeviceId'
            if ($MaxDevices -gt 0) {
                $devicesUri += "&`$top=$MaxDevices"
            }

            $devices = @(Invoke-NCGraphAllPagesCore -Uri $devicesUri)
            if ($FilterByPlatform -ne 'All') {
                $devices = @($devices | Where-Object { $_.operatingSystem -like "$FilterByPlatform*" })
            }

            $managedDeviceLabel = if ($devices.Count -eq 1) { 'managed device' } else { 'managed devices' }
            Write-NCMessage "Found $($devices.Count) $managedDeviceLabel" -Level INFO

            $appDeviceMap = @{}
            $processedDevices = 0
            $batchNoticeWritten = $false

            Write-NCMessage "Scanning device applications ..." -Level INFO
            if ($devices.Count -gt 0) {
                $batchNoticeWritten = Write-NCGraphBatchNotice -Count ($devices.Count) -Noun 'device(s)' -PassThru
            }

            for ($offset = 0; $offset -lt $devices.Count; $offset += 20) {
                $deviceChunk = @($devices[$offset..([Math]::Min($offset + 20, $devices.Count) - 1)])

                $Percentage = Get-NCProgressPercent -Current $processedDevices -Total $devices.Count
                Write-Progress -Activity 'Processing Devices' -Status "$($deviceChunk[0].deviceName) - $processedDevices of $($devices.Count) devices - $Percentage%" -PercentComplete $Percentage

                $appRequests = @(for ($i = 0; $i -lt $deviceChunk.Count; $i++) {
                        @{ Id = "d$i"; Method = 'GET'; Url = "/deviceManagement/managedDevices/$([uri]::EscapeDataString([string]$deviceChunk[$i].id))?`$select=id,deviceName,operatingSystem,userPrincipalName,azureADDeviceId,azureActiveDirectoryDeviceId&`$expand=detectedApps" }
                    })
                $appResponses = @(Invoke-NCGraphBatch -Requests $appRequests -ApiVersion 'beta' -Activity 'Reading detected apps')

                for ($i = 0; $i -lt $deviceChunk.Count; $i++) {
                    $device = $deviceChunk[$i]
                    $processedDevices++
                    $Percentage = Get-NCProgressPercent -Current $processedDevices -Total $devices.Count

                    if (-not $appResponses[$i].Success) {
                        Write-NCMessage "Error processing device $($device.deviceName): $($appResponses[$i].ErrorMessage)" -Level WARNING
                        continue
                    }

                    $deviceWithApps = $appResponses[$i].Body
                    if ($deviceWithApps.detectedApps) {
                        foreach ($app in $deviceWithApps.detectedApps) {
                            if ($app.displayName -like $ApplicationName) {
                                Write-Progress -Activity 'Processing Devices' -Status "$($device.deviceName) / $($app.displayName) - $processedDevices of $($devices.Count) devices - $Percentage%" -PercentComplete $Percentage

                                if ($MinimumVersion -and $app.version) {
                                    if (-not (Test-NCIntuneVersionAtLeast -CurrentVersion $app.version -MinimumVersion $MinimumVersion)) {
                                        continue
                                    }
                                }

                                $appKey = $app.displayName
                                if (-not $appDeviceMap.ContainsKey($appKey)) {
                                    $appDeviceMap[$appKey] = @{
                                        Devices    = @()
                                        Versions   = @{}
                                        Publishers = @{}
                                        DeviceIds  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                                    }
                                }

                                if ($appDeviceMap[$appKey].DeviceIds.Add([string]$device.id)) {
                                    $appDeviceMap[$appKey].Devices += [ordered]@{
                                        DeviceId   = $device.id
                                        DeviceName = $device.deviceName
                                        Platform   = $device.operatingSystem
                                        User       = $device.userPrincipalName
                                        Version    = $app.version
                                        Publisher  = $app.publisher
                                        Source     = 'DetectedApps'
                                    }
                                }

                                if ($app.version) {
                                    if ($appDeviceMap[$appKey].Versions.ContainsKey($app.version)) {
                                        $appDeviceMap[$appKey].Versions[$app.version]++
                                    }
                                    else {
                                        $appDeviceMap[$appKey].Versions[$app.version] = 1
                                    }
                                }

                                if ($app.publisher) {
                                    if ($appDeviceMap[$appKey].Publishers.ContainsKey($app.publisher)) {
                                        $appDeviceMap[$appKey].Publishers[$app.publisher]++
                                    }
                                    else {
                                        $appDeviceMap[$appKey].Publishers[$app.publisher] = 1
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Write-Progress -Activity 'Processing Devices' -Completed

            if ($FilterByType -ne 'All' -or $OnlySuccessfulInstalls.IsPresent) {
                Write-NCMessage "Retrieving deployed application data ..." -Level INFO
                $appsUri = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
                $deployedApps = @(Invoke-NCGraphAllPagesCore -Uri $appsUri)

                $deployedCandidates = @(foreach ($app in $deployedApps) {
                        if ($app.displayName -like $ApplicationName) {
                            $candidateType = Get-NCIntuneAppTypeFromODataType -ODataType ([string]$app.'@odata.type')
                            if ($FilterByType -ne 'All' -and $candidateType -ne $FilterByType) {
                                continue
                            }
                            [pscustomobject]@{ App = $app; AppType = $candidateType }
                        }
                    })

                if ($deployedCandidates.Count -gt 0 -and -not $batchNoticeWritten) {
                    $batchNoticeWritten = Write-NCGraphBatchNotice -Count ($deployedCandidates.Count) -Noun 'app(s)' -PassThru
                }

                for ($offset = 0; $offset -lt $deployedCandidates.Count; $offset += 20) {
                    $candidateChunk = @($deployedCandidates[$offset..([Math]::Min($offset + 20, $deployedCandidates.Count) - 1)])
                    $statusRequests = @(for ($i = 0; $i -lt $candidateChunk.Count; $i++) {
                            @{ Id = "s$i"; Method = 'GET'; Url = "/deviceAppManagement/mobileApps/$([uri]::EscapeDataString([string]$candidateChunk[$i].App.id))/deviceStatuses" }
                        })
                    $statusResponses = @(Invoke-NCGraphBatchCollection -Requests $statusRequests -ApiVersion 'beta' -Activity 'Reading app deployment status')

                    for ($i = 0; $i -lt $candidateChunk.Count; $i++) {
                        if (-not $statusResponses[$i].Success) {
                            throw [System.InvalidOperationException]::new([string]$statusResponses[$i].ErrorMessage)
                        }

                        $app = $candidateChunk[$i].App
                        $appType = $candidateChunk[$i].AppType
                        $deviceStatuses = @($statusResponses[$i].Items)

                        foreach ($status in $deviceStatuses) {
                            if ($OnlySuccessfulInstalls.IsPresent -and $status.installState -ne 'installed') {
                                continue
                            }

                            $matchingDevice = $devices | Where-Object { $_.id -eq $status.deviceId } | Select-Object -First 1
                            if ($matchingDevice) {
                                $appKey = $app.displayName
                                if (-not $appDeviceMap.ContainsKey($appKey)) {
                                    $appDeviceMap[$appKey] = @{
                                        Devices    = @()
                                        Versions   = @{}
                                        Publishers = @{}
                                        AppType    = $appType
                                        DeviceIds  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                                    }
                                }

                                if ($appDeviceMap[$appKey].DeviceIds.Add([string]$matchingDevice.id)) {
                                    $appDeviceMap[$appKey].Devices += [ordered]@{
                                        DeviceId     = $matchingDevice.id
                                        DeviceName   = $matchingDevice.deviceName
                                        Platform     = $matchingDevice.operatingSystem
                                        User         = $matchingDevice.userPrincipalName
                                        InstallState = $status.installState
                                        AppType      = $appType
                                        Source       = 'DeploymentStatus'
                                    }
                                }
                            }
                        }
                    }
                }
            }
            elseif ($FilterByType -ne 'All') {
                Write-Verbose 'FilterByType applies only when deployment data is used. Skipping type filter for detected apps only.'
            }

            Write-NCMessage "Preparing groups for $($appDeviceMap.Count) applications ..." -Level INFO
            $groupsCreated = 0
            $groupsUpdated = 0
            $totalDevicesProcessed = 0
            $groupTargets = @()
            if ($useExplicitGroupName) {
                $allDevices = @($appDeviceMap.Values | ForEach-Object { $_.Devices } | Where-Object { $_ })
                $allVersions = @{}
                $allPublishers = @{}
                foreach ($appInfo in $appDeviceMap.Values) {
                    foreach ($version in $appInfo.Versions.Keys) {
                        if (-not $allVersions.ContainsKey($version)) {
                            $allVersions[$version] = 0
                        }
                        $allVersions[$version] += $appInfo.Versions[$version]
                    }
                    foreach ($publisher in $appInfo.Publishers.Keys) {
                        if (-not $allPublishers.ContainsKey($publisher)) {
                            $allPublishers[$publisher] = 0
                        }
                        $allPublishers[$publisher] += $appInfo.Publishers[$publisher]
                    }
                }

                $groupTargets += [ordered]@{
                    AppName    = if ($appDeviceMap.Keys.Count -gt 1) { 'Matching apps' } else { $appDeviceMap.Keys | Select-Object -First 1 }
                    Devices    = $allDevices
                    Versions   = $allVersions
                    Publishers = $allPublishers
                    GroupName  = $GroupName
                    GroupScope = 'ExplicitGroupName'
                }
            }
            else {
                foreach ($appName in $appDeviceMap.Keys) {
                    $groupTargets += [ordered]@{
                        AppName    = $appName
                        Devices    = $appDeviceMap[$appName].Devices
                        Versions   = $appDeviceMap[$appName].Versions
                        Publishers = $appDeviceMap[$appName].Publishers
                        GroupName  = $null
                        GroupScope = 'PerApp'
                    }
                }
            }

            # Device and group names per target (targets without devices are skipped)
            $targetInfos = [System.Collections.Generic.List[object]]::new()
            foreach ($target in $groupTargets) {
                $uniqueDeviceIds = @(
                    $target.Devices |
                        ForEach-Object { Get-NCCoreProperty -Object $_ -Names @('DeviceId', 'deviceId', 'Id', 'id') } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                        Select-Object -Unique
                )

                if ($uniqueDeviceIds.Count -eq 0) {
                    $targetInfos.Add($null)
                    continue
                }

                if ($useExplicitGroupName) {
                    $targetGroupName = Get-NCIntuneAppBasedGroupName -GroupName $GroupName
                    $targetGroupDescription = 'Devices matching selected Intune apps (Created via Nebula.Core)'
                }
                else {
                    $targetGroupName = Get-NCIntuneAppBasedGroupName -AppName $target.AppName -GroupPrefix $GroupPrefix -GroupSuffix $GroupSuffix
                    $targetGroupDescription = "Devices with $($target.AppName) installed (Created via Nebula.Core)"
                }

                $targetInfos.Add([pscustomobject]@{
                        UniqueDeviceIds  = $uniqueDeviceIds
                        GroupName        = $targetGroupName
                        GroupDescription = $targetGroupDescription
                    })
            }

            # Existing groups are looked up for all targets in Graph batches
            $groupLookup = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
            # Failed lookups (name -> error): the group may exist, so the target is skipped rather than created
            $groupLookupFailed = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            if (-not $DryRun.IsPresent) {
                $lookupNames = @($targetInfos | Where-Object { $_ } | ForEach-Object { $_.GroupName } | Select-Object -Unique)
                if ($lookupNames.Count -gt 0 -and -not $batchNoticeWritten) {
                    $batchNoticeWritten = Write-NCGraphBatchNotice -Count ($lookupNames.Count) -Noun 'group(s)' -PassThru
                }

                if ($lookupNames.Count -gt 0) {
                    $groupRequests = @(for ($i = 0; $i -lt $lookupNames.Count; $i++) {
                            $groupFilter = "displayName eq '$($lookupNames[$i].Replace("'", "''"))'"
                            @{ Id = "g$i"; Method = 'GET'; Url = "/groups?`$filter=$([uri]::EscapeDataString($groupFilter))&`$select=id,displayName" }
                        })
                    $groupResponses = @(Invoke-NCGraphBatchCollection -Requests $groupRequests -ApiVersion 'v1.0' -Activity 'Looking up existing groups')
                    for ($i = 0; $i -lt $lookupNames.Count; $i++) {
                        if ($groupResponses[$i].Success) {
                            $groupLookup[$lookupNames[$i]] = @($groupResponses[$i].Items) | Select-Object -First 1
                        }
                        else {
                            $groupLookupFailed[$lookupNames[$i]] = [string]$groupResponses[$i].ErrorMessage
                        }
                    }
                }
            }

            # Membership writes (one POST/DELETE per device inside the batch, per-device outcome)
            $memberStats = @{ Added = 0; Removed = 0 }
            $addGroupMembers = {
                param([string]$GroupId, [string]$GroupLabel, [object[]]$Members, [string]$FailureFormat)

                $addRequests = @(for ($i = 0; $i -lt $Members.Count; $i++) {
                        @{ Id = "m$i"; Method = 'POST'; Url = "/groups/$([uri]::EscapeDataString($GroupId))/members/`$ref"; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id ([string]$Members[$i].EntraDeviceId)) } }
                    })
                $addResponses = @(Invoke-NCGraphBatch -Requests $addRequests -Activity "Adding devices to $GroupLabel")

                for ($i = 0; $i -lt $Members.Count; $i++) {
                    $memberLabel = $Members[$i].DeviceName
                    if ($addResponses[$i].Success) {
                        $memberStats.Added++
                        Write-Verbose "Added device '$memberLabel' to group '$GroupLabel'"
                    }
                    elseif ($addResponses[$i].ErrorMessage -match 'added object references already exist') {
                        Write-NCMessage "Device '$memberLabel' is already a member of '$GroupLabel'" -Level WARNING
                    }
                    else {
                        Write-NCMessage ($FailureFormat -f $memberLabel, $addResponses[$i].ErrorMessage) -Level ERROR
                    }
                }
            }
            $removeGroupMembers = {
                param([string]$GroupId, [string]$GroupLabel, [object[]]$MemberIds, [hashtable]$MemberNames)

                $removeRequests = @(for ($i = 0; $i -lt $MemberIds.Count; $i++) {
                        @{ Id = "r$i"; Method = 'DELETE'; Url = "/groups/$([uri]::EscapeDataString($GroupId))/members/$([uri]::EscapeDataString([string]$MemberIds[$i]))/`$ref" }
                    })
                $removeResponses = @(Invoke-NCGraphBatch -Requests $removeRequests -Activity "Removing members from $GroupLabel")

                for ($i = 0; $i -lt $MemberIds.Count; $i++) {
                    if ($removeResponses[$i].Success) {
                        $memberStats.Removed++
                    }
                    elseif ($removeResponses[$i].Status -eq 404) {
                        Write-Verbose "Member '$($MemberIds[$i])' is no longer a member of '$GroupLabel'"
                    }
                    else {
                        $removeLabel = if ($MemberNames.ContainsKey([string]$MemberIds[$i])) { $MemberNames[[string]$MemberIds[$i]] } else { [string]$MemberIds[$i] }
                        Write-NCMessage "Failed to remove device '$removeLabel' from '$GroupLabel': $($removeResponses[$i].ErrorMessage)" -Level ERROR
                    }
                }
            }

            $processedApps = 0
            for ($targetIndex = 0; $targetIndex -lt $groupTargets.Count; $targetIndex++) {
                $target = $groupTargets[$targetIndex]
                $processedApps++
                $appName = $target.AppName
                $appInfo = $target
                $targetInfo = $targetInfos[$targetIndex]

                if (-not $targetInfo) {
                    continue
                }

                $uniqueDeviceIds = $targetInfo.UniqueDeviceIds
                $deviceCount = $uniqueDeviceIds.Count
                $groupName = $targetInfo.GroupName
                $groupDescription = $targetInfo.GroupDescription

                $Percentage = ($processedApps / [Math]::Max($groupTargets.Count, 1)) * 100
                Write-Progress -Activity 'Processing App Groups' -Status "$appName / $groupName - $processedApps of $($groupTargets.Count) groups - $Percentage%" -PercentComplete $Percentage

                if ($DryRun.IsPresent) {
                    Write-NCMessage "[DRY RUN] Would create/update group: $groupName" -Level INFO
                    Write-NCMessage "Total devices with app matches: $deviceCount" -Level INFO
                    Write-Verbose "Devices to be added:"
                    foreach ($device in $appInfo.Devices) {
                        Write-Verbose "  - $($device.DeviceName) ($($device.Platform))"
                    }

                    if ($appInfo.Versions.Count -gt 0) {
                        Write-NCMessage "Versions found: $($appInfo.Versions.Keys -join ', ')" -Level INFO
                    }

                    $totalDevicesProcessed += $deviceCount
                    continue
                }

                $existingGroup = $null
                if ($groupLookupFailed.ContainsKey($groupName)) {
                    Write-NCMessage "Unable to look up existing group '$groupName', skipping it to avoid creating a duplicate: $($groupLookupFailed[$groupName])" -Level ERROR
                    continue
                }
                elseif ($groupLookup.ContainsKey($groupName)) {
                    $existingGroup = $groupLookup[$groupName]
                }

                if ($existingGroup -and -not $UpdateExisting.IsPresent) {
                    Write-NCMessage "Group '$groupName' already exists. Use -UpdateExisting to update it." -Level WARNING
                    continue
                }

                $entraDevices = @()
                # Removals are only safe when every device was resolved without a Graph error
                $resolutionIncomplete = $false
                $seenEntraDeviceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                try {
                    Write-Progress -Activity 'Resolving Entra Devices' -Status "$groupName - $deviceCount devices" -PercentComplete 0
                    foreach ($resolved in @(Resolve-NCIntuneManagedDeviceEntraMembers -ManagedDevices $devices -DeviceIds @($uniqueDeviceIds))) {
                        if ($resolved.LookupFailed) {
                            $resolutionIncomplete = $true
                        }
                        # Two Intune devices can map to the same Entra device: add it once
                        if ($resolved.Resolution -and $seenEntraDeviceIds.Add([string]$resolved.Resolution.EntraDeviceId)) {
                            $entraDevices += @{
                                IntuneDeviceId = $resolved.Resolution.IntuneDeviceId
                                EntraDeviceId  = $resolved.Resolution.EntraDeviceId
                                DeviceName     = $resolved.Resolution.DeviceName
                            }
                        }
                    }
                }
                catch {
                    Write-NCMessage "Error looking up Entra ID devices for group $($groupName): $($_.Exception.Message)" -Level ERROR
                    $resolutionIncomplete = $true
                }

                Write-Progress -Activity 'Resolving Entra Devices' -Completed

                if ($existingGroup -and $UpdateExisting.IsPresent) {
                    if ($PSCmdlet.ShouldProcess($groupName, 'Update group members')) {
                        try {
                            $currentMembersUri = "https://graph.microsoft.com/v1.0/groups/$($existingGroup.id)/members"
                            $currentMembers = @(Invoke-NCGraphAllPagesCore -Uri $currentMembersUri)
                            $currentMemberIds = $currentMembers | ForEach-Object { $_.id }
                            $currentMemberNames = @{}
                            foreach ($currentMember in $currentMembers) {
                                if ($currentMember.id -and $currentMember.displayName) { $currentMemberNames[[string]$currentMember.id] = [string]$currentMember.displayName }
                            }

                            $entraDeviceIds = $entraDevices | ForEach-Object { $_.EntraDeviceId }
                            $deviceIdsToAdd = @($entraDeviceIds | Where-Object { $_ -notin $currentMemberIds })
                            $deviceIdsToRemove = @($currentMemberIds | Where-Object { $_ -notin $entraDeviceIds })
                            if ($resolutionIncomplete -and $deviceIdsToRemove.Count -gt 0) {
                                Write-NCMessage "Entra device resolution for group '$groupName' is incomplete. No members will be removed from it in this run." -Level WARNING
                                $deviceIdsToRemove = @()
                            }
                            $devicesToAdd = @($entraDevices | Where-Object { $_.EntraDeviceId -in $deviceIdsToAdd })

                            $memberStats.Added = 0
                            $memberStats.Removed = 0
                            if ($devicesToAdd.Count -gt 0) {
                                & $addGroupMembers -GroupId ([string]$existingGroup.id) -GroupLabel $groupName -Members $devicesToAdd -FailureFormat "Failed to add device '{0}' to '$($groupName.Replace('{', '{{').Replace('}', '}}'))': {1}"
                            }

                            if ($deviceIdsToRemove.Count -gt 0) {
                                & $removeGroupMembers -GroupId ([string]$existingGroup.id) -GroupLabel $groupName -MemberIds $deviceIdsToRemove -MemberNames $currentMemberNames
                            }

                            Write-NCMessage "Updated group: $groupName (Added: $($memberStats.Added), Removed: $($memberStats.Removed))" -Level SUCCESS
                            if ($deviceIdsToAdd.Count -gt 0) {
                                Write-Verbose "Added devices:"
                                foreach ($deviceId in $deviceIdsToAdd) {
                                    $deviceInfo = $entraDevices | Where-Object { $_.EntraDeviceId -eq $deviceId } | Select-Object -First 1
                                    if ($deviceInfo) {
                                        Write-Verbose "  - $($deviceInfo.DeviceName)"
                                    }
                                }
                            }

                            $groupsUpdated++
                        }
                        catch {
                            Write-NCMessage "Failed to update group: $($_.Exception.Message)" -Level ERROR
                        }
                    }
                }
                else {
                    if ($PSCmdlet.ShouldProcess($groupName, 'Create new group')) {
                        try {
                            $groupBody = @{
                                displayName     = $groupName
                                mailEnabled     = $false
                                mailNickname    = ($groupName -replace '[^a-zA-Z0-9]', '')
                                securityEnabled = $true
                                description     = $groupDescription
                            } | ConvertTo-Json -Depth 10

                            $newGroup = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/groups' -Method POST -Body $groupBody -ContentType 'application/json'
                            Write-NCMessage "Created group: $groupName $($newGroup.id)" -Level SUCCESS
                            $groupLookup[$groupName] = [pscustomobject]@{ id = $newGroup.id; displayName = $groupName }
                            $null = $groupLookupFailed.Remove($groupName)

                            if ($entraDevices.Count -gt 0) {
                                try {
                                    $memberStats.Added = 0
                                    & $addGroupMembers -GroupId ([string]$newGroup.id) -GroupLabel $groupName -Members $entraDevices -FailureFormat "Group created but failed to add device '{0}': {1}"

                                    if ($memberStats.Added -gt 0) {
                                        $groupDeviceLabel = if ($memberStats.Added -eq 1) { 'device' } else { 'devices' }
                                        Write-NCMessage "Added $($memberStats.Added) $groupDeviceLabel to group" -Level SUCCESS
                                        Write-Verbose "Added devices:"
                                        foreach ($device in $entraDevices) {
                                            Write-Verbose "  - $($device.DeviceName)"
                                        }
                                    }
                                }
                                catch {
                                    Write-NCMessage "Group created but failed to add members: $($_.Exception.Message)" -Level ERROR
                                }
                            }

                            $groupsCreated++
                        }
                        catch {
                            Write-NCMessage "Failed to create group: $($_.Exception.Message)" -Level ERROR
                            Write-Verbose "Group body: $groupBody"
                        }
                    }
                }

                $totalDevicesProcessed += $deviceCount
            }

            Write-Progress -Activity 'Processing App Groups' -Completed

            # Add-EmptyLine
            # Write-NCMessage " - Applications matched: $($appDeviceMap.Count)" -Level INFO
            # Write-NCMessage " - Total devices processed: $totalDevicesProcessed" -Level INFO
            # Write-NCMessage " - Groups created: $groupsCreated" -Level INFO
            # Write-NCMessage " - Groups updated: $groupsUpdated" -Level INFO

            if ($DryRun.IsPresent) {
                Write-NCMessage "[DRY RUN] No changes were made" -Level INFO
            }

            if ($appDeviceMap.Count -gt 0) {
                Add-EmptyLine
                Write-NCMessage "Top Applications by Device Count:" -Level INFO
                $appDeviceMap.GetEnumerator() |
                    Sort-Object { if ($_.Value.DeviceIds) { $_.Value.DeviceIds.Count } else { $_.Value.Devices.Count } } -Descending |
                    Select-Object -First 10 |
                    ForEach-Object {
                        $deviceCount = if ($_.Value.DeviceIds) { $_.Value.DeviceIds.Count } else { ($_.Value.Devices | Select-Object -Property DeviceId -Unique).Count }
                        $deviceLabel = if ($deviceCount -eq 1) { 'device' } else { 'devices' }
                        Write-NCMessage "  - $($_.Key): $deviceCount $deviceLabel" -Level INFO
                    }
            }

            Add-EmptyLine
            Write-NCMessage "App-based group creation completed successfully." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Script execution failed: $($_.Exception.Message)" -Level ERROR
            return
        }
    }
}


