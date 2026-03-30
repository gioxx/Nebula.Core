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
        Write-NCMessage "Intune profiles found for '$normalizedSearch': $($results.Count)" -Level VERBOSE

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
    .PARAMETER PivotSummary
        Print a per-app summary after the report is built.
    .EXAMPLE
        Export-IntuneAppInventory -ApplicationName "TeamViewer"
    .EXAMPLE
        Export-IntuneAppInventory -ApplicationName "Microsoft*" -IncludeDeployedApps -FilterByType Win32 -OutputCsvPath "apps.csv"
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

        [ValidateRange(0, [int]::MaxValue)]
        [int]$MaxDevices = 0,

        [string]$OutputCsvPath,
        [string]$OutputJsonPath,
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

            Write-NCMessage "Starting app inventory reporting ..." -Level INFO

            # Pull devices
            $devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
            if ($MaxDevices -gt 0) {
                $devicesUri += "?`$top=$MaxDevices"
            }
            $devices = @(Invoke-NCGraphAllPagesCore -Uri $devicesUri)
            if ($FilterByPlatform -ne "All") {
                $devices = $devices | Where-Object { $_.operatingSystem -like "$FilterByPlatform*" }
            }
            Write-NCMessage "Devices retrieved: $($devices.Count)" -Level INFO

            # Build app --> device mapping from Detected Apps
            $appDeviceMap = @{}
            $processed = 0

            foreach ($device in $devices) {
                $processed++
                Write-Progress -Activity "Reading Detected Apps" -Status "$processed / $($devices.Count) devices" -CurrentOperation "Device: $($device.deviceName)" -PercentComplete (($processed / [Math]::Max($devices.Count,1)) * 100)
                try {
                    $deviceAppsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($device.id)?`$expand=detectedApps"
                    $deviceWithApps = Invoke-MgGraphRequest -Uri $deviceAppsUri -Method GET -ErrorAction Stop
                    
                    foreach ($app in ($deviceWithApps.detectedApps | Where-Object { $_.displayName -like $ApplicationName })) {
                        Write-Progress -Activity "Reading Detected Apps" -Status "$processed / $($devices.Count) devices" -CurrentOperation "Device: $($device.deviceName) | App: $($app.displayName)" -PercentComplete (($processed / [Math]::Max($devices.Count,1)) * 100)
                        if ($MinimumVersion -and $app.version) {
                            if (-not (Test-NCIntuneVersionAtLeast -CurrentVersion $app.version -MinimumVersion $MinimumVersion)) {
                                continue
                            }
                        }

                        $key = $app.displayName
                        if (-not $appDeviceMap.ContainsKey($key)) {
                            $appDeviceMap[$key] = [ordered]@{ Devices = @(); Versions = @{}; Publishers = @{} }
                        }
                        
                        $appDeviceMap[$key].Devices += [ordered]@{
                            DeviceId   = $device.id
                            DeviceName = $device.deviceName
                            Platform   = $device.operatingSystem
                            User       = $device.userPrincipalName
                            Version    = $app.version
                            # Publisher  = $app.publisher
                            Source     = "DetectedApps"
                        }
                        
                        if ($app.version) { 
                            $appDeviceMap[$key].Versions[$app.version] = ($appDeviceMap[$key].Versions[$app.version] + 1)
                        }
                        
                        if ($app.publisher) { 
                            $appDeviceMap[$key].Publishers[$app.publisher] = ($appDeviceMap[$key].Publishers[$app.publisher] + 1)
                        }
                    }
                    Start-Sleep -Milliseconds 40
                }
                catch {
                    if ($_.Exception.Message -like "*429*") {
                        Write-NCMessage "`nRate limit hit, waiting 60 seconds ..." -Level INFO
                        Start-Sleep -Seconds 60
                        $processed--
                        Write-Progress -Activity "Reading Detected Apps" -Status "$processed / $($devices.Count) devices" -CurrentOperation "Waiting after rate limit" -PercentComplete (($processed / [Math]::Max($devices.Count,1)) * 100)
                        continue
                    }
                    Write-NCMessage "Error reading apps for $($device.deviceName): $($_.Exception.Message)" -Level WARNING
                }
            }
            Write-Progress -Activity "Reading Detected Apps" -Completed

            # Optionally incorporate deployment statuses (broadens coverage)
            if ($IncludeDeployedApps) {
                Write-NCMessage "Including deployed apps device status ..." -Level INFO
                $appsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"
                $allApps = @(Invoke-NCGraphAllPagesCore -Uri $appsUri)
                
                foreach ($app in $allApps | Where-Object { $_.displayName -like $ApplicationName }) {
                    $appType = Get-NCIntuneAppTypeFromODataType -ODataType $app.'@odata.type'
                    if ($FilterByType -ne "All" -and $appType -ne $FilterByType) {
                        continue
                    }

                    $statusUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/deviceStatuses"
                    $statuses = @(Invoke-NCGraphAllPagesCore -Uri $statusUri)
                    
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
                            $appDeviceMap[$key].Devices += [ordered]@{
                                DeviceId     = $d.id
                                DeviceName   = $d.deviceName
                                Platform     = $d.operatingSystem
                                User         = $d.userPrincipalName
                                Version      = $null
                                # Publisher    = $null
                                # InstallState = $s.installState
                                # AppType      = $appType
                                Source       = "DeploymentStatus"
                            }
                        }
                    }
                }
            }

            # Optional app type filter when only Detected Apps were used
            if (-not $IncludeDeployedApps -and $FilterByType -ne "All") {
                Write-NCMessage "FilterByType applies only when -IncludeDeployedApps is used. Skipping type filter for Detected Apps only." -Level VERBOSE
            }

            # Build flat rows
            $rows = @()
            foreach ($appName in $appDeviceMap.Keys) {
                foreach ($dev in $appDeviceMap[$appName].Devices) {
                    $rows += [pscustomobject]@{
                        AppName      = $appName
                        Version      = $dev.Version
                        # Publisher    = $dev.Publisher
                        # AppType      = $dev.AppType
                        DeviceName   = $dev.DeviceName
                        DeviceId     = $dev.DeviceId
                        Platform     = $dev.Platform
                        User         = $dev.User
                        # InstallState = $dev.InstallState
                        Source       = $dev.Source
                    }
                }
            }

            if (-not $rows -or $rows.Count -eq 0) {
                Write-NCMessage "No matches found for '$ApplicationName' with the provided filters." -Level WARNING
                return
            }

            # Console table output
            $rows | Sort-Object AppName, DeviceName | Format-Table -AutoSize AppName, Version, Publisher, AppType, DeviceName, Platform, User, InstallState, Source

            # Optional exports
            if ($OutputCsvPath) {
                try {
                    $rows | Sort-Object AppName, DeviceName | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
                    Write-NCMessage "CSV exported to $OutputCsvPath" -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Failed to export CSV: $($_.Exception.Message)" -Level WARNING
                }
            }

            if ($OutputJsonPath) {
                try {
                    $rows | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputJsonPath -Encoding UTF8
                    Write-NCMessage "JSON exported to $OutputJsonPath" -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Failed to export JSON: $($_.Exception.Message)" -Level WARNING
                }
            }

            # Optional per-app pivot/summary
            if ($PivotSummary) {
                Write-NCMessage "`n=== SUMMARY BY APPLICATION ===" -Level INFO
                $rows | Group-Object AppName | Sort-Object Count -Descending | ForEach-Object {
                    $app = $_.Name
                    $count = $_.Count
                    $versions = ($_.Group | Where-Object Version | Group-Object Version | Sort-Object Count -Descending | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Count }) -join ", "
                    $publishers = ($_.Group | Where-Object Publisher | Group-Object Publisher | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Count }) -join ", "
                    "• {0}: {1} devices`n    Versions: {2}`n    Top publishers: {3}" -f $app, $count, ($(if ($versions) { $versions } else { 'n/a' })), ($(if ($publishers) { $publishers } else { 'n/a' }))
                }
            }
        }
        catch {
            Write-NCMessage "Script execution failed: $($_.Exception.Message)" -Level ERROR
            exit 1
        }
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
                Write-NCMessage 'GroupName was supplied; GroupPrefix and GroupSuffix will be ignored.' -Level VERBOSE
            }

            Write-NCMessage "Starting app-based group creation process ..." -Level INFO

            Write-NCMessage "Retrieving managed devices ..." -Level INFO
            $devicesUri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices'
            if ($MaxDevices -gt 0) {
                $devicesUri += "?`$top=$MaxDevices"
            }

            $devices = @(Invoke-NCGraphAllPagesCore -Uri $devicesUri)
            if ($FilterByPlatform -ne 'All') {
                $devices = @($devices | Where-Object { $_.operatingSystem -like "$FilterByPlatform*" })
            }

            Write-NCMessage "Found $($devices.Count) managed devices" -Level INFO

            $appDeviceMap = @{}
            $processedDevices = 0

            Write-NCMessage "Processing device applications ..." -Level INFO
            foreach ($device in $devices) {
                $processedDevices++
                Write-Progress -Activity 'Processing Devices' -Status "$processedDevices of $($devices.Count) devices" -CurrentOperation "Device: $($device.deviceName)" -PercentComplete (($processedDevices / [Math]::Max($devices.Count, 1)) * 100)

                try {
                    $deviceAppsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($device.id)?`$expand=detectedApps"
                    $deviceWithApps = Invoke-MgGraphRequest -Uri $deviceAppsUri -Method GET

                    if ($deviceWithApps.detectedApps) {
                        foreach ($app in $deviceWithApps.detectedApps) {
                            if ($app.displayName -like $ApplicationName) {
                                Write-Progress -Activity 'Processing Devices' -Status "$processedDevices of $($devices.Count) devices" -CurrentOperation "Device: $($device.deviceName) | App: $($app.displayName)" -PercentComplete (($processedDevices / [Math]::Max($devices.Count, 1)) * 100)

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
                                    }
                                }

                                $appDeviceMap[$appKey].Devices += [ordered]@{
                                    DeviceId   = $device.id
                                    DeviceName = $device.deviceName
                                    Platform   = $device.operatingSystem
                                    User       = $device.userPrincipalName
                                    Version    = $app.version
                                    Publisher  = $app.publisher
                                    Source     = 'DetectedApps'
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

                    Start-Sleep -Milliseconds 50
                }
                catch {
                    if ($_.Exception.Message -like '*429*') {
                        Write-NCMessage "Rate limit hit, waiting 60 seconds ..." -Level INFO
                        Start-Sleep -Seconds 60
                        $processedDevices--
                        continue
                    }

                    Write-NCMessage "Error processing device $($device.deviceName): $($_.Exception.Message)" -Level WARNING
                }
            }

            Write-Progress -Activity 'Processing Devices' -Completed

            if ($FilterByType -ne 'All' -or $OnlySuccessfulInstalls.IsPresent) {
                Write-NCMessage "Retrieving deployed application data ..." -Level INFO
                $appsUri = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
                $deployedApps = @(Invoke-NCGraphAllPagesCore -Uri $appsUri)

                foreach ($app in $deployedApps) {
                    if ($app.displayName -like $ApplicationName) {
                        $appType = Get-NCIntuneAppTypeFromODataType -ODataType ([string]$app.'@odata.type')
                        if ($FilterByType -ne 'All' -and $appType -ne $FilterByType) {
                            continue
                        }

                        $statusUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/deviceStatuses"
                        $deviceStatuses = @(Invoke-NCGraphAllPagesCore -Uri $statusUri)

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
                                    }
                                }

                                $existingDevice = $appDeviceMap[$appKey].Devices | Where-Object { $_.DeviceId -eq $matchingDevice.id }
                                if (-not $existingDevice) {
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
                Write-NCMessage 'FilterByType applies only when deployment data is used. Skipping type filter for detected apps only.' -Level VERBOSE
            }

            Write-NCMessage "Processing groups for $($appDeviceMap.Count) applications ..." -Level INFO
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

            $processedApps = 0
            foreach ($target in $groupTargets) {
                $processedApps++
                $appName = $target.AppName
                $appInfo = $target
                $uniqueDevices = $appInfo.Devices | Select-Object -Property DeviceId -Unique
                $deviceCount = $uniqueDevices.Count

                if ($deviceCount -eq 0) {
                    continue
                }

                if ($useExplicitGroupName) {
                    $groupName = Get-NCIntuneAppBasedGroupName -GroupName $GroupName
                    $groupDescription = 'Devices matching selected Intune apps (Created via Nebula.Core)'
                }
                else {
                    $groupName = Get-NCIntuneAppBasedGroupName -AppName $appName -GroupPrefix $GroupPrefix -GroupSuffix $GroupSuffix
                    $groupDescription = "Devices with $appName installed (Created via Nebula.Core)"
                }
                Write-Progress -Activity 'Processing App Groups' -Status "$processedApps of $($groupTargets.Count) groups" -CurrentOperation "App: $appName | Group: $groupName" -PercentComplete (($processedApps / [Math]::Max($groupTargets.Count, 1)) * 100)
                if ($DryRun.IsPresent) {
                    Write-NCMessage "[DRY RUN] Would create/update group: $groupName" -Level INFO
                    Write-NCMessage "Total devices with app matches: $deviceCount" -Level INFO
                    Write-NCMessage "Devices to be added:" -Level INFO
                    foreach ($device in $appInfo.Devices) {
                        Write-NCMessage "  • $($device.DeviceName) ($($device.Platform))" -Level INFO
                    }

                    if ($appInfo.Versions.Count -gt 0) {
                        Write-NCMessage "Versions found: $($appInfo.Versions.Keys -join ', ')" -Level INFO
                    }

                    $totalDevicesProcessed += $deviceCount
                    continue
                }

                $existingGroup = $null
                try {
                    $groupFilter = "displayName eq '$($groupName.Replace("'", "''"))'"
                    $existingGroup = Get-MgGroup -Filter $groupFilter -All -ErrorAction Stop | Select-Object -First 1
                }
                catch {
                    Write-NCMessage "No existing group found with name: $groupName" -Level WARNING
                }

                if ($existingGroup -and -not $UpdateExisting.IsPresent) {
                    Write-NCMessage "Group '$groupName' already exists. Use -UpdateExisting to update it." -Level WARNING
                    continue
                }

                $memberIds = @()
                $entraDevices = @()
                $processedMembers = 0

                foreach ($device in $uniqueDevices) {
                    $processedMembers++

                    try {
                        $intuneDevice = $devices | Where-Object { $_.id -eq $device.DeviceId } | Select-Object -First 1
                        $deviceLabel = if ($intuneDevice -and -not [string]::IsNullOrWhiteSpace($intuneDevice.deviceName)) {
                            $intuneDevice.deviceName
                        }
                        else {
                            $device.DeviceId
                        }

                        Write-Progress -Activity 'Resolving Entra Devices' -Status "$processedMembers of $($uniqueDevices.Count) devices" -CurrentOperation "Device: $deviceLabel" -PercentComplete (($processedMembers / [Math]::Max($uniqueDevices.Count, 1)) * 100)

                        if ($intuneDevice -and $intuneDevice.azureADDeviceId) {
                            $filter = "deviceId eq '$($intuneDevice.azureADDeviceId)'"
                            $entraDeviceUri = "https://graph.microsoft.com/v1.0/devices?`$filter=$filter"
                            $entraDeviceResponse = Invoke-MgGraphRequest -Uri $entraDeviceUri -Method GET

                            if ($entraDeviceResponse.value -and $entraDeviceResponse.value.Count -gt 0) {
                                $entraDevice = $entraDeviceResponse.value[0]
                                $memberIds += "https://graph.microsoft.com/v1.0/directoryObjects/$($entraDevice.id)"
                                $entraDevices += @{
                                    IntuneDeviceId = $device.DeviceId
                                    EntraDeviceId  = $entraDevice.id
                                    DeviceName     = $intuneDevice.deviceName
                                }
                                Write-NCMessage "Found Entra ID device: $($intuneDevice.deviceName) -> $($entraDevice.id)" -Level VERBOSE
                            }
                            else {
                                Write-NCMessage "Device not found in Entra ID: $($intuneDevice.deviceName) (Azure AD Device ID: $($intuneDevice.azureADDeviceId))" -Level WARNING
                            }
                        }
                        else {
                            Write-NCMessage "No Azure AD Device ID for: $($intuneDevice.deviceName)" -Level WARNING
                        }
                    }
                    catch {
                        Write-NCMessage "Error looking up Entra ID device for $($intuneDevice.deviceName): $($_.Exception.Message)" -Level ERROR
                    }
                }

                Write-Progress -Activity 'Resolving Entra Devices' -Completed

                if ($existingGroup -and $UpdateExisting.IsPresent) {
                    if ($PSCmdlet.ShouldProcess($groupName, 'Update group members')) {
                        try {
                            $currentMembersUri = "https://graph.microsoft.com/v1.0/groups/$($existingGroup.id)/members"
                            $currentMembers = @(Invoke-NCGraphAllPagesCore -Uri $currentMembersUri)
                            $currentMemberIds = $currentMembers | ForEach-Object { $_.id }

                            $entraDeviceIds = $entraDevices | ForEach-Object { $_.EntraDeviceId }
                            $deviceIdsToAdd = $entraDeviceIds | Where-Object { $_ -notin $currentMemberIds }
                            $deviceIdsToRemove = $currentMemberIds | Where-Object { $_ -notin $entraDeviceIds }

                            if ($deviceIdsToAdd.Count -gt 0) {
                                $batchSize = 20
                                for ($i = 0; $i -lt $deviceIdsToAdd.Count; $i += $batchSize) {
                                    $batch = $deviceIdsToAdd[$i..([Math]::Min($i + $batchSize - 1, $deviceIdsToAdd.Count - 1))]
                                    $addBody = @{
                                        'members@odata.bind' = $batch | ForEach-Object {
                                            "https://graph.microsoft.com/v1.0/directoryObjects/$_"
                                        }
                                    } | ConvertTo-Json -Depth 10

                                    Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($existingGroup.id)" -Method PATCH -Body $addBody -ContentType 'application/json'
                                    Write-Verbose "Added batch of $($batch.Count) members"
                                }
                            }

                            foreach ($memberId in $deviceIdsToRemove) {
                                $removeUri = "https://graph.microsoft.com/v1.0/groups/$($existingGroup.id)/members/$memberId/`$ref"
                                Invoke-MgGraphRequest -Uri $removeUri -Method DELETE
                            }

                            Write-NCMessage "Updated group: $groupName (Added: $($deviceIdsToAdd.Count), Removed: $($deviceIdsToRemove.Count))" -Level SUCCESS
                            if ($deviceIdsToAdd.Count -gt 0) {
                                Write-NCMessage "Added devices:" -Level INFO
                                foreach ($deviceId in $deviceIdsToAdd) {
                                    $deviceInfo = $entraDevices | Where-Object { $_.EntraDeviceId -eq $deviceId } | Select-Object -First 1
                                    if ($deviceInfo) {
                                        Write-NCMessage "  • $($deviceInfo.DeviceName)" -Level INFO
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
                            Write-NCMessage "Created group: $groupName" -Level SUCCESS
                            Write-NCMessage "Group ID: $($newGroup.id)" -Level INFO

                            if ($memberIds.Count -gt 0) {
                                try {
                                    $batchSize = 20
                                    for ($i = 0; $i -lt $memberIds.Count; $i += $batchSize) {
                                        $batch = $memberIds[$i..([Math]::Min($i + $batchSize - 1, $memberIds.Count - 1))]
                                        $addMembersBody = @{
                                            'members@odata.bind' = $batch
                                        } | ConvertTo-Json -Depth 10

                                        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($newGroup.id)" -Method PATCH -Body $addMembersBody -ContentType 'application/json'
                                        Write-Verbose "Added batch of $($batch.Count) members"
                                    }

                                    Write-NCMessage "Added $($memberIds.Count) devices to group" -Level SUCCESS
                                    Write-NCMessage "Added devices:" -Level INFO
                                    foreach ($device in $entraDevices) {
                                        Write-NCMessage "  • $($device.DeviceName)" -Level INFO
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
                            Write-NCMessage "Group body: $groupBody" -Level VERBOSE
                        }
                    }
                }

                $totalDevicesProcessed += $deviceCount
            }

            Write-Progress -Activity 'Processing App Groups' -Completed

            Write-NCMessage "APP-BASED GROUP CREATION SUMMARY" -Level INFO
            Write-NCMessage "===================================" -Level INFO
            Write-NCMessage "Applications matched: $($appDeviceMap.Count)" -Level INFO
            Write-NCMessage "Total devices processed: $totalDevicesProcessed" -Level INFO
            Write-NCMessage "Groups created: $groupsCreated" -Level INFO
            Write-NCMessage "Groups updated: $groupsUpdated" -Level INFO

            if ($DryRun.IsPresent) {
                Write-NCMessage "[DRY RUN] No changes were made" -Level INFO
            }

            if ($appDeviceMap.Count -gt 0) {
                Write-NCMessage "Top Applications by Device Count:" -Level INFO
                $appDeviceMap.GetEnumerator() |
                    Sort-Object { $_.Value.Devices.Count } -Descending |
                    Select-Object -First 10 |
                    ForEach-Object {
                        $deviceCount = ($_.Value.Devices | Select-Object -Property DeviceId -Unique).Count
                        Write-NCMessage "  • $($_.Key): $deviceCount devices" -Level INFO
                    }
            }

            Write-NCMessage "App-based group creation completed successfully." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Script execution failed: $($_.Exception.Message)" -Level ERROR
            return
        }
    }
}
