#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Intune helpers =======================================================================================================================

function Get-IntuneProfileAssignmentsByGroup {
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

            Write-Information "Starting app inventory reporting..." -InformationAction Continue

            # 1) Pull devices
            $devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
            if ($MaxDevices -gt 0) {
                $devicesUri += "?`$top=$MaxDevices"
            }
            $devices = @(Invoke-NCGraphAllPagesCore -Uri $devicesUri)
            if ($FilterByPlatform -ne "All") {
                $devices = $devices | Where-Object { $_.operatingSystem -like "$FilterByPlatform*" }
            }
            Write-Information "✓ Devices retrieved: $($devices.Count)" -InformationAction Continue

            # 2) Build app->device mapping from Detected Apps
            $appDeviceMap = @{}
            $processed = 0
            foreach ($device in $devices) {
                $processed++
                Write-Progress -Activity "Reading Detected Apps" -Status "$processed / $($devices.Count)" -PercentComplete (($processed / [Math]::Max($devices.Count,1)) * 100)
                try {
                    $deviceAppsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($device.id)?`$expand=detectedApps"
                    $deviceWithApps = Invoke-MgGraphRequest -Uri $deviceAppsUri -Method GET -ErrorAction Stop
                    foreach ($app in ($deviceWithApps.detectedApps | Where-Object { $_.displayName -like $ApplicationName })) {
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
                            Publisher  = $app.publisher
                            Source     = "DetectedApps"
                        }
                        if ($app.version)   { $appDeviceMap[$key].Versions[$app.version]   = ($appDeviceMap[$key].Versions[$app.version]   + 1) }
                        if ($app.publisher) { $appDeviceMap[$key].Publishers[$app.publisher] = ($appDeviceMap[$key].Publishers[$app.publisher] + 1) }
                    }
                    Start-Sleep -Milliseconds 40
                }
                catch {
                    if ($_.Exception.Message -like "*429*") {
                        Write-Information "`nRate limit hit, waiting 60 seconds..." -InformationAction Continue
                        Start-Sleep -Seconds 60
                        $processed--
                        continue
                    }
                    Write-Warning "Error reading apps for $($device.deviceName): $($_.Exception.Message)"
                }
            }
            Write-Progress -Activity "Reading Detected Apps" -Completed

            # 3) Optionally incorporate deployment statuses (broadens coverage)
            if ($IncludeDeployedApps) {
                Write-Information "Including deployed apps device status..." -InformationAction Continue
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
                                Publisher    = $null
                                InstallState = $s.installState
                                AppType      = $appType
                                Source       = "DeploymentStatus"
                            }
                        }
                    }
                }
            }

            # 4) Optional app type filter when only Detected Apps were used
            if (-not $IncludeDeployedApps -and $FilterByType -ne "All") {
                Write-Verbose "FilterByType applies only when -IncludeDeployedApps is used. Skipping type filter for Detected Apps only."
            }

            # 5) Build flat rows
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
                        User         = $dev.User
                        InstallState = $dev.InstallState
                        Source       = $dev.Source
                    }
                }
            }

            if (-not $rows -or $rows.Count -eq 0) {
                Write-Warning "No matches found for '$ApplicationName' with the provided filters."
                return
            }

            # 6) Console table output
            $rows | Sort-Object AppName, DeviceName | Format-Table -AutoSize AppName, Version, Publisher, AppType, DeviceName, Platform, User, InstallState, Source

            # 7) Optional exports
            if ($OutputCsvPath) {
                try {
                    $rows | Sort-Object AppName, DeviceName | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
                    Write-Information "✓ CSV exported to $OutputCsvPath" -InformationAction Continue
                }
                catch {
                    Write-Warning "Failed to export CSV: $($_.Exception.Message)"
                }
            }

            if ($OutputJsonPath) {
                try {
                    $rows | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputJsonPath -Encoding UTF8
                    Write-Information "✓ JSON exported to $OutputJsonPath" -InformationAction Continue
                }
                catch {
                    Write-Warning "Failed to export JSON: $($_.Exception.Message)"
                }
            }

            # 8) Optional per-app pivot/summary
            if ($PivotSummary) {
                Write-Host "`n=== SUMMARY BY APPLICATION ===" -ForegroundColor Cyan
                $rows | Group-Object AppName | Sort-Object Count -Descending | ForEach-Object {
                    $app = $_.Name
                    $count = $_.Count
                    $versions = ($_.Group | Where-Object Version | Group-Object Version | Sort-Object Count -Descending | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Count }) -join ", "
                    $publishers = ($_.Group | Where-Object Publisher | Group-Object Publisher | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "{0} ({1})" -f $_.Name, $_.Count }) -join ", "
                    "• {0}: {1} devices`n    Versions: {2}`n    Top publishers: {3}" -f $app, $count, ($(if ($versions) { $versions } else { 'n/a' })), ($(if ($publishers) { $publishers } else { 'n/a' }))
                }
            }

            Write-Information "`n🎉 Reporting completed successfully!" -InformationAction Continue
        }
        catch {
            Write-Error "Script execution failed: $($_.Exception.Message)"
            exit 1
        }
    }
}
