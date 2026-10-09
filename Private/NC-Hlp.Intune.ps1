#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Intune helpers =============================================================================================================

function Convert-NCGraphObjectToPlainObject {
    param($Object)

    if ($null -eq $Object) { return $null }

    try {
        return (($Object | ConvertTo-Json -Depth 20 -Compress) | ConvertFrom-Json -Depth 20)
    }
    catch {
        return $Object
    }
}

function Get-NCCoreProperty {
    param($Object, [string[]]$Names)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($Object.Contains($name) -and -not [string]::IsNullOrWhiteSpace([string]$Object[$name])) {
                return $Object[$name]
            }
        }
    }

    $additionalProperties = $Object.PSObject.Properties['AdditionalProperties']
    if ($additionalProperties -and $additionalProperties.Value -is [System.Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($additionalProperties.Value.Contains($name) -and -not [string]::IsNullOrWhiteSpace([string]$additionalProperties.Value[$name])) {
                return $additionalProperties.Value[$name]
            }
        }
    }

    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return $property.Value
        }
    }
    return $null
}

function Resolve-NCIntuneManagedDeviceEntraMembers {
    <#
    .SYNOPSIS
        Resolves the Entra device object of many Intune managed devices using Graph batches.
    .DESCRIPTION
        Refreshes the Azure AD device id of devices that lack it (beta managedDevices, 20 per request) and
        looks up the matching Entra device (v1.0 devices filter, 20 per request). Returns one object per
        input id, in input order, with DeviceId, Resolution (null when the device cannot be resolved) and
        LookupFailed (true when the device could not be resolved for any reason: a failed Graph read, a missing
        Azure AD device ID, or no matching Entra device. Its Entra state is then unknown, so callers must not
        treat it as absent).
        Warnings and errors use the same text as the single-device resolver.
    .PARAMETER ManagedDevices
        Managed devices already retrieved from Intune.
    .PARAMETER DeviceIds
        Intune managed device ids to resolve.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ManagedDevices,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$DeviceIds
    )

    if (@($DeviceIds).Count -eq 0) {
        return
    }

    $devicesById = @{}
    foreach ($managedDevice in @($ManagedDevices)) {
        $key = [string]$managedDevice.id
        if (-not $devicesById.ContainsKey($key)) { $devicesById[$key] = $managedDevice }
    }

    $entries = @(foreach ($deviceId in $DeviceIds) {
            $plain = if ($devicesById.ContainsKey([string]$deviceId)) { Convert-NCGraphObjectToPlainObject -Object $devicesById[[string]$deviceId] } else { $null }
            $label = Get-NCCoreProperty -Object $plain -Names @('deviceName', 'DeviceName', 'displayName', 'DisplayName', 'id', 'Id')
            if ([string]::IsNullOrWhiteSpace($label)) { $label = [string]$deviceId }
            [pscustomobject]@{
                DeviceId        = [string]$deviceId
                Plain           = $plain
                Label           = $label
                AzureAdDeviceId = (Get-NCCoreProperty -Object $plain -Names @('azureADDeviceId', 'azureActiveDirectoryDeviceId', 'azureAdDeviceId'))
                DetailError     = $null
                EntraDevice     = $null
                LookupError     = $null
            }
        })

    # Refresh the Azure AD device id for devices that do not carry it
    $refresh = @($entries | Where-Object { $_.Plain -and -not $_.AzureAdDeviceId })
    if ($refresh.Count -gt 0) {
        $refreshRequests = @(for ($i = 0; $i -lt $refresh.Count; $i++) {
                @{ Id = "r$i"; Method = 'GET'; Url = "/deviceManagement/managedDevices/$([uri]::EscapeDataString($refresh[$i].DeviceId))?`$select=id,deviceName,operatingSystem,userPrincipalName,azureADDeviceId,azureActiveDirectoryDeviceId" }
            })
        $refreshResponses = @(Invoke-NCGraphBatch -Requests $refreshRequests -ApiVersion 'beta' -Activity 'Reading Intune device details')
        for ($i = 0; $i -lt $refresh.Count; $i++) {
            if ($refreshResponses[$i].Success) {
                $detailsPlain = Convert-NCGraphObjectToPlainObject -Object $refreshResponses[$i].Body
                $refresh[$i].AzureAdDeviceId = Get-NCCoreProperty -Object $detailsPlain -Names @('azureADDeviceId', 'azureActiveDirectoryDeviceId', 'azureAdDeviceId')
            }
            else {
                $refresh[$i].DetailError = $refreshResponses[$i].ErrorMessage
            }
        }
    }

    # Look up the Entra device objects
    $lookup = @($entries | Where-Object { $_.Plain -and $_.AzureAdDeviceId })
    if ($lookup.Count -gt 0) {
        $lookupRequests = @(for ($i = 0; $i -lt $lookup.Count; $i++) {
                $filter = "deviceId eq '$(([string]$lookup[$i].AzureAdDeviceId).Replace("'", "''"))'"
                @{ Id = "e$i"; Method = 'GET'; Url = "/devices?`$filter=$([uri]::EscapeDataString($filter))" }
            })
        $lookupResponses = @(Invoke-NCGraphBatch -Requests $lookupRequests -ApiVersion 'v1.0' -Activity 'Resolving Entra devices')
        for ($i = 0; $i -lt $lookup.Count; $i++) {
            if ($lookupResponses[$i].Success) {
                $values = @($lookupResponses[$i].Body.value)
                if ($values.Count -gt 0 -and $null -ne $values[0]) { $lookup[$i].EntraDevice = $values[0] }
            }
            else {
                $lookup[$i].LookupError = $lookupResponses[$i].ErrorMessage
            }
        }
    }

    foreach ($entry in $entries) {
        if ($entry.DetailError) {
            Write-NCMessage "Unable to refresh Intune device details for $($entry.Label): $($entry.DetailError)" -Level WARNING
        }

        $resolution = $null
        if (-not $entry.Plain -or -not $entry.AzureAdDeviceId) {
            Write-NCMessage "No Azure AD Device ID for: $($entry.Label)" -Level WARNING
        }
        elseif ($entry.LookupError) {
            Write-NCMessage "Error looking up Entra ID device for $($entry.Label): $($entry.LookupError)" -Level ERROR
        }
        elseif ($entry.EntraDevice) {
            Write-Verbose "Found Entra ID device: $($entry.Label) -> $($entry.EntraDevice.id)"
            $resolution = [pscustomobject]@{
                IntuneDeviceId  = $entry.DeviceId
                EntraDeviceId   = $entry.EntraDevice.id
                DeviceName      = $entry.Label
                AzureAdDeviceId = $entry.AzureAdDeviceId
            }
        }
        else {
            Write-NCMessage "Device not found in Entra ID: $($entry.Label) (Azure AD Device ID: $($entry.AzureAdDeviceId))" -Level WARNING
        }

        # Any unresolved device leaves its Entra state unknown, not only a failed Graph read
        [pscustomobject]@{ DeviceId = $entry.DeviceId; Resolution = $resolution; LookupFailed = ($null -eq $resolution) }
    }
}

function Resolve-NCIntuneManagedDeviceEntraMember {
    <#
    .SYNOPSIS
        Resolves the Entra device object of one Intune managed device.
    .DESCRIPTION
        Single-device wrapper around Resolve-NCIntuneManagedDeviceEntraMembers.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ManagedDevices,
        [Parameter(Mandatory = $true)]
        [string]$DeviceId
    )

    return (@(Resolve-NCIntuneManagedDeviceEntraMembers -ManagedDevices $ManagedDevices -DeviceIds @($DeviceId)) | Select-Object -First 1).Resolution
}

function Invoke-NCIntuneGroupUsageCore {
    <#
    .SYNOPSIS
        Core Intune group usage lookup logic.
    .DESCRIPTION
        Resolves the requested Entra group, scans supported Intune surfaces, and returns matching assignment
        records. This is the implementation behind the public group usage command.
    .PARAMETER ParameterSetName
        Active parameter set name from the public wrapper.
    .PARAMETER GroupName
        Target group display name.
    .PARAMETER GroupId
        Target group object ID.
    .PARAMETER ProfileName
        Optional filter for the profile or app display name.
    .PARAMETER ProfileId
        Optional filter for a specific Intune object ID.
    .PARAMETER IncludeNestedGroups
        Also include parent groups that contain the requested group.
    .PARAMETER GridView
        Show results in Out-GridView when requested.
    .PARAMETER Diagnostic
        Include diagnostic columns in the returned objects.
    #>
    [CmdletBinding()]
    param(
        [string]$ParameterSetName,
        [string]$GroupName,
        [string]$GroupId,
        [string]$ProfileName,
        [string]$ProfileId,
        [switch]$IncludeNestedGroups,
        [switch]$GridView,
        [switch]$Diagnostic
    )

    function Invoke-NCGraphPagedRequestCore {
        param([Parameter(Mandatory = $true)][string]$Uri)

        $items = [System.Collections.Generic.List[object]]::new()
        $nextUri = $Uri
        while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
            $response = Invoke-MgGraphRequest -Uri $nextUri -Method Get -ErrorAction Stop
            $pageItems = @()

            if ($response -is [System.Collections.IDictionary]) {
                if ($response.Contains('value')) { $pageItems = @($response['value']) }
                elseif ($response.Contains('id')) { $pageItems = @($response) }
                if ($response.Contains('@odata.nextLink')) { $nextUri = [string]$response['@odata.nextLink'] } else { $nextUri = $null }
            }
            else {
                if ($response.PSObject.Properties['value']) { $pageItems = @($response.value) }
                elseif ($response.PSObject.Properties['id']) { $pageItems = @($response) }
                if ($response.PSObject.Properties['@odata.nextLink']) { $nextUri = [string]$response.'@odata.nextLink' } else { $nextUri = $null }
            }

            foreach ($item in $pageItems) {
                if ($null -ne $item) { $items.Add($item) | Out-Null }
            }
        }

        return @($items)
    }

    function Test-NCCoreNameMatch {
        param([string]$CandidateName, [string]$FilterName)
        if ([string]::IsNullOrWhiteSpace($FilterName)) { return $true }
        if ([string]::IsNullOrWhiteSpace($CandidateName)) { return $false }
        return $CandidateName -like "*$FilterName*"
    }

    function Resolve-NCCoreGroupName {
        param([string]$Id)
        if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
        if ($script:NCIntuneGroupNameCache -and $script:NCIntuneGroupNameCache.ContainsKey($Id)) {
            return $script:NCIntuneGroupNameCache[$Id]
        }
        if (-not $script:NCIntuneGroupNameCache) {
            $script:NCIntuneGroupNameCache = @{}
        }
        $name = $null
        try {
            $groupInfo = Get-MgGroup -GroupId $Id -ErrorAction Stop
            $name = $groupInfo.DisplayName
        }
        catch {
            $name = $null
        }
        $script:NCIntuneGroupNameCache[$Id] = $name
        return $name
    }

    function Get-NCCoreEffectiveGroupIds {
        param([string]$RootGroupId)

        $ids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        [void]$ids.Add($RootGroupId)

        if (-not $IncludeNestedGroups.IsPresent) {
            return $ids
        }

        try {
            $parents = @(Get-MgGroupTransitiveMemberOf -GroupId $RootGroupId -All -ErrorAction Stop)
        }
        catch {
            try {
                $parents = @(Get-MgGroupMemberOf -GroupId $RootGroupId -All -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to resolve parent groups for '$RootGroupId': $($_.Exception.Message)" -Level WARNING
                return $ids
            }
        }

        foreach ($parent in $parents) {
            $odataType = $null
            if ($parent.AdditionalProperties -and $parent.AdditionalProperties.ContainsKey('@odata.type')) {
                $odataType = [string]$parent.AdditionalProperties['@odata.type']
            }
            if ($odataType -and $odataType -notmatch 'group') { continue }
            $parentId = [string](Get-NCCoreProperty -Object $parent -Names @('id', 'Id'))
            if (-not [string]::IsNullOrWhiteSpace($parentId)) {
                [void]$ids.Add($parentId)
            }
        }

        return $ids
    }

    function Initialize-NCCoreGroupNames {
        param([string[]]$Ids)

        if (-not $script:NCIntuneGroupNameCache) {
            $script:NCIntuneGroupNameCache = @{}
        }

        $missing = @($Ids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $script:NCIntuneGroupNameCache.ContainsKey($_) } | Select-Object -Unique)
        if ($missing.Count -eq 0) { return }

        $requests = @(for ($i = 0; $i -lt $missing.Count; $i++) {
                @{ Id = "g$i"; Method = 'GET'; Url = "/groups/$([uri]::EscapeDataString($missing[$i]))?`$select=id,displayName" }
            })
        $responses = @(Invoke-NCGraphBatch -Requests $requests -ApiVersion 'v1.0' -Activity 'Resolving group names')
        for ($i = 0; $i -lt $missing.Count; $i++) {
            $name = $null
            if ($responses[$i].Success -and $null -ne $responses[$i].Body) {
                $name = $responses[$i].Body.displayName
            }
            $script:NCIntuneGroupNameCache[$missing[$i]] = $name
        }
    }

    function Get-NCCoreAssignmentRecordsBatch {
        param(
            [string]$EntityType,
            [string[]]$EntityIds,
            [System.Collections.Generic.HashSet[string]]$EffectiveGroupIds,
            [string]$RequestedGroupId
        )

        $map = @{}
        $ids = @($EntityIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        if ($ids.Count -eq 0) { return $map }

        $segment = switch ($EntityType) {
            'deviceConfigurations' { 'deviceManagement/deviceConfigurations' }
            'configurationPolicies' { 'deviceManagement/configurationPolicies' }
            'mobileApps' { 'deviceAppManagement/mobileApps' }
            default { $null }
        }
        if (-not $segment) {
            foreach ($id in $ids) { $map[$id] = @() }
            return $map
        }

        $requests = @(for ($i = 0; $i -lt $ids.Count; $i++) {
                @{ Id = "a$i"; Method = 'GET'; Url = "/$segment('$([uri]::EscapeDataString($ids[$i]))')/assignments" }
            })
        $responses = @(Invoke-NCGraphBatchCollection -Requests $requests -ApiVersion 'beta' -Activity "Reading $EntityType assignments")

        $pending = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $ids.Count; $i++) {
            $map[$ids[$i]] = @()
            if (-not $responses[$i].Success) {
                Write-NCMessage "Unable to read assignments for '$EntityType' object '$($ids[$i])': $($responses[$i].ErrorMessage)" -Level WARNING
                continue
            }

            foreach ($assignment in @($responses[$i].Items)) {
                if ($null -eq $assignment) { continue }
                $target = Get-NCCoreProperty -Object $assignment -Names @('target', 'Target')
                if (-not $target) { continue }

                $odataType = [string](Get-NCCoreProperty -Object $target -Names @('@odata.type'))
                $targetGroupId = [string](Get-NCCoreProperty -Object $target -Names @('groupId', 'GroupId'))
                $intent = [string](Get-NCCoreProperty -Object $assignment -Names @('intent', 'Intent'))
                if (-not [string]::IsNullOrWhiteSpace($intent)) { $intent = $intent.ToLowerInvariant() }

                $reason = $null
                switch ($odataType) {
                    '#microsoft.graph.groupAssignmentTarget' {
                        if ($EffectiveGroupIds.Contains($targetGroupId)) {
                            $reason = if ($targetGroupId -eq $RequestedGroupId) { 'Direct Assignment' } else { 'Group Assignment' }
                        }
                    }
                    '#microsoft.graph.exclusionGroupAssignmentTarget' {
                        if ($EffectiveGroupIds.Contains($targetGroupId)) {
                            $reason = if ($targetGroupId -eq $RequestedGroupId) { 'Direct Exclusion' } else { 'Group Exclusion' }
                        }
                    }
                }

                if (-not $reason -and -not $Diagnostic.IsPresent) { continue }

                $pending.Add([pscustomobject]@{
                        EntityId        = $ids[$i]
                        Reason          = $reason
                        GroupId         = $targetGroupId
                        AssignmentId    = [string](Get-NCCoreProperty -Object $assignment -Names @('id', 'Id'))
                        TargetODataType = $odataType
                        Intent          = $intent
                    }) | Out-Null
            }
        }

        Initialize-NCCoreGroupNames -Ids @($pending | ForEach-Object { $_.GroupId })

        $records = @{}
        foreach ($entry in $pending) {
            if (-not $records.ContainsKey($entry.EntityId)) {
                $records[$entry.EntityId] = [System.Collections.Generic.List[object]]::new()
            }
            $records[$entry.EntityId].Add([pscustomobject]@{
                    Reason          = $entry.Reason
                    GroupId         = $entry.GroupId
                    GroupName       = Resolve-NCCoreGroupName -Id $entry.GroupId
                    AssignmentId    = $entry.AssignmentId
                    TargetODataType = $entry.TargetODataType
                    Intent          = $entry.Intent
                }) | Out-Null
        }
        foreach ($id in $records.Keys) {
            $map[$id] = @($records[$id])
        }

        return $map
    }

    function Resolve-NCCoreAssignmentValue {
        param([object[]]$Assignments)

        $hasInclude = @($Assignments | Where-Object { $_.Reason -in @('Direct Assignment', 'Group Assignment') }).Count -gt 0
        $hasExclude = @($Assignments | Where-Object { $_.Reason -in @('Direct Exclusion', 'Group Exclusion') }).Count -gt 0

        if ($hasInclude -and $hasExclude) { return 'Include; Exclude' }
        if ($hasExclude) { return 'Exclude' }
        if ($hasInclude) { return 'Include' }
        return $null
    }

    function Add-NCCoreResult {
        param(
            [System.Collections.Generic.List[object]]$Results,
            [string]$Category,
            $Item,
            [string]$Source,
            [object[]]$Assignments,
            $ResolvedGroup,
            [string]$AppIntent
        )

        $itemId = [string](Get-NCCoreProperty -Object $Item -Names @('id', 'Id'))
        $itemName = [string](Get-NCCoreProperty -Object $Item -Names @('displayName', 'DisplayName', 'name', 'Name'))
        $itemType = [string](Get-NCCoreProperty -Object $Item -Names @('@odata.type'))

        if ($ProfileId -and $itemId -ne $ProfileId) { return }
        if (-not (Test-NCCoreNameMatch -CandidateName $itemName -FilterName $ProfileName)) { return }

        $assignmentValue = Resolve-NCCoreAssignmentValue -Assignments $Assignments
        if (-not $assignmentValue -and -not $Diagnostic.IsPresent) { return }

        $row = [ordered]@{
            'Category'     = $Category
            'Profile Name' = $itemName
            'Profile Type' = $itemType
            'Assignment'   = $assignmentValue
        }

        if ($GridView.IsPresent -or $Diagnostic.IsPresent) {
            $row['Profile Id'] = $itemId
            $row['Source'] = $Source
            $row['Group Name'] = $ResolvedGroup.DisplayName
            $row['Group Id'] = $ResolvedGroup.Id
            $row['Assignment Id'] = (($Assignments | ForEach-Object { $_.AssignmentId } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
            $row['Target OData Type'] = (($Assignments | ForEach-Object { $_.TargetODataType } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
            $row['Target Group Id'] = (($Assignments | ForEach-Object { $_.GroupId } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
            $row['Target Group Name'] = (($Assignments | ForEach-Object { $_.GroupName } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
            $row['Matched Requested Group'] = (@($Assignments | Where-Object { $_.Reason }).Count -gt 0)
            if ($AppIntent) { $row['App Intent'] = $AppIntent }
        }

        $resultObject = [pscustomobject]$row
        $resultObject.PSObject.TypeNames.Insert(0, 'Nebula.Core.IntuneGroupUsage')
        $Results.Add($resultObject) | Out-Null
    }

    if (-not (Test-MgGraphConnection -Scopes @('DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All', 'Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false)) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return
    }

    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
        return
    }

    try {
        if ($ParameterSetName -eq 'ById') {
            $resolvedGroup = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        }
        else {
            $groupCandidates = @(Get-MgGroup -Filter "displayName eq '$GroupName'" -ConsistencyLevel eventual -CountVariable ignored -All -ErrorAction Stop)
            if ($groupCandidates.Count -eq 0) {
                Write-NCMessage "No Entra group found with display name '$GroupName'." -Level WARNING
                return
            }
            if ($groupCandidates.Count -gt 1) {
                Write-NCMessage "Multiple Entra groups found with display name '$GroupName'. Use -GroupId instead." -Level ERROR
                return
            }
            $resolvedGroup = $groupCandidates[0]
        }
    }
    catch {
        Write-NCMessage "Unable to resolve target group: $($_.Exception.Message)" -Level ERROR
        return
    }

    $effectiveGroupIds = Get-NCCoreEffectiveGroupIds -RootGroupId $resolvedGroup.Id
    $results = [System.Collections.Generic.List[object]]::new()

    $deviceConfigurations = @()
    $configurationPolicies = @()
    $mobileApps = @()

    try { $deviceConfigurations = @(Invoke-NCGraphPagedRequestCore -Uri 'beta/deviceManagement/deviceConfigurations') }
    catch { Write-NCMessage "Unable to retrieve Intune device configurations: $($_.Exception.Message)" -Level WARNING }

    try { $configurationPolicies = @(Invoke-NCGraphPagedRequestCore -Uri 'beta/deviceManagement/configurationPolicies') }
    catch { Write-NCMessage "Unable to retrieve Intune configuration policies: $($_.Exception.Message)" -Level WARNING }

    try { $mobileApps = @(Invoke-NCGraphPagedRequestCore -Uri "beta/deviceAppManagement/mobileApps?`$filter=isAssigned eq true") }
    catch { Write-NCMessage "Unable to retrieve Intune mobile apps: $($_.Exception.Message)" -Level WARNING }

    $scannedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($collection in @($deviceConfigurations, $configurationPolicies, $mobileApps)) {
        foreach ($entry in @($collection)) {
            $entryId = [string](Get-NCCoreProperty -Object $entry -Names @('id', 'Id'))
            if (-not [string]::IsNullOrWhiteSpace($entryId)) { [void]$scannedIds.Add($entryId) }
        }
    }

    Add-EmptyLine
    Write-Verbose "Scanning $($scannedIds.Count) Intune profile(s) for group '$($resolvedGroup.DisplayName)' ..."

    $getEntityId = { param($entity) [string](Get-NCCoreProperty -Object $entity -Names @('id', 'Id')) }
    $configurationEntities = @($deviceConfigurations | Where-Object { -not [string]::IsNullOrWhiteSpace((& $getEntityId $_)) })
    $policyEntities = @($configurationPolicies | Where-Object { -not [string]::IsNullOrWhiteSpace((& $getEntityId $_)) })
    $appEntities = @($mobileApps | Where-Object { -not ($_.isFeatured -or $_.isBuiltIn) -and -not [string]::IsNullOrWhiteSpace((& $getEntityId $_)) })

    $profilesToRead = $configurationEntities.Count + $policyEntities.Count + $appEntities.Count
    if ($profilesToRead -gt 0) {
        Write-NCGraphBatchNotice -Count $profilesToRead -Noun 'Intune profile(s)'
    }

    $configurationAssignments = Get-NCCoreAssignmentRecordsBatch -EntityType 'deviceConfigurations' -EntityIds @($configurationEntities | ForEach-Object { & $getEntityId $_ }) -EffectiveGroupIds $effectiveGroupIds -RequestedGroupId $resolvedGroup.Id
    foreach ($entity in $configurationEntities) {
        $assignments = @($configurationAssignments[(& $getEntityId $entity)])
        Add-NCCoreResult -Results $results -Category 'Device Configuration' -Item $entity -Source 'deviceConfigurations' -Assignments $assignments -ResolvedGroup $resolvedGroup
    }

    $policyAssignments = Get-NCCoreAssignmentRecordsBatch -EntityType 'configurationPolicies' -EntityIds @($policyEntities | ForEach-Object { & $getEntityId $_ }) -EffectiveGroupIds $effectiveGroupIds -RequestedGroupId $resolvedGroup.Id
    foreach ($entity in $policyEntities) {
        $assignments = @($policyAssignments[(& $getEntityId $entity)])
        Add-NCCoreResult -Results $results -Category 'Settings Catalog Policy' -Item $entity -Source 'configurationPolicies' -Assignments $assignments -ResolvedGroup $resolvedGroup
    }

    $appAssignments = Get-NCCoreAssignmentRecordsBatch -EntityType 'mobileApps' -EntityIds @($appEntities | ForEach-Object { & $getEntityId $_ }) -EffectiveGroupIds $effectiveGroupIds -RequestedGroupId $resolvedGroup.Id
    foreach ($entity in $appEntities) {
        $entityId = & $getEntityId $entity
        $assignments = @($appAssignments[$entityId])
        if ($assignments.Count -eq 0 -and -not $Diagnostic.IsPresent) { continue }

        $intentGroups = @{}
        foreach ($assignment in $assignments) {
            if ([string]::IsNullOrWhiteSpace($assignment.Intent)) { continue }
            if (-not $intentGroups.ContainsKey($assignment.Intent)) {
                $intentGroups[$assignment.Intent] = [System.Collections.Generic.List[object]]::new()
            }
            $intentGroups[$assignment.Intent].Add($assignment) | Out-Null
        }

        foreach ($intent in $intentGroups.Keys) {
            $category = switch ($intent) {
                'required' { 'Required App' }
                'available' { 'Available App' }
                'uninstall' { 'Uninstall App' }
                default { 'Assigned App' }
            }
            Add-NCCoreResult -Results $results -Category $category -Item $entity -Source 'mobileApps' -Assignments @($intentGroups[$intent]) -ResolvedGroup $resolvedGroup -AppIntent $intent
        }
    }

    $sorted = $results | Sort-Object 'Category', 'Profile Name', 'Assignment' -Unique

    Add-EmptyLine
    Write-Verbose "Intune profiles found for '$($resolvedGroup.DisplayName)': $($sorted.Count)"

    if ($sorted.Count -eq 0) {
        Write-NCMessage "No Intune profiles found for group '$($resolvedGroup.DisplayName)'." -Level WARNING
        return
    }

    if ($GridView.IsPresent) {
        $sorted | Out-NCGridView -Title "Intune Profiles - $($resolvedGroup.DisplayName)"
    }
    else {
        $sorted
    }
}

function Invoke-NCGraphCollectionRequest {
    <#
    .SYNOPSIS
        Retrieves a Graph collection with pagination.
    .DESCRIPTION
        Repeatedly calls Microsoft Graph until no next link remains, with light throttling and basic 429 retry
        handling.
    .PARAMETER Uri
        Initial Graph request URI.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    $requestCount = 0

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        try {
            if ($requestCount -gt 0) {
                Start-Sleep -Milliseconds 100
            }

            $response = Invoke-MgGraphRequest -Uri $nextUri -Method GET -ErrorAction Stop
            $requestCount++

            if ($response.PSObject.Properties['value']) {
                $items.AddRange(@($response.value))
            }
            else {
                $items.AddRange(@($response))
            }

            if ($requestCount % 10 -eq 0) {
                Write-Verbose "."
            }

            $nextLink = $response.'@odata.nextLink'
            $nextUri = if ($nextLink) { [string]$nextLink } else { $null }
        }
        catch {
            if ($_.Exception.Message -like '*429*') {
                Write-Verbose "`nRate limit hit, waiting 60 seconds ..."
                Start-Sleep -Seconds 60
                continue
            }

            # A failed page leaves the collection incomplete: let the caller decide instead of returning partial data
            throw
        }
    }

    return @($items)
}

function Invoke-NCGraphAllPagesCore {
    <#
    .SYNOPSIS
        Retrieves all pages from a Graph endpoint.
    .DESCRIPTION
        Follows @odata.nextLink until the collection is exhausted and returns all items as an array.
    .PARAMETER Uri
        Initial Graph request URI.
    .PARAMETER DelayMs
        Delay in milliseconds between follow-up page requests.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [int]$DelayMs = 100
    )

    $allResults = @()
    $nextLink = $Uri
    $requestCount = 0

    do {
        try {
            if ($requestCount -gt 0) {
                Start-Sleep -Milliseconds $DelayMs
            }

            $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
            $requestCount++

            $hasValueProperty = $false
            if ($response -is [hashtable] -or $response -is [System.Collections.IDictionary]) {
                $hasValueProperty = $response.ContainsKey('value')
            }
            else {
                $hasValueProperty = $null -ne $response.PSObject.Properties['value']
            }

            if ($hasValueProperty) {
                $allResults += @($response.value)
            }
            else {
                $allResults += $response
            }

            $nextLink = $response.'@odata.nextLink'
            if ($requestCount % 10 -eq 0) {
                Write-Verbose "."
            }
        }
        catch {
            if ($_.Exception.Message -like "*429*") {
                Write-Verbose "`nRate limit hit, waiting 60 seconds ..."
                Start-Sleep -Seconds 60
                continue
            }

            # A failed page leaves the collection incomplete: let the caller decide instead of returning partial data
            throw
        }
    }
    while ($nextLink)

    return $allResults
}

function Get-NCIntuneItemName {
    <#
    .SYNOPSIS
        Returns the best display name for an Intune object.
    .DESCRIPTION
        Reads common display-name properties and returns the first non-empty value.
    .PARAMETER Item
        Intune object to inspect.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Item
    )

    if (-not $Item) { return $null }

    foreach ($propertyName in @('displayName', 'DisplayName', 'name', 'Name')) {
        $property = $Item.PSObject.Properties[$propertyName]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return $null
}

function Get-NCIntuneItemId {
    <#
    .SYNOPSIS
        Returns the best identifier for an Intune object.
    .DESCRIPTION
        Reads common id properties and returns the first non-empty value.
    .PARAMETER Item
        Intune object to inspect.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Item
    )

    if (-not $Item) { return $null }

    foreach ($propertyName in @('id', 'Id')) {
        $property = $Item.PSObject.Properties[$propertyName]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return $null
}

function Get-NCIntuneItemODataType {
    <#
    .SYNOPSIS
        Returns the OData type for an Intune object.
    .DESCRIPTION
        Reads the standard @odata.type property or falls back to AdditionalProperties when available.
    .PARAMETER Item
        Intune object to inspect.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Item
    )

    if (-not $Item) { return $null }

    $odataProperty = $Item.PSObject.Properties['@odata.type']
    if ($odataProperty -and -not [string]::IsNullOrWhiteSpace([string]$odataProperty.Value)) {
        return [string]$odataProperty.Value
    }

    $additionalProperties = $Item.PSObject.Properties['AdditionalProperties']
    if ($additionalProperties -and $additionalProperties.Value -and $additionalProperties.Value.ContainsKey('@odata.type')) {
        return [string]$additionalProperties.Value['@odata.type']
    }

    return $null
}

function Get-NCIntuneSearchFields {
    <#
    .SYNOPSIS
        Returns searchable fields for an Intune object.
    .DESCRIPTION
        Collects common identifying and descriptive properties into a normalized search field list.
    .PARAMETER Item
        Intune object to inspect.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Item
    )

    $fields = [System.Collections.Generic.List[object]]::new()
    if (-not $Item) { return $fields.ToArray() }

    foreach ($propertyName in @('id', 'Id', 'displayName', 'DisplayName', 'name', 'Name', 'description', 'Description', 'networkName', 'NetworkName', 'ssid', 'Ssid')) {
        $property = $Item.PSObject.Properties[$propertyName]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            $fields.Add([pscustomobject]@{
                    Property = $propertyName
                    Value    = [string]$property.Value
                }) | Out-Null
        }
    }

    return $fields.ToArray()
}

function Get-NCIntuneAppTypeFromODataType {
    <#
    .SYNOPSIS
        Maps an Intune app OData type to a friendly app type.
    .DESCRIPTION
        Translates common Microsoft Graph Intune app type names into a short label used by the module.
    .PARAMETER ODataType
        OData type string to translate.
    #>
    [CmdletBinding()]
    param([string]$ODataType)

    switch ($ODataType) {
        '#microsoft.graph.win32LobApp' { return 'Win32' }
        '#microsoft.graph.microsoftStoreForBusinessApp' { return 'Store' }
        '#microsoft.graph.webApp' { return 'Web' }
        '#microsoft.graph.officeSuiteApp' { return 'Office' }
        '#microsoft.graph.winGetApp' { return 'WinGet' }
        '#microsoft.graph.iosLobApp' { return 'iOS' }
        '#microsoft.graph.iosStoreApp' { return 'iOS' }
        '#microsoft.graph.androidManagedStoreApp' { return 'Android' }
        '#microsoft.graph.androidLobApp' { return 'Android' }
        '#microsoft.graph.macOSLobApp' { return 'macOS' }
        '#microsoft.graph.macOSOfficeSuiteApp' { return 'macOS' }
        default { return 'Other' }
    }
}

function Test-NCIntuneVersionAtLeast {
    <#
    .SYNOPSIS
        Compares two version strings.
    .DESCRIPTION
        Returns $true when the current version is greater than or equal to the minimum version. Falls back
        to string comparison if the inputs cannot be parsed as [version].
    .PARAMETER CurrentVersion
        Current version string.
    .PARAMETER MinimumVersion
        Minimum version string to compare against.
    #>
    [CmdletBinding()]
    param(
        [string]$CurrentVersion,
        [string]$MinimumVersion
    )

    if ([string]::IsNullOrWhiteSpace($MinimumVersion)) {
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($CurrentVersion)) {
        return $false
    }

    $currentParsed = [version]'0.0'
    $minimumParsed = [version]'0.0'
    if ([version]::TryParse($CurrentVersion, [ref]$currentParsed) -and [version]::TryParse($MinimumVersion, [ref]$minimumParsed)) {
        return $currentParsed -ge $minimumParsed
    }

    return $CurrentVersion -ge $MinimumVersion
}

function Get-NCIntuneAppBasedGroupName {
    <#
    .SYNOPSIS
        Builds a sanitized Entra group name for app-based Intune groups.
    .DESCRIPTION
        Uses an explicit group name when provided, otherwise removes invalid characters, collapses
        separators, trims edges, and enforces the Entra group name length limit while preserving the
        requested prefix and suffix.
    .PARAMETER GroupName
        Explicit full group name to use as-is instead of generating one from prefix and suffix.
    .PARAMETER AppName
        Application display name used as the base for the group name.
    .PARAMETER GroupPrefix
        Prefix prepended to the sanitized application name.
    .PARAMETER GroupSuffix
        Suffix appended to the sanitized application name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$GroupName,
        [Parameter(Mandatory = $false)]
        [string]$AppName,
        [Parameter(Mandatory = $false)]
        [string]$GroupPrefix = 'Devices-With-',
        [Parameter(Mandatory = $false)]
        [string]$GroupSuffix = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($GroupName)) {
        return $GroupName
    }

    $sanitized = $AppName -replace '[^\w\s-]', ''
    $sanitized = $sanitized -replace '\s+', '-'
    $sanitized = $sanitized -replace '-+', '-'
    $sanitized = $sanitized.Trim('-')

    $maxLength = 256 - $GroupPrefix.Length - $GroupSuffix.Length
    if ($sanitized.Length -gt $maxLength) {
        $sanitized = $sanitized.Substring(0, $maxLength)
    }

    return "${GroupPrefix}${sanitized}${GroupSuffix}"
}
