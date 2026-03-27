#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Intune helpers =============================================================================================================

function Invoke-NCIntuneGroupUsageCore {
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
        foreach ($name in $Names) {
            $property = $Object.PSObject.Properties[$name]
            if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                return $property.Value
            }
        }
        return $null
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

    function Get-NCCoreAssignmentRecords {
        param(
            [string]$EntityType,
            [string]$EntityId,
            [System.Collections.Generic.HashSet[string]]$EffectiveGroupIds,
            [string]$RequestedGroupId
        )

        $uri = switch ($EntityType) {
            'deviceConfigurations' { "beta/deviceManagement/deviceConfigurations('$EntityId')/assignments" }
            'configurationPolicies' { "beta/deviceManagement/configurationPolicies('$EntityId')/assignments" }
            'mobileApps' { "beta/deviceAppManagement/mobileApps('$EntityId')/assignments" }
            default { $null }
        }
        if (-not $uri) { return @() }

        try {
            $assignments = @(Invoke-NCGraphPagedRequestCore -Uri $uri)
        }
        catch {
            Write-NCMessage "Unable to read assignments for '$EntityType' object '$EntityId': $($_.Exception.Message)" -Level WARNING
            return @()
        }

        $records = [System.Collections.Generic.List[object]]::new()
        foreach ($assignment in $assignments) {
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

            $records.Add([pscustomobject]@{
                    Reason          = $reason
                    GroupId         = $targetGroupId
                    GroupName       = Resolve-NCCoreGroupName -Id $targetGroupId
                    AssignmentId    = [string](Get-NCCoreProperty -Object $assignment -Names @('id', 'Id'))
                    TargetODataType = $odataType
                    Intent          = $intent
                }) | Out-Null
        }

        return @($records)
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
    Write-NCMessage "Scanning $($scannedIds.Count) Intune profile(s) for group '$($resolvedGroup.DisplayName)' ..." -Level VERBOSE

    foreach ($entity in $deviceConfigurations) {
        $entityId = [string](Get-NCCoreProperty -Object $entity -Names @('id', 'Id'))
        if ([string]::IsNullOrWhiteSpace($entityId)) { continue }
        $assignments = @(Get-NCCoreAssignmentRecords -EntityType 'deviceConfigurations' -EntityId $entityId -EffectiveGroupIds $effectiveGroupIds -RequestedGroupId $resolvedGroup.Id)
        Add-NCCoreResult -Results $results -Category 'Device Configuration' -Item $entity -Source 'deviceConfigurations' -Assignments $assignments -ResolvedGroup $resolvedGroup
    }

    foreach ($entity in $configurationPolicies) {
        $entityId = [string](Get-NCCoreProperty -Object $entity -Names @('id', 'Id'))
        if ([string]::IsNullOrWhiteSpace($entityId)) { continue }
        $assignments = @(Get-NCCoreAssignmentRecords -EntityType 'configurationPolicies' -EntityId $entityId -EffectiveGroupIds $effectiveGroupIds -RequestedGroupId $resolvedGroup.Id)
        Add-NCCoreResult -Results $results -Category 'Settings Catalog Policy' -Item $entity -Source 'configurationPolicies' -Assignments $assignments -ResolvedGroup $resolvedGroup
    }

    foreach ($entity in $mobileApps) {
        if ($entity.isFeatured -or $entity.isBuiltIn) { continue }
        $entityId = [string](Get-NCCoreProperty -Object $entity -Names @('id', 'Id'))
        if ([string]::IsNullOrWhiteSpace($entityId)) { continue }

        $assignments = @(Get-NCCoreAssignmentRecords -EntityType 'mobileApps' -EntityId $entityId -EffectiveGroupIds $effectiveGroupIds -RequestedGroupId $resolvedGroup.Id)
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
    Write-NCMessage "Intune profiles found for '$($resolvedGroup.DisplayName)': $($sorted.Count)" -Level VERBOSE

    if ($sorted.Count -eq 0) {
        Write-NCMessage "No Intune profiles found for group '$($resolvedGroup.DisplayName)'." -Level WARNING
        return
    }

    if ($GridView.IsPresent) {
        $sorted | Out-GridView -Title "Intune Profiles - $($resolvedGroup.DisplayName)"
    }
    else {
        $sorted
    }
}

function Invoke-NCGraphCollectionRequest {
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
                Write-Information "." -InformationAction Continue
            }

            $nextLink = $response.'@odata.nextLink'
            $nextUri = if ($nextLink) { [string]$nextLink } else { $null }
        }
        catch {
            if ($_.Exception.Message -like '*429*') {
                Write-Information "`nRate limit hit, waiting 60 seconds..." -InformationAction Continue
                Start-Sleep -Seconds 60
                continue
            }

            Write-Warning "Error fetching data: $($_.Exception.Message)"
            break
        }
    }

    return @($items)
}

function Invoke-NCGraphAllPagesCore {
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

            if ($response.value) {
                $allResults += $response.value
            }
            else {
                $allResults += $response
            }

            $nextLink = $response.'@odata.nextLink'
            if ($requestCount % 10 -eq 0) {
                Write-Information "." -InformationAction Continue
            }
        }
        catch {
            if ($_.Exception.Message -like "*429*") {
                Write-Information "`nRate limit hit, waiting 60 seconds..." -InformationAction Continue
                Start-Sleep -Seconds 60
                continue
            }

            Write-Warning "Error fetching data: $($_.Exception.Message)"
            break
        }
    }
    while ($nextLink)

    return $allResults
}

function Get-NCIntuneItemName {
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
