#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Intune ===============================================================================================================================

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

        function Invoke-NCGraphCollectionRequestLocal {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Uri
            )

            $items = [System.Collections.Generic.List[object]]::new()
            $nextUri = $Uri

            while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop

                $pageItems = @()
                if ($response -is [System.Collections.IDictionary]) {
                    if ($response.Contains('value')) {
                        $value = $response['value']
                        if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                            foreach ($entry in $value) {
                                if ($null -ne $entry) { $pageItems += $entry }
                            }
                        }
                        elseif ($null -ne $value) {
                            $pageItems = @($value)
                        }
                    }
                    elseif (($response.Contains('id') -or $response.Contains('Id')) -and $null -ne $response) {
                        $pageItems = @($response)
                    }
                }
                elseif ($response -is [System.Array]) {
                    $pageItems = @($response)
                }
                elseif ($response.PSObject.Properties.Name -contains 'value') {
                    $value = $response.value
                    if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                        foreach ($entry in $value) {
                            if ($null -ne $entry) { $pageItems += $entry }
                        }
                    }
                    elseif ($null -ne $value) {
                        $pageItems = @($value)
                    }
                }
                elseif ((($response.PSObject.Properties.Name -contains 'id') -or ($response.PSObject.Properties.Name -contains 'Id')) -and $null -ne $response) {
                    $pageItems = @($response)
                }

                foreach ($item in $pageItems) {
                    $items.Add($item) | Out-Null
                }

                $nextLink = $null
                if ($response -is [System.Collections.IDictionary]) {
                    if ($response.Contains('@odata.nextLink')) {
                        $nextLink = [string]$response['@odata.nextLink']
                    }
                }
                else {
                    $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
                    if ($nextLinkProperty) {
                        $nextLink = [string]$nextLinkProperty.Value
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($nextLink)) {
                    $nextUri = $nextLink
                }
                else {
                    $nextUri = $null
                }
            }

            return $items.ToArray()
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
                    $items = @(Invoke-NCGraphCollectionRequestLocal -Uri $endpointUri)
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

function Search-IntuneRawEndpoint {
    <#
    .SYNOPSIS
        Probes raw Intune Microsoft Graph endpoints and filters results by profile name.
    .DESCRIPTION
        Connects to Microsoft Graph, queries a curated set of Intune deviceManagement endpoints,
        and returns any items whose common identifying properties or object ID match the provided search text.
        Use this as a low-level discovery tool when a profile is not found by the higher-level
        Intune helper functions.
    .PARAMETER SearchText
        Text to search for in the returned profile names.
    .PARAMETER Endpoint
        Optional custom endpoint URIs to probe. When omitted, a built-in curated list is used.
    .PARAMETER Exact
        Match the profile name exactly instead of using a contains search.
    .PARAMETER GridView
        Show the results in Out-GridView instead of returning objects.
    .EXAMPLE
        Search-IntuneRawEndpoint -SearchText "iOS - Wi-Fi M-Smartphone"
    .EXAMPLE
        Search-IntuneRawEndpoint -SearchText "Wi-Fi" -GridView
    .EXAMPLE
        Search-IntuneRawEndpoint -SearchText "Wi-Fi" -Endpoint 'beta/deviceManagement/configurationPolicies?$top=100'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'DisplayName', 'ProfileName', 'Query')]
        [string]$SearchText,
        [string[]]$Endpoint,
        [switch]$Exact,
        [switch]$GridView
    )

    begin {
        $graphConnected = $null

        function Invoke-NCRawGraphCollectionRequest {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Uri
            )

            $items = [System.Collections.Generic.List[object]]::new()
            $nextUri = $Uri

            while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop

                $pageItems = @()
                if ($response -is [System.Collections.IDictionary]) {
                    if ($response.Contains('value')) {
                        $value = $response['value']
                        if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                            foreach ($entry in $value) {
                                if ($null -ne $entry) { $pageItems += $entry }
                            }
                        }
                        elseif ($null -ne $value) {
                            $pageItems = @($value)
                        }
                    }
                    elseif (($response.Contains('id') -or $response.Contains('Id')) -and $null -ne $response) {
                        $pageItems = @($response)
                    }
                }
                elseif ($response -is [System.Array]) {
                    $pageItems = @($response)
                }
                elseif ($response.PSObject.Properties.Name -contains 'value') {
                    $value = $response.value
                    if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                        foreach ($entry in $value) {
                            if ($null -ne $entry) { $pageItems += $entry }
                        }
                    }
                    elseif ($null -ne $value) {
                        $pageItems = @($value)
                    }
                }
                elseif ((($response.PSObject.Properties.Name -contains 'id') -or ($response.PSObject.Properties.Name -contains 'Id')) -and $null -ne $response) {
                    $pageItems = @($response)
                }

                foreach ($item in $pageItems) {
                    $items.Add($item) | Out-Null
                }

                $nextLink = $null
                if ($response -is [System.Collections.IDictionary]) {
                    if ($response.Contains('@odata.nextLink')) {
                        $nextLink = [string]$response['@odata.nextLink']
                    }
                }
                else {
                    $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
                    if ($nextLinkProperty) {
                        $nextLink = [string]$nextLinkProperty.Value
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($nextLink)) {
                    $nextUri = $nextLink
                }
                else {
                    $nextUri = $null
                }
            }

            return $items.ToArray()
        }

        function Get-NCRawItemName {
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

        function Get-NCRawItemId {
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

        function Get-NCRawItemODataType {
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

        function Get-NCRawSearchFields {
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

        $endpointsToProbe = if ($Endpoint -and $Endpoint.Count -gt 0) {
            @($Endpoint)
        }
        else {
            @(
                'v1.0/deviceManagement/deviceConfigurations?$top=100',
                'beta/deviceManagement/deviceConfigurations?$top=100',
                'v1.0/deviceManagement/configurationPolicies?$top=100',
                'beta/deviceManagement/configurationPolicies?$top=100',
                'beta/deviceManagement/groupPolicyConfigurations?$top=100',
                'beta/deviceManagement/resourceAccessProfiles?$top=100',
                'v1.0/deviceManagement/deviceCompliancePolicies?$top=100',
                'v1.0/deviceManagement/deviceEnrollmentConfigurations?$top=100',
                'beta/deviceManagement/deviceHealthScripts?$top=100',
                'beta/deviceManagement/deviceManagementScripts?$top=100',
                'beta/deviceManagement/deviceShellScripts?$top=100',
                'beta/deviceManagement/intent?$top=100',
                'beta/deviceManagement/templates?$top=100',
                'beta/deviceAppManagement/mobileApps?$top=100'
            )
        }

        $normalizedSearch = $SearchText.Trim()
        $results = [System.Collections.Generic.List[object]]::new()

        foreach ($endpointUri in $endpointsToProbe) {
            $items = @()
            try {
                $items = @(Invoke-NCRawGraphCollectionRequest -Uri $endpointUri)
            }
            catch {
                Write-NCMessage "Unable to query $endpointUri : $($_.Exception.Message)" -Level WARNING
                continue
            }

            foreach ($item in $items) {
                $itemName = Get-NCRawItemName -Item $item
                $searchFields = @(Get-NCRawSearchFields -Item $item)
                if ($searchFields.Count -eq 0) {
                    continue
                }

                $matchedField = $null
                foreach ($field in $searchFields) {
                    $isMatch = if ($Exact.IsPresent) {
                        $field.Value -eq $normalizedSearch
                    }
                    else {
                        $field.Value -like "*$normalizedSearch*"
                    }

                    if ($isMatch) {
                        $matchedField = $field
                        break
                    }
                }

                if (-not $matchedField) {
                    continue
                }

                $results.Add([pscustomobject][ordered]@{
                        'Profile Name'      = $itemName
                        'Endpoint'          = $endpointUri
                        'Matched Property'  = $matchedField.Property
                        'Matched Value'     = $matchedField.Value
                        'Profile Id'        = Get-NCRawItemId -Item $item
                        'Profile Type'      = Get-NCRawItemODataType -Item $item
                    }) | Out-Null
            }
        }

        Add-EmptyLine
        Write-NCMessage "Raw Intune matches found for '$normalizedSearch': $($results.Count)" -Level VERBOSE

        if ($results.Count -eq 0) {
            Write-NCMessage "No raw Intune matches found for '$normalizedSearch' in the probed endpoints." -Level WARNING
            return
        }

        $sorted = $results | Sort-Object 'Profile Name', 'Endpoint' -Unique
        if ($GridView.IsPresent) {
            $sorted | Out-GridView -Title "Intune Raw Endpoint Search - $normalizedSearch"
        }
        else {
            $sorted
        }
    }
}

function Search-IntuneObjectById {
    <#
    .SYNOPSIS
        Probes raw Intune Microsoft Graph object endpoints by ID.
    .DESCRIPTION
        Connects to Microsoft Graph and attempts direct GET requests against a curated set of
        Intune deviceManagement object endpoints using the provided ID. Use this when an object
        does not appear in collection listings but you have an identifier from the Intune portal.
    .PARAMETER Id
        Object ID to probe across Intune Graph endpoints.
    .PARAMETER Endpoint
        Optional custom object endpoint prefixes. When omitted, a built-in curated list is used.
    .PARAMETER GridView
        Show the results in Out-GridView instead of returning objects.
    .EXAMPLE
        Search-IntuneObjectById -Id "cbbf1b23-3a92-47d6-974b-9d8295b9978d"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('ObjectId', 'ProfileId')]
        [string]$Id,
        [string[]]$Endpoint,
        [switch]$GridView
    )

    process {
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

        if ([string]::IsNullOrWhiteSpace($Id)) {
            Write-NCMessage "Id cannot be empty." -Level WARNING
            return
        }

        $endpointPrefixes = if ($Endpoint -and $Endpoint.Count -gt 0) {
            @($Endpoint)
        }
        else {
            @(
                'v1.0/deviceManagement/deviceConfigurations',
                'beta/deviceManagement/deviceConfigurations',
                'v1.0/deviceManagement/configurationPolicies',
                'beta/deviceManagement/configurationPolicies',
                'beta/deviceManagement/groupPolicyConfigurations',
                'beta/deviceManagement/resourceAccessProfiles',
                'v1.0/deviceManagement/deviceCompliancePolicies',
                'v1.0/deviceManagement/deviceEnrollmentConfigurations',
                'beta/deviceManagement/deviceHealthScripts',
                'beta/deviceManagement/deviceManagementScripts',
                'beta/deviceManagement/deviceShellScripts',
                'beta/deviceManagement/intents',
                'beta/deviceManagement/templates',
                'beta/deviceAppManagement/mobileApps'
            )
        }

        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($prefix in $endpointPrefixes) {
            $uri = "$prefix/$Id"
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            }
            catch {
                continue
            }

            $name = $null
            foreach ($propertyName in @('displayName', 'DisplayName', 'name', 'Name')) {
                $property = $response.PSObject.Properties[$propertyName]
                if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                    $name = [string]$property.Value
                    break
                }
            }

            $odataType = $null
            $odataProperty = $response.PSObject.Properties['@odata.type']
            if ($odataProperty -and -not [string]::IsNullOrWhiteSpace([string]$odataProperty.Value)) {
                $odataType = [string]$odataProperty.Value
            }
            else {
                $additionalProperties = $response.PSObject.Properties['AdditionalProperties']
                if ($additionalProperties -and $additionalProperties.Value -and $additionalProperties.Value.ContainsKey('@odata.type')) {
                    $odataType = [string]$additionalProperties.Value['@odata.type']
                }
            }

            $results.Add([pscustomobject][ordered]@{
                    'Endpoint'     = $prefix
                    'Profile Id'   = $Id
                    'Profile Name' = $name
                    'Profile Type' = $odataType
                }) | Out-Null
        }

        Add-EmptyLine
        Write-NCMessage "Direct Intune endpoint matches found for '$Id': $($results.Count)" -Level VERBOSE

        if ($results.Count -eq 0) {
            Write-NCMessage "No direct Intune endpoint matches found for '$Id' in the probed endpoints." -Level WARNING
            return
        }

        $sorted = $results | Sort-Object 'Endpoint' -Unique
        if ($GridView.IsPresent) {
            $sorted | Out-GridView -Title "Intune Direct Object Search - $Id"
        }
        else {
            $sorted
        }
    }
}

function Get-IntuneProfileAssignmentsRaw {
    <#
    .SYNOPSIS
        Returns raw Intune profile assignments for a specific profile ID.
    .DESCRIPTION
        Probes multiple Microsoft Graph assignment endpoints for the provided profile ID and returns
        the raw assignment target details. Use this to understand how Intune exposes assignment data
        for profiles that do not behave as expected through higher-level helper functions.
    .PARAMETER ProfileId
        Intune profile ID to inspect.
    .PARAMETER GridView
        Show the results in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-IntuneProfileAssignmentsRaw -ProfileId "00000000-0000-0000-0000-000000000000"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Id', 'ObjectId')]
        [string]$ProfileId,
        [switch]$GridView
    )

    begin {
        function Invoke-NCRawAssignmentCollectionRequest {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$Uri
            )

            $items = [System.Collections.Generic.List[object]]::new()
            $nextUri = $Uri

            while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop

                $pageItems = @()
                if ($response -is [System.Collections.IDictionary]) {
                    if ($response.Contains('value')) {
                        $value = $response['value']
                        if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                            foreach ($entry in $value) {
                                if ($null -ne $entry) { $pageItems += $entry }
                            }
                        }
                        elseif ($null -ne $value) {
                            $pageItems = @($value)
                        }
                    }
                    elseif (($response.Contains('id') -or $response.Contains('Id')) -and $null -ne $response) {
                        $pageItems = @($response)
                    }
                }
                elseif ($response -is [System.Array]) {
                    $pageItems = @($response)
                }
                elseif ($response.PSObject.Properties.Name -contains 'value') {
                    $value = $response.value
                    if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                        foreach ($entry in $value) {
                            if ($null -ne $entry) { $pageItems += $entry }
                        }
                    }
                    elseif ($null -ne $value) {
                        $pageItems = @($value)
                    }
                }
                elseif ((($response.PSObject.Properties.Name -contains 'id') -or ($response.PSObject.Properties.Name -contains 'Id')) -and $null -ne $response) {
                    $pageItems = @($response)
                }

                foreach ($item in $pageItems) {
                    $items.Add($item) | Out-Null
                }

                $nextLink = $null
                if ($response -is [System.Collections.IDictionary]) {
                    if ($response.Contains('@odata.nextLink')) {
                        $nextLink = [string]$response['@odata.nextLink']
                    }
                }
                else {
                    $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
                    if ($nextLinkProperty) {
                        $nextLink = [string]$nextLinkProperty.Value
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($nextLink)) {
                    $nextUri = $nextLink
                }
                else {
                    $nextUri = $null
                }
            }

            return $items.ToArray()
        }
    }

    process {
        $graphConnected = Test-MgGraphConnection -Scopes @('DeviceManagementConfiguration.Read.All', 'Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
            return
        }

        $assignmentEndpoints = @(
            "beta/deviceManagement/deviceConfigurations('$ProfileId')/assignments?`$top=100",
            "v1.0/deviceManagement/deviceConfigurations('$ProfileId')/assignments?`$top=100",
            "beta/deviceManagement/configurationPolicies('$ProfileId')/assignments?`$top=100",
            "beta/deviceManagement/groupPolicyConfigurations('$ProfileId')/assignments?`$top=100",
            "beta/deviceManagement/resourceAccessProfiles('$ProfileId')/assignments?`$top=100"
        )

        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($endpoint in $assignmentEndpoints) {
            $assignments = @()
            try {
                $assignments = @(Invoke-NCRawAssignmentCollectionRequest -Uri $endpoint)
            }
            catch {
                continue
            }

            foreach ($assignment in $assignments) {
                $target = $null
                if ($assignment -is [System.Collections.IDictionary]) {
                    foreach ($key in @('target', 'Target')) {
                        if ($assignment.Contains($key) -and $null -ne $assignment[$key]) {
                            $target = $assignment[$key]
                            break
                        }
                    }
                }
                elseif ($assignment.PSObject.Properties['Target']) {
                    $target = $assignment.Target
                }
                elseif ($assignment.PSObject.Properties['target']) {
                    $target = $assignment.PSObject.Properties['target'].Value
                }

                $targetProps = @{}
                if ($target) {
                    foreach ($prop in $target.PSObject.Properties) {
                        if ($prop.Name -ne 'AdditionalProperties') {
                            $targetProps[$prop.Name] = $prop.Value
                        }
                    }

                    $additionalTargetProperties = $target.PSObject.Properties['AdditionalProperties']
                    if ($additionalTargetProperties -and $additionalTargetProperties.Value) {
                        foreach ($key in $additionalTargetProperties.Value.Keys) {
                            $targetProps[$key] = $additionalTargetProperties.Value[$key]
                        }
                    }
                }

                $assignmentId = $null
                if ($assignment -is [System.Collections.IDictionary]) {
                    foreach ($key in @('id', 'Id')) {
                        if ($assignment.Contains($key) -and -not [string]::IsNullOrWhiteSpace([string]$assignment[$key])) {
                            $assignmentId = [string]$assignment[$key]
                            break
                        }
                    }
                }
                elseif ($assignment.PSObject.Properties['Id']) {
                    $assignmentId = $assignment.Id
                }
                elseif ($assignment.PSObject.Properties['id']) {
                    $assignmentId = $assignment.PSObject.Properties['id'].Value
                }

                $results.Add([pscustomobject][ordered]@{
                        'Endpoint'         = $endpoint
                        'Assignment Id'    = $assignmentId
                        'Target Summary'   = ($targetProps.GetEnumerator() | Sort-Object Name | ForEach-Object { "{0}={1}" -f $_.Name, $_.Value }) -join '; '
                        'Target Properties' = [pscustomobject]$targetProps
                    }) | Out-Null
            }
        }

        Add-EmptyLine
        Write-NCMessage "Raw assignment rows found for '$ProfileId': $($results.Count)" -Level VERBOSE

        if ($results.Count -eq 0) {
            Write-NCMessage "No raw assignments found for '$ProfileId' in the probed endpoints." -Level WARNING
            return
        }

        $sorted = $results | Sort-Object 'Endpoint', 'Assignment Id' -Unique
        if ($GridView.IsPresent) {
            $sorted | Out-GridView -Title "Intune Raw Assignments - $ProfileId"
        }
        else {
            $sorted
        }
    }
}
