BeforeAll {
    function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
    function Add-EmptyLine {}
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Get-MgGroup { param([string]$GroupId, [string]$Filter, [switch]$All, [string]$ConsistencyLevel, [string]$CountVariable, [string]$ErrorAction) }
    function Invoke-MgGraphRequest {
        param(
            [string]$Uri,
            [string]$Method,
            [object]$Body,
            [string]$ContentType,
            [string]$OutputType
        )
    }

    function New-TestBatchResponse {
        param(
            [object]$Body,
            [scriptblock]$Responder
        )
        $payload = $Body | ConvertFrom-Json
        @{
            responses = @(foreach ($request in $payload.requests) {
                    $answer = & $Responder $request
                    @{ id = $request.id; status = $answer.status; headers = $answer.headers; body = $answer.body }
                })
        }
    }

    . "$PSScriptRoot/../../Private/NC-Hlp.Intune.ps1"
    . "$PSScriptRoot/../../Private/NC-Hlp.GraphBatch.ps1"
}

Describe 'Invoke-NCGraphAllPagesCore' {
    It 'returns an empty array when Graph reports zero items, not the raw response wrapper' {
        Mock Invoke-MgGraphRequest {
            [pscustomobject]@{
                '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#owners'
                value            = @()
            }
        }

        $result = @(Invoke-NCGraphAllPagesCore -Uri 'https://graph.microsoft.com/v1.0/applications/app-1/owners')

        $result.Count | Should -Be 0
    }

    It 'returns the actual items when Graph reports one or more' {
        Mock Invoke-MgGraphRequest {
            [pscustomobject]@{
                value = @(
                    [pscustomobject]@{ id = 'owner-1'; displayName = 'Jane Doe' }
                    [pscustomobject]@{ id = 'owner-2'; displayName = 'John Smith' }
                )
            }
        }

        $result = @(Invoke-NCGraphAllPagesCore -Uri 'https://graph.microsoft.com/v1.0/applications/app-1/owners')

        $result.Count | Should -Be 2
        $result[0].id | Should -Be 'owner-1'
        $result[1].id | Should -Be 'owner-2'
    }

    It 'still returns a single non-paged object as-is when the response has no value property' {
        Mock Invoke-MgGraphRequest {
            [pscustomobject]@{ id = 'single-object-id'; displayName = 'Not a collection' }
        }

        $result = @(Invoke-NCGraphAllPagesCore -Uri 'https://graph.microsoft.com/v1.0/applications/app-1')

        $result.Count | Should -Be 1
        $result[0].id | Should -Be 'single-object-id'
    }
}

Describe 'Resolve-NCIntuneManagedDeviceEntraMember batching' {
    BeforeEach {
        Mock Write-NCMessage {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                $global:SeenRequests.Add($request)
                $url = [string]$request.url
                if ($url -match '^/deviceManagement/managedDevices/(dev\d+)\?\$select=') {
                    return @{ status = 200; body = @{ id = $Matches[1]; azureADDeviceId = "az-$($Matches[1])" } }
                }
                if ($url -match '^/devices\?\$filter=(.+)$') {
                    $filter = [uri]::UnescapeDataString($Matches[1])
                    if ($filter -match "deviceId eq 'az-gone'") { return @{ status = 200; body = @{ value = @() } } }
                    $null = $filter -match "deviceId eq '([^']+)'"
                    return @{ status = 200; body = @{ value = @(@{ id = "ent-$($Matches[1])" }) } }
                }
                @{ status = 500 }
            }
        }
    }

    It 'resolves 45 devices with 3 lookup batches and keeps input order' {
        $devices = @(1..45 | ForEach-Object { [pscustomobject]@{ id = "dev$_"; deviceName = "PC$_"; azureADDeviceId = "az$_" } })
        $results = @(Resolve-NCIntuneManagedDeviceEntraMembers -ManagedDevices $devices -DeviceIds @($devices.id))

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
        $results.Count | Should -Be 45
        $results[0].DeviceId | Should -Be 'dev1'
        $results[0].Resolution.EntraDeviceId | Should -Be 'ent-az1'
        $results[44].Resolution.DeviceName | Should -Be 'PC45'
    }

    It 'refreshes missing Azure AD ids with a beta batch and warns for devices not in Entra' {
        $devices = @(
            [pscustomobject]@{ id = 'dev1'; deviceName = 'PC1' },
            [pscustomobject]@{ id = 'dev2'; deviceName = 'PC2'; azureADDeviceId = 'az-gone' }
        )
        $results = @(Resolve-NCIntuneManagedDeviceEntraMembers -ManagedDevices $devices -DeviceIds @('dev1', 'dev2', 'unknown'))

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        $results[0].Resolution.EntraDeviceId | Should -Be 'ent-az-dev1'
        $results[1].Resolution | Should -BeNullOrEmpty
        $results[2].Resolution | Should -BeNullOrEmpty
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -eq 'Device not found in Entra ID: PC2 (Azure AD Device ID: az-gone)' }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -eq 'No Azure AD Device ID for: unknown' }
    }

    It 'marks every unresolved device as a failed lookup, not only Graph errors' {
        $devices = @(
            [pscustomobject]@{ id = 'dev1'; deviceName = 'PC1'; azureADDeviceId = 'az1' },
            [pscustomobject]@{ id = 'dev2'; deviceName = 'PC2'; azureADDeviceId = 'az-gone' }
        )
        $results = @(Resolve-NCIntuneManagedDeviceEntraMembers -ManagedDevices $devices -DeviceIds @('dev1', 'dev2', 'unknown'))

        $results[0].LookupFailed | Should -BeFalse
        $results[1].LookupFailed | Should -BeTrue
        $results[2].LookupFailed | Should -BeTrue
    }

    It 'keeps the single-device wrapper working' {
        $devices = @([pscustomobject]@{ id = 'dev1'; deviceName = 'PC1'; azureADDeviceId = 'az1' })
        $resolution = Resolve-NCIntuneManagedDeviceEntraMember -ManagedDevices $devices -DeviceId 'dev1'

        $resolution.EntraDeviceId | Should -Be 'ent-az1'
        $resolution.IntuneDeviceId | Should -Be 'dev1'
    }
}

Describe 'Invoke-NCIntuneGroupUsageCore batching' {
    BeforeEach {
        Mock Write-NCMessage {}
        Mock Add-EmptyLine {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Test-MgGraphConnection { $true }
        Mock Get-MgGroup { [pscustomobject]@{ Id = 'root'; DisplayName = 'Root group' } }
        $script:NCIntuneGroupNameCache = @{}
        $global:SeenUris = [System.Collections.Generic.List[string]]::new()
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        Mock Invoke-MgGraphRequest {
            $global:SeenUris.Add($Uri)
            if ($Uri -like '*$batch') {
                return New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($url -match "^/deviceManagement/deviceConfigurations\('(cfg\d+)'\)/assignments$") {
                        return @{ status = 200; body = @{ value = @(@{ id = "as-$($Matches[1])"; target = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'root' } }) } }
                    }
                    if ($url -eq '/groups/root?$select=id,displayName') {
                        return @{ status = 200; body = @{ id = 'root'; displayName = 'Root group' } }
                    }
                    @{ status = 500 }
                }
            }
            if ($Uri -like '*deviceConfigurations') {
                return @{ value = @(1..45 | ForEach-Object { @{ id = "cfg$_"; displayName = "Profile $_" } }) }
            }
            @{ value = @() }
        }
    }

    It 'reads the assignments of 45 device configurations in 3 beta batches and resolves the group name once' {
        $results = @(Invoke-NCIntuneGroupUsageCore -ParameterSetName 'ById' -GroupId 'root' -Diagnostic)
        Should -Invoke Get-MgGroup -Times 1 -Exactly -Scope It
        @($global:SeenUris | Where-Object { $_ -like '*beta*$batch' }).Count | Should -Be 3
    }
}
