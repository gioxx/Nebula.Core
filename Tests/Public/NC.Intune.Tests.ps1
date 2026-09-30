BeforeAll {
    $global:NCVars = @{ DateTimeString_Full = 'yyyy-MM-dd HH:mm:ss'; CSV_DefaultLimiter = ',' }
    function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
    function Add-EmptyLine {}
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Invoke-MgGraphRequest {
        param(
            [string]$Uri,
            [string]$Method,
            [object]$Body,
            [string]$ContentType,
            [string]$OutputType
        )
    }
    function Get-MgContext {}

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

    . "$PSScriptRoot/../../Private/NC-Hlp.ModuleUtils.ps1"
    function Format-NCDateTime { param($Value, [switch]$AsLocalTime) "FMT:$Value" }
    . "$PSScriptRoot/../../Private/NC-Hlp.Intune.ps1"
    . "$PSScriptRoot/../../Private/NC-Hlp.GraphBatch.ps1"
    . "$PSScriptRoot/../../Public/NC.Intune.ps1"
}

Describe 'Export-IntuneAppInventory batching' {
    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Write-NCMessage {}
        Mock Add-EmptyLine {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        $global:SeenUris = [System.Collections.Generic.List[string]]::new()
        $global:FailDevice = ''
        $global:JsonPath = Join-Path $TestDrive 'out.json'
    }

    BeforeAll {
        function Set-IntuneDevices {
            param([int]$Count, [switch]$WithLastSync)
            $global:TestDevices = @($(if ($Count -gt 0) { 1..$Count }) | ForEach-Object {
                    $d = [pscustomobject]@{ id = "dev$_"; deviceName = "PC$_"; operatingSystem = 'Windows'; userPrincipalName = "user$_@contoso.com"; lastSyncDateTime = $null }
                    if ($WithLastSync) { $d.lastSyncDateTime = '2026-01-02T03:04:05Z' }
                    $d
                })
            Mock Invoke-NCGraphAllPagesCore { $global:TestDevices }
        }

        function Set-IntuneGraphMock {
            Mock Invoke-MgGraphRequest {
                $global:SeenUris.Add($Uri)
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($url -match '^/deviceManagement/managedDevices/(dev\d+)\?\$expand=detectedApps$') {
                        if ($Matches[1] -eq $global:FailDevice) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 200; body = @{ detectedApps = @(
                                    @{ displayName = 'Java 8'; version = '8.0.1'; publisher = 'Oracle' },
                                    @{ displayName = 'Other'; version = '1.0'; publisher = 'X' }
                                ) } }
                    }
                    if ($url -match '^/deviceManagement/managedDevices/(dev\d+)\?\$select=lastSyncDateTime$') {
                        return @{ status = 200; body = @{ lastSyncDateTime = '2026-03-04T05:06:07Z' } }
                    }
                    @{ status = 500 }
                }
            }
        }
    }

    It 'reads detected apps for 45 devices in 3 beta batch calls and keeps row order' {
        Set-IntuneDevices -Count 45
        Set-IntuneGraphMock
        Export-IntuneAppInventory -ApplicationName 'Java*' -OutputJsonPath $global:JsonPath | Out-Null

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
        @($global:SeenUris | Where-Object { $_ -like '*beta*$batch' }).Count | Should -Be 3
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing 45 device(s) in Graph batches (20 per request) ...' }

        $rows = @(Get-Content -LiteralPath $global:JsonPath -Raw | ConvertFrom-Json)
        $rows.Count | Should -Be 45
        $rows[0].AppName | Should -Be 'Java 8'
        $rows[0].Version | Should -Be '8.0.1'
        $rows[0].Publisher | Should -Be 'Oracle'
        $rows[0].Source | Should -Be 'DetectedApps'
        @($rows.DeviceId) | Should -Be @(1..45 | ForEach-Object { "dev$_" })
    }

    It 'skips a failing device with the existing message and processes the others' {
        Set-IntuneDevices -Count 25
        Set-IntuneGraphMock
        $global:FailDevice = 'dev3'
        Export-IntuneAppInventory -ApplicationName 'Java*' -OutputJsonPath $global:JsonPath | Out-Null

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -like 'Error reading apps for PC3: *denied*' }
        $rows = @(Get-Content -LiteralPath $global:JsonPath -Raw | ConvertFrom-Json)
        $rows.Count | Should -Be 24
        $rows.DeviceId | Should -Not -Contain 'dev3'
    }

    It 'reads missing last inventory dates in v1.0 batches of 20' {
        Set-IntuneDevices -Count 45
        Set-IntuneGraphMock
        Export-IntuneAppInventory -ApplicationName 'Java*' -LastInventory -OutputJsonPath $global:JsonPath | Out-Null

        Should -Invoke Invoke-MgGraphRequest -Times 6 -Exactly -Scope It
        @($global:SeenUris | Where-Object { $_ -like '*v1.0*$batch' }).Count | Should -Be 3
        $rows = @(Get-Content -LiteralPath $global:JsonPath -Raw | ConvertFrom-Json)
        $rows.Count | Should -Be 45
        $rows[0].LastInventory | Should -Be 'FMT:2026-03-04T05:06:07Z'
    }

    It 'uses the last sync date from the device list without extra v1.0 calls' {
        Set-IntuneDevices -Count 45 -WithLastSync
        Set-IntuneGraphMock
        Export-IntuneAppInventory -ApplicationName 'Java*' -LastInventory -OutputJsonPath $global:JsonPath | Out-Null

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
        $rows = @(Get-Content -LiteralPath $global:JsonPath -Raw | ConvertFrom-Json)
        $rows[0].LastInventory | Should -Be 'FMT:2026-01-02T03:04:05Z'
    }

    It 'writes no start line and makes no batch call when there are no devices' {
        Set-IntuneDevices -Count 0
        Set-IntuneGraphMock
        Export-IntuneAppInventory -ApplicationName 'Java*' | Out-Null

        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Message -like 'Processing*' }
    }
}

Describe 'New-IntuneAppBasedGroup batching' {
    BeforeAll {
        function Get-MgGroup { param([string]$Filter, [switch]$All, [string]$ErrorAction) }

        function Set-AppGroupDevices {
            param([int]$Count)
            $global:TestDevices = @($(if ($Count -gt 0) { 1..$Count }) | ForEach-Object {
                    [pscustomobject]@{ id = "dev$_"; deviceName = "PC$_"; operatingSystem = 'Windows'; userPrincipalName = "u$_@contoso.com"; azureADDeviceId = "az$_" }
                })
        }

        function Set-AppGroupGraphMock {
            Mock Invoke-NCGraphAllPagesCore {
                if ($Uri -like '*/managedDevices?*') { return $global:TestDevices }
                if ($Uri -like '*/groups/G1/members') { return $global:CurrentMembers }
                @()
            }
            Mock Invoke-MgGraphRequest {
                $global:SeenUris.Add($Uri)
                if ($Uri -like '*$batch') {
                    return New-TestBatchResponse -Body $Body -Responder {
                        param($request)
                        $global:SeenRequests.Add($request)
                        $url = [string]$request.url
                        $method = [string]$request.method
                        if ($method -eq 'GET' -and $url -match '^/deviceManagement/managedDevices/(dev\d+)\?\$select=.*&\$expand=detectedApps$') {
                            return @{ status = 200; body = @{ detectedApps = @(@{ displayName = 'Java 8'; version = '8.0.1'; publisher = 'Oracle' }) } }
                        }
                        if ($method -eq 'GET' -and $url -match '^/devices\?\$filter=(.+)$') {
                            $filter = [uri]::UnescapeDataString($Matches[1])
                            $null = $filter -match "deviceId eq 'az(\d+)'"
                            return @{ status = 200; body = @{ value = @(@{ id = "ent$($Matches[1])" }) } }
                        }
                        if ($method -eq 'GET' -and $url -match '^/groups\?\$filter=') {
                            if ($global:ExistingGroup) { return @{ status = 200; body = @{ value = @(@{ id = 'G1'; displayName = 'Devices - Java' }) } } }
                            return @{ status = 200; body = @{ value = @() } }
                        }
                        if ($method -eq 'POST' -and ($url -eq '/groups/G1/members/$ref' -or $url -eq '/groups/NEW1/members/$ref')) {
                            $odataId = [string]$request.body.'@odata.id'
                            if ($odataId -like "*/$($global:ExistsEntraId)") {
                                return @{ status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'One or more added object references already exist for the following modified properties: members.' } } }
                            }
                            return @{ status = 204 }
                        }
                        if ($method -eq 'DELETE' -and $url -match '^/groups/G1/members/([^/]+)/\$ref$') {
                            return @{ status = 204 }
                        }
                        @{ status = 500; body = @{ error = @{ code = 'x'; message = "unexpected $method $url" } } }
                    }
                }
                if ($Uri -eq 'https://graph.microsoft.com/v1.0/groups' -and $Method -eq 'POST') {
                    return @{ id = 'NEW1' }
                }
                throw "unexpected direct call $Method $Uri"
            }
        }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Write-NCMessage {}
        Mock Add-EmptyLine {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Get-MgGroup { throw 'Get-MgGroup must not be called' }
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        $global:SeenUris = [System.Collections.Generic.List[string]]::new()
        $global:ExistingGroup = $false
        $global:ExistsEntraId = ''
        $global:CurrentMembers = @()
    }

    It 'creates a group and adds 25 devices with 2 batches of POST members/$ref, reading detected apps and resolving devices in batches' {
        Set-AppGroupDevices -Count 25
        Set-AppGroupGraphMock
        New-IntuneAppBasedGroup -ApplicationName 'Java*' -GroupName 'Devices - Java' -Confirm:$false

        # 2 detectedApps + 2 device lookups + 1 group lookup + 1 create (direct) + 2 member adds
        Should -Invoke Invoke-MgGraphRequest -Times 8 -Exactly -Scope It
        @($global:SeenUris | Where-Object { $_ -like '*beta*$batch' }).Count | Should -Be 2
        $adds = @($global:SeenRequests | Where-Object { $_.method -eq 'POST' })
        $adds.Count | Should -Be 25
        @($adds | ForEach-Object { $_.url } | Select-Object -Unique) | Should -Be @('/groups/NEW1/members/$ref')
        $adds[0].body.'@odata.id' | Should -Be 'https://graph.microsoft.com/v1.0/directoryObjects/ent1'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing 25 device(s) in Graph batches (20 per request) ...' }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Added 25 devices to group' -and $Level -eq 'SUCCESS' }
        Should -Invoke Get-MgGroup -Times 0 -Exactly -Scope It
    }

    It 'reports a device that is already a member and still adds the others' {
        Set-AppGroupDevices -Count 25
        Set-AppGroupGraphMock
        $global:ExistsEntraId = 'ent7'
        New-IntuneAppBasedGroup -ApplicationName 'Java*' -GroupName 'Devices - Java' -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -eq "Device 'PC7' is already a member of 'Devices - Java'" }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Added 24 devices to group' -and $Level -eq 'SUCCESS' }
    }

    It 'updates an existing group with batched adds and removes' {
        Set-AppGroupDevices -Count 25
        Set-AppGroupGraphMock
        $global:ExistingGroup = $true
        $global:CurrentMembers = @([pscustomobject]@{ id = 'old1'; displayName = 'Old 1' }, [pscustomobject]@{ id = 'ent1'; displayName = 'PC1' })
        New-IntuneAppBasedGroup -ApplicationName 'Java*' -GroupName 'Devices - Java' -UpdateExisting -Confirm:$false

        @($global:SeenRequests | Where-Object { $_.method -eq 'POST' }).Count | Should -Be 24
        @($global:SeenRequests | Where-Object { $_.method -eq 'DELETE' }).Count | Should -Be 1
        @($global:SeenRequests | Where-Object { $_.method -eq 'DELETE' })[0].url | Should -Be '/groups/G1/members/old1/$ref'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Updated group: Devices - Java (Added: 24, Removed: 1)' -and $Level -eq 'SUCCESS' }
    }

    It 'sends no write requests with -WhatIf' {
        Set-AppGroupDevices -Count 3
        Set-AppGroupGraphMock
        New-IntuneAppBasedGroup -ApplicationName 'Java*' -GroupName 'Devices - Java' -WhatIf

        @($global:SeenRequests | Where-Object { $_.method -in 'POST', 'DELETE' }).Count | Should -Be 0
    }

    It 'writes no start line and makes no calls when there are no devices' {
        Set-AppGroupDevices -Count 0
        Set-AppGroupGraphMock
        New-IntuneAppBasedGroup -ApplicationName 'Java*' -GroupName 'Devices - Java' -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Message -like 'Processing*' }
    }
}
