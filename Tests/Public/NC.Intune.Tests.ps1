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
