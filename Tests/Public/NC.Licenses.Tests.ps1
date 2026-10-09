BeforeAll {
    function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
    function Add-EmptyLine {}
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Set-ProgressAndInfoPreferences {}
    function Restore-ProgressAndInfoPreferences {}
    function Get-LicenseCatalog { param([switch]$IncludeMetadata, [switch]$ForceRefresh) }
    function Get-LicenseDisplayName { param($Lookup, $SkuPartNumber, $FallbackLookup, $MatchSource) }
    function Get-MgSubscribedSku { param([switch]$All) }
    function Find-UserRecipient { param([string]$UserPrincipalName, [switch]$PreferGraphIdentity, [switch]$SkipDirectGraphLookup) }
    function Invoke-NCRetry {
        param([scriptblock]$Action, [int]$MaxAttempts, [int]$DelaySeconds, [string]$OperationDescription, [scriptblock]$OnError)
        & $Action
    }
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
    function Get-NCProgressPercent { param($Current, $Total) if ($Total) { [int](100 * $Current / $Total) } else { 0 } }
    function Test-Folder { param($Path) $Path }
    function New-File { param($Path) $Path }
    function Get-MgUser { [CmdletBinding()] param($Filter, $ConsistencyLevel, $CountVariable, [switch]$All, $Property) }
    function Get-MgEnvironment { param([string]$Name) }
    function Update-MgUser { [CmdletBinding()] param($UserId, $UsageLocation) }
    function Set-MgUserLicense { [CmdletBinding()] param($UserId, $AddLicenses, $RemoveLicenses) }

    # Builds a $batch response by asking $Responder for each sub-request ({ param($request) @{ status; body } }).
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
    . "$PSScriptRoot/../../Public/NC.Licenses.ps1"

    $global:NCVars = @{ UsageLocation = 'IT' }
    $global:skuId = '6fd2c87f-b296-42f0-b197-1e91e994b900'
}

Describe 'License assignment batching' {
    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Write-NCMessage {}
        Mock Add-EmptyLine {}
        Mock Set-ProgressAndInfoPreferences {}
        Mock Restore-ProgressAndInfoPreferences {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Find-UserRecipient {}
        Mock Get-LicenseCatalog { $null }
        Mock Get-MgSubscribedSku {
            @([pscustomobject]@{ SkuId = [guid]$global:skuId; SkuPartNumber = 'ENTERPRISEPACK'; PrepaidUnits = [pscustomobject]@{ Enabled = 100 }; ConsumedUnits = 0 })
        }
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        $global:UsageLocationForUsers = 'IT'
        $global:PatchFailsFor = ''
        $global:UsageLocationOverride = @{}
        $global:AssignedFor = @()
    }

    BeforeAll {
        function New-Upns { param([int]$Count) 1..$Count | ForEach-Object { "user$_@contoso.com" } }

        function Set-LicenseGraphMock {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($request.method -eq 'GET' -and $url -match '^/users/user(\d+)%40contoso\.com') {
                        $n = $Matches[1]
                        $loc = if ($global:UsageLocationOverride.ContainsKey($n)) { $global:UsageLocationOverride[$n] } elseif ($global:UsageLocationForUsers) { $global:UsageLocationForUsers } else { $null }
                        $owned = if ($global:AssignedFor -contains $n) { @(@{ skuId = $global:skuId; disabledPlans = @() }) } else { @() }
                        return @{ status = 200; body = @{ id = "id$n"; userPrincipalName = "user$n@contoso.com"; displayName = "User $n"; usageLocation = $loc; assignedLicenses = $owned } }
                    }
                    if ($request.method -eq 'GET' -and $url -match '^/users/(id\d+)/licenseDetails') {
                        return @{ status = 200; body = @{ value = @(@{ skuId = $global:skuId; skuPartNumber = 'ENTERPRISEPACK' }) } }
                    }
                    if ($request.method -eq 'PATCH' -and $url -match '^/users/(id\d+)$') {
                        if ($Matches[1] -eq $global:PatchFailsFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 204 }
                    }
                    if ($request.method -eq 'POST' -and $url -match '^/users/(id\d+)/assignLicense$') {
                        return @{ status = 200; body = @{ id = $Matches[1] } }
                    }
                    @{ status = 500 }
                }
            }
        }
    }

    It 'reads tenant data once and uses 2 Graph calls for 14 users with usage location set' {
        Set-LicenseGraphMock
        New-Upns 14 | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        Should -Invoke Get-MgSubscribedSku -Times 1 -Exactly -Scope It
        Should -Invoke Get-LicenseCatalog -Times 1 -Exactly -Scope It
        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 14 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' }
    }

    It 'adds a PATCH pass (3 calls) when usage location is empty' {
        $global:UsageLocationForUsers = ''
        Set-LicenseGraphMock
        New-Upns 14 | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
        $patches = @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' })
        $patches.Count | Should -Be 14
        ($patches | ForEach-Object { $_.body.usageLocation } | Select-Object -Unique) | Should -Be 'IT'
        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 14
        @($assigns[0].body.addLicenses)[0].skuId | Should -Be $global:skuId
    }

    It 'skips assignLicense for a user whose usage location PATCH failed' {
        $global:UsageLocationForUsers = ''
        $global:PatchFailsFor = 'id2'
        Set-LicenseGraphMock
        @('user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com') | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 2
        ($assigns.url -contains '/users/id2/assignLicense') | Should -BeFalse
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -eq 'Unable to set usage location (IT) for user2@contoso.com: denied'
        }
    }

    It 'reports no availability for the second user when only one unit is left' {
        Mock Get-MgSubscribedSku {
            @([pscustomobject]@{ SkuId = [guid]$global:skuId; SkuPartNumber = 'ENTERPRISEPACK'; PrepaidUnits = [pscustomobject]@{ Enabled = 1 }; ConsumedUnits = 0 })
        }
        Set-LicenseGraphMock
        @('user1@contoso.com', 'user2@contoso.com') | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 1
        $assigns[0].url | Should -Be '/users/id1/assignLicense'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'WARNING' -and $Message -eq 'No available units for license ENTERPRISEPACK (ENTERPRISEPACK) (available: 0)'
        }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -eq 'No licenses to assign: none available. Requested: ENTERPRISEPACK'
        }
    }

    It 'neither reserves a seat for nor reassigns a license the user already has' {
        Mock Get-MgSubscribedSku {
            @([pscustomobject]@{ SkuId = [guid]$global:skuId; SkuPartNumber = 'ENTERPRISEPACK'; PrepaidUnits = [pscustomobject]@{ Enabled = 1 }; ConsumedUnits = 0 })
        }
        $global:AssignedFor = @('1')
        Set-LicenseGraphMock
        @('user1@contoso.com', 'user2@contoso.com') | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 1
        $assigns[0].url | Should -Be '/users/id2/assignLicense'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'WARNING' -and $Message -eq 'user1@contoso.com already has license(s): ENTERPRISEPACK'
        }
        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Message -like 'No available units*' }
    }

    It 'processes a user listed twice once, so the duplicate takes no seat' {
        Mock Get-MgSubscribedSku {
            @([pscustomobject]@{ SkuId = [guid]$global:skuId; SkuPartNumber = 'ENTERPRISEPACK'; PrepaidUnits = [pscustomobject]@{ Enabled = 2 }; ConsumedUnits = 0 })
        }
        Set-LicenseGraphMock
        @('user1@contoso.com', 'USER1@contoso.com', 'user2@contoso.com') | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        @($assigns.url) | Should -Be @('/users/id1/assignLicense', '/users/id2/assignLicense')
        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Message -like 'No available units*' }
    }
    It 'processes a user repeated in a later batch once' {
        Mock Get-MgSubscribedSku {
            @([pscustomobject]@{ SkuId = [guid]$global:skuId; SkuPartNumber = 'ENTERPRISEPACK'; PrepaidUnits = [pscustomobject]@{ Enabled = 21 }; ConsumedUnits = 0 })
        }
        Set-LicenseGraphMock
        @((New-Upns 20) + 'user1@contoso.com' + 'user21@contoso.com') | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 21
        @($assigns.url) | Should -Contain '/users/id21/assignLicense'
        @($assigns | Where-Object { $_.url -eq '/users/id1/assignLicense' }).Count | Should -Be 1
    }
    It 'writes the unresolved-user message once for a user Graph cannot find' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } }
            }
        }
        'ghost@contoso.com' | Add-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -eq 'Unable to resolve user recipient for ghost@contoso.com'
        }
    }

    It 'removes matched licenses for 3 users with 3 Graph calls' {
        Set-LicenseGraphMock
        New-Upns 3 | Remove-UserMsolAccountSku -License 'ENTERPRISEPACK' -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 3
        foreach ($assign in $assigns) {
            @($assign.body.removeLicenses) | Should -Be @($global:skuId)
            @($assign.body.addLicenses).Count | Should -Be 0
        }
        Should -Invoke Write-NCMessage -Times 3 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' }
    }

    It 'removes every license with -All' {
        Set-LicenseGraphMock
        'user1@contoso.com' | Remove-UserMsolAccountSku -All -Confirm:$false

        $assigns = @($global:SeenRequests | Where-Object { $_.url -like '*/assignLicense' })
        $assigns.Count | Should -Be 1
        @($assigns[0].body.removeLicenses) | Should -Be @($global:skuId)
    }
    It 'Get-UserMsolAccountSku uses 2 Graph calls for 14 users' {
        Set-LicenseGraphMock
        New-Upns 14 | Get-UserMsolAccountSku

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 14 -Exactly -Scope It -ParameterFilter { $Message -like '*Processing user: User *' }
        Should -Invoke Write-NCMessage -Times 14 -Exactly -Scope It -ParameterFilter { $Message -like "*($global:skuId)" }
        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Message -like '*in Graph batches*' }
    }

    It 'Get-UserMsolAccountSku reports a missing user once and keeps going' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                $url = [string]$request.url
                if ($url -like '/users/ghost*') { return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } } }
                if ($url -match '^/users/user1%40') { return @{ status = 200; body = @{ id = 'id1'; userPrincipalName = 'user1@contoso.com'; displayName = 'User 1' } } }
                if ($url -match '^/users/id1/licenseDetails') { return @{ status = 200; body = @{ value = @(@{ skuId = $global:skuId; skuPartNumber = 'ENTERPRISEPACK' }) } } }
                @{ status = 500 }
            }
        }
        @('ghost@contoso.com', 'user1@contoso.com') | Get-UserMsolAccountSku

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -eq 'Unable to resolve user recipient for ghost@contoso.com'
        }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -like '*Processing user: User 1*' }
    }

    It 'Get-UserUsageLocation resolves 14 users with 1 Graph call and keeps output shape' {
        Set-LicenseGraphMock
        $result = @(New-Upns 14 | Get-UserUsageLocation)

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
        $result.Count | Should -Be 14
        ($result[0].PSObject.Properties.Name -join ',') | Should -Be 'UserPrincipalName,DisplayName,UsageLocation,ConfiguredDefaultUsageLocation,MatchesConfiguredDefault'
        $result[0].UserPrincipalName | Should -Be 'user1@contoso.com'
        $result[13].UserPrincipalName | Should -Be 'user14@contoso.com'
        $result[0].MatchesConfiguredDefault | Should -BeTrue
    }

    It 'Set-UserUsageLocation uses 2 Graph calls for 14 users and 1 call with -WhatIf' {
        $global:UsageLocationForUsers = ''
        Set-LicenseGraphMock
        $result = @(New-Upns 14 | Set-UserUsageLocation -UsageLocation DE -PassThru -Confirm:$false)

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        $result.Count | Should -Be 14
        ($result.Action | Select-Object -Unique) | Should -Be 'Updated'
        $patches = @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' })
        $patches.Count | Should -Be 14
        ($patches | ForEach-Object { $_.body.usageLocation } | Select-Object -Unique) | Should -Be 'DE'
    }

    It 'Set-UserUsageLocation -WhatIf only reads' {
        $global:UsageLocationForUsers = ''
        Set-LicenseGraphMock
        New-Upns 14 | Set-UserUsageLocation -UsageLocation DE -WhatIf

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
        @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' }).Count | Should -Be 0
    }

    It 'Set-UserUsageLocation reports a failed PATCH and omits that user from PassThru' {
        $global:UsageLocationForUsers = ''
        $global:PatchFailsFor = 'id2'
        Set-LicenseGraphMock
        $result = @('user1@contoso.com', 'user2@contoso.com' | Set-UserUsageLocation -UsageLocation DE -PassThru -Confirm:$false)

        $result.Count | Should -Be 1
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -eq 'Unable to set usage location (DE) for user2@contoso.com: denied'
        }
    }

    It 'Set-UserUsageLocation -PassThru keeps input order when skipped and updated users are mixed' {
        $global:UsageLocationForUsers = ''
        $global:UsageLocationOverride = @{ '2' = 'DE'; '4' = 'DE' }
        $global:PatchFailsFor = 'id3'
        Set-LicenseGraphMock
        $result = @(New-Upns 5 | Set-UserUsageLocation -UsageLocation DE -PassThru -Confirm:$false)

        ($result.UserPrincipalName -join ',') | Should -Be 'user1@contoso.com,user2@contoso.com,user4@contoso.com,user5@contoso.com'
        ($result.Action -join ',') | Should -Be 'Updated,Skipped,Skipped,Updated'
    }

    It 'Export-MsolAccountSku reads license details for 45 licensed users in 3 batch calls' {
        Mock Get-MgUser { 1..45 | ForEach-Object { [pscustomobject]@{ Id = "id$_"; DisplayName = "User $_"; UserPrincipalName = "user$_@contoso.com"; Mail = "user$_@contoso.com" } } }
        Mock Test-Folder { $TestDrive }
        Mock New-File { Join-Path $TestDrive 'report.csv' }
        $global:NCVars.DateTimeString_CSV = 'yyyyMMdd'
        $global:NCVars.CSV_DefaultLimiter = ';'
        $global:NCVars.CSV_Encoding = 'UTF8'
        Set-LicenseGraphMock

        Export-MsolAccountSku -CSVFolder $TestDrive -Domain 'contoso.com'

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
        $rows = @(Import-Csv -LiteralPath (Join-Path $TestDrive 'report.csv') -Delimiter ';')
        $rows.Count | Should -Be 45
        ($rows[0].PSObject.Properties.Name -join ',') | Should -Be 'DisplayName,UserPrincipalName,PrimarySmtpAddress,Licenses'
        $rows[0].DisplayName | Should -Be 'User 1'
        $rows[44].UserPrincipalName | Should -Be 'user45@contoso.com'
        ($rows.Licenses | Select-Object -Unique) | Should -Be 'ENTERPRISEPACK'
    }

    Context 'Copy and Move user licenses' {
        BeforeEach {
            Mock Update-MgUser {}
            Mock Set-MgUserLicense {}
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $url = [string]$request.url
                    if ($url -like '/users/ghost*') { return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } } }
                    if ($url -match '^/users/src%40contoso\.com') { return @{ status = 200; body = @{ id = 'idsrc'; userPrincipalName = 'src@contoso.com'; displayName = 'Src'; usageLocation = 'IT' } } }
                    if ($url -match '^/users/dst%40contoso\.com') { return @{ status = 200; body = @{ id = 'iddst'; userPrincipalName = 'dst@contoso.com'; displayName = 'Dst'; usageLocation = 'IT' } } }
                    if ($url -match '^/users/idsrc/licenseDetails') { return @{ status = 200; body = @{ value = @(@{ skuId = $global:skuId; skuPartNumber = 'ENTERPRISEPACK' }) } } }
                    if ($url -match '^/users/iddst/licenseDetails') { return @{ status = 200; body = @{ value = @() } } }
                    @{ status = 500 }
                }
            }
        }

        It 'Copy-UserMsolAccountSku uses 2 Graph calls before the write' {
            Copy-UserMsolAccountSku -Source 'src@contoso.com' -Destination 'dst@contoso.com' -Confirm:$false

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            Should -Invoke Set-MgUserLicense -Times 1 -Exactly -Scope It -ParameterFilter { $UserId -eq 'iddst' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' -and $Message -like 'Copied licenses to dst@contoso.com*' }
        }

        It 'Copy-UserMsolAccountSku reports a missing source once and stops' {
            Copy-UserMsolAccountSku -Source 'ghost@contoso.com' -Destination 'dst@contoso.com' -Confirm:$false

            Should -Invoke Set-MgUserLicense -Times 0 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like 'Unable to retrieve source user ghost@contoso.com*' }
        }

        It 'Copy-UserMsolAccountSku aborts when source and destination are the same' {
            Copy-UserMsolAccountSku -Source 'src@contoso.com' -Destination 'SRC@contoso.com' -Confirm:$false

            Should -Invoke Set-MgUserLicense -Times 0 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Source and destination users are the same. Aborting.' }
        }

        It 'Move-UserMsolAccountSku uses 2 Graph calls before the writes and removes from the source after the assignment' {
            Move-UserMsolAccountSku -Source 'src@contoso.com' -Destination 'dst@contoso.com' -Confirm:$false

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            Should -Invoke Set-MgUserLicense -Times 1 -Exactly -Scope It -ParameterFilter { $UserId -eq 'iddst' }
            Should -Invoke Set-MgUserLicense -Times 1 -Exactly -Scope It -ParameterFilter { $UserId -eq 'idsrc' }
        }

        It 'Move-UserMsolAccountSku does not remove from the source when the destination assignment fails' {
            Mock Set-MgUserLicense { throw 'assign denied' }
            Move-UserMsolAccountSku -Source 'src@contoso.com' -Destination 'dst@contoso.com' -Confirm:$false

            Should -Invoke Set-MgUserLicense -Times 1 -Exactly -Scope It
            Should -Invoke Set-MgUserLicense -Times 0 -Exactly -Scope It -ParameterFilter { $UserId -eq 'idsrc' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like 'License assignment to dst@contoso.com failed. Aborting removal from source.*' }
        }

        It 'Move-UserMsolAccountSku reports a missing destination once' {
            Move-UserMsolAccountSku -Source 'src@contoso.com' -Destination 'ghost@contoso.com' -Confirm:$false

            Should -Invoke Set-MgUserLicense -Times 0 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq 'Unable to resolve destination user recipient for ghost@contoso.com' }
        }
    }
}
