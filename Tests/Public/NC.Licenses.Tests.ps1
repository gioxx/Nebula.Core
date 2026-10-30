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
        $global:LicenseChangesFailFor = ''
        $global:UsageLocationOverride = @{}
        $global:AssignedFor = @()
        $global:LicenseAssignmentStatesForUsers = @(@{ skuId = $global:skuId; assignedByGroup = $null; disabledPlans = @('plan-old') })
        $global:LicenseDetailsForUsers = @(@{ skuId = $global:skuId; skuPartNumber = 'ENTERPRISEPACK' })
        $global:OmitAssignmentStates = $false
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
                        $userBody = @{ id = "id$n"; userPrincipalName = "user$n@contoso.com"; displayName = "User $n"; usageLocation = $loc; assignedLicenses = $owned }
                        if (-not $global:OmitAssignmentStates) { $userBody.licenseAssignmentStates = $global:LicenseAssignmentStatesForUsers }
                        return @{ status = 200; body = $userBody }
                    }
                    if ($request.method -eq 'GET' -and $url -match '^/users/(id\d+)/licenseDetails') {
                        return @{ status = 200; body = @{ value = $global:LicenseDetailsForUsers } }
                    }
                    if ($request.method -eq 'PATCH' -and $url -match '^/users/(id\d+)$') {
                        if ($Matches[1] -eq $global:PatchFailsFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 204 }
                    }
                    if ($request.method -eq 'POST' -and $url -match '^/users/(id\d+)/assignLicense$') {
                        if ($Matches[1] -eq $global:LicenseChangesFailFor) { return @{ status = 400; body = @{ error = @{ code = 'LicenseConflict'; message = 'license conflict' } } } }
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

    Context 'Set user licenses in one operation' {
        BeforeEach {
            $global:newSkuId = '11111111-1111-1111-1111-111111111111'
            Mock Get-MgSubscribedSku {
                @(
                    [pscustomobject]@{ SkuId = $global:skuId; SkuPartNumber = 'ENTERPRISEPACK' }
                    [pscustomobject]@{ SkuId = $global:newSkuId; SkuPartNumber = 'NEWPACK' }
                )
            }
            Set-LicenseGraphMock
        }

        It 'combines additions and removals in a single user write' {
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            $writes[0].url | Should -Be '/users/id1/assignLicense'
            $writes[0].body.addLicenses.skuId | Should -Be $global:newSkuId
            $writes[0].body.removeLicenses | Should -Be @($global:skuId)
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'SUCCESS' -and $Message -like '*added: NEWPACK*' -and $Message -like '*removed: ENTERPRISEPACK*'
            }
        }

        It 'requests Graph scopes for license reads and writes and user updates' {
            Set-UserMsolAccountSku user1@contoso.com -RemoveLicense ENTERPRISEPACK -Confirm:$false
            Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter {
                $Scopes -contains 'LicenseAssignment.Read.All' -and
                $Scopes -contains 'LicenseAssignment.ReadWrite.All' -and
                $Scopes -contains 'User.ReadWrite.All'
            }
        }

        It 'merges repeated users across more than twenty pipeline rows' {
            $rows = @(1..21 | ForEach-Object { [pscustomobject]@{ UserPrincipalName = 'user1@contoso.com'; AddLicense = 'NEWPACK' } })
            $rows += [pscustomobject]@{ UserPrincipalName = 'user1@contoso.com'; RemoveLicense = 'ENTERPRISEPACK' }
            $rows | Set-UserMsolAccountSku -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            @($writes[0].body.addLicenses).Count | Should -Be 1
            $writes[0].body.removeLicenses | Should -Be @($global:skuId)
        }

        It 'rejects overlapping SKU operations even when identified differently' {
            Set-UserMsolAccountSku user1@contoso.com -AddLicense ENTERPRISEPACK -RemoveLicense $global:skuId -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        }

        It 'rejects all changes for a user with an unresolved license on another row' {
            @(
                [pscustomobject]@{ UserPrincipalName = 'user1@contoso.com'; AddLicense = 'NEWPACK' }
                [pscustomobject]@{ UserPrincipalName = 'user1@contoso.com'; RemoveLicense = 'missing' }
            ) | Set-UserMsolAccountSku -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        }

        It 'skips users whose license state cannot be read' {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    if ($request.url -like '*/licenseDetails*') { return @{ status = 403; body = @{ error = @{ message = 'denied' } } } }
                    @{ status = 200; body = @{ id = 'id1'; userPrincipalName = 'user1@contoso.com' } }
                }
            }
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        }

        It 'previews the combined operation without writing' {
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -WhatIf
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'INFO' -and $Message -like '*changed successfully: 0; failed: 0; unchanged: 0; skipped/not executed: 1*'
            }
        }

        It 'supplies an empty opposite array for <Operation> only' -ForEach @(
            @{ Operation = 'Add'; Parameters = @{ AddLicense = 'NEWPACK' }; Empty = 'removeLicenses'; Populated = 'addLicenses' }
            @{ Operation = 'Remove'; Parameters = @{ RemoveLicense = 'ENTERPRISEPACK' }; Empty = 'addLicenses'; Populated = 'removeLicenses' }
        ) {
            Set-UserMsolAccountSku user1@contoso.com @Parameters -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            ($writes[0].body.$Empty -is [array]) | Should -BeTrue
            @($writes[0].body.$Empty).Count | Should -Be 0
            @($writes[0].body.$Populated).Count | Should -Be 1
        }

        It 'does not fetch tenant SKUs for remove-only input with <Case> additions' -ForEach @(
            @{ Case = 'omitted'; Arguments = @{} }
            @{ Case = 'blank'; Arguments = @{ AddLicense = @($null, '', ' ') } }
        ) {
            Set-UserMsolAccountSku user1@contoso.com -RemoveLicense ENTERPRISEPACK @Arguments -Confirm:$false
            Should -Invoke Get-MgSubscribedSku -Times 0 -Exactly -Scope It
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            $writes[0].body.removeLicenses | Should -Be @($global:skuId)
            @($writes[0].body.addLicenses).Count | Should -Be 0
        }

        It 'rejects an empty change' {
            Set-UserMsolAccountSku user1@contoso.com -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        }

        It 'preserves service plans by omitting licenses already assigned directly' {
            Set-UserMsolAccountSku user1@contoso.com -AddLicense ENTERPRISEPACK,NEWPACK -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            @($writes[0].body.addLicenses).Count | Should -Be 1
            $writes[0].body.addLicenses.skuId | Should -Be $global:newSkuId
            ($writes[0].body.addLicenses[0].disabledPlans -is [array]) | Should -BeTrue
            @($writes[0].body.addLicenses[0].disabledPlans).Count | Should -Be 0
        }

        It 'adds a direct target assignment even when that SKU is already inherited from a group' {
            $global:LicenseAssignmentStatesForUsers += @{ skuId = $global:newSkuId; assignedByGroup = 'group1'; disabledPlans = @('plan-group') }
            $global:LicenseDetailsForUsers += @{ skuId = $global:newSkuId; skuPartNumber = 'NEWPACK' }
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            $writes[0].body.addLicenses.skuId | Should -Be $global:newSkuId
            $writes[0].body.removeLicenses | Should -Be @($global:skuId)
        }

        It 'rejects a removal when the SKU is inherited only from a group' {
            $global:LicenseAssignmentStatesForUsers = @(@{ skuId = $global:skuId; assignedByGroup = 'group1'; disabledPlans = @() })
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*inherited-only*' }
        }

        It 'treats an already removed tenant SKU as a no-op on a repeated replacement' {
            $global:LicenseAssignmentStatesForUsers = @(@{ skuId = $global:newSkuId; assignedByGroup = $null; disabledPlans = @() })
            $global:LicenseDetailsForUsers = @(@{ skuId = $global:newSkuId; skuPartNumber = 'NEWPACK' })

            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false

            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'SUCCESS' -and $Message -like '*No license changes needed*' -and
                $Message -like '*already assigned: NEWPACK*' -and $Message -like '*already absent: ENTERPRISEPACK*'
            }
            Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        }

        It 'suppresses per-user no-change messages when QuietNoChanges is set' {
            $global:LicenseAssignmentStatesForUsers = @(@{ skuId = $global:newSkuId; assignedByGroup = $null; disabledPlans = @() })
            $global:LicenseDetailsForUsers = @(@{ skuId = $global:newSkuId; skuPartNumber = 'NEWPACK' })

            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -QuietNoChanges -Confirm:$false

            Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter {
                $Message -like 'No license changes needed for*'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'INFO' -and $Message -like '*changed successfully: 0; failed: 0; unchanged: 1; skipped/not executed: 0*'
            }
        }

        It 'sets the configured usage location before assigning a new license' {
            $global:UsageLocationForUsers = ''
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object { $_.method -in @('PATCH', 'POST') })
            $writes.Count | Should -Be 2
            $writes[0].method | Should -Be 'PATCH'
            $writes[0].body.usageLocation | Should -Be 'IT'
            $writes[1].method | Should -Be 'POST'
            $writes[1].body.addLicenses.skuId | Should -Be $global:newSkuId
        }

        It 'does not assign a license if setting usage location fails' {
            $global:UsageLocationForUsers = ''
            $global:PatchFailsFor = 'id1'
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*usage location*' }
        }

        It 'skips writes if direct assignment state is <Case>' -ForEach @(
            @{ Case = 'omitted'; Omit = $true; States = @() }
            @{ Case = 'null'; Omit = $false; States = $null }
            @{ Case = 'missing an assigned SKU'; Omit = $false; States = @() }
            @{ Case = 'missing assignment source'; Omit = $false; States = @(@{ skuId = '6fd2c87f-b296-42f0-b197-1e91e994b900'; disabledPlans = @() }) }
            @{ Case = 'missing SKU ID'; Omit = $false; States = @(@{ assignedByGroup = $null; disabledPlans = @() }) }
            @{ Case = 'missing disabled plans'; Omit = $false; States = @(@{ skuId = '6fd2c87f-b296-42f0-b197-1e91e994b900'; assignedByGroup = $null }) }
        ) {
            $global:OmitAssignmentStates = $Omit
            $global:LicenseAssignmentStatesForUsers = $States
            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false
            @($global:SeenRequests | Where-Object method -eq POST).Count | Should -Be 0
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*assignment state*' }
        }

        It 'resolves friendly names through the catalog' {
            Mock Get-LicenseCatalog { [pscustomobject]@{ Lookup = @{ NEWPACK = 'New Product' }; CustomLookup = @{} } }
            Mock Get-LicenseDisplayName { if ($SkuPartNumber -eq 'NEWPACK') { 'New Product' } else { 'Old Product' } }
            Set-UserMsolAccountSku user1@contoso.com -AddLicense ' New Product ' -RemoveLicense 'old product' -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            $writes[0].body.addLicenses.skuId | Should -Be $global:newSkuId
            $writes[0].body.removeLicenses | Should -Be @($global:skuId)
        }

        It 'merges different identifiers that resolve to the same user ID' {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    if ($request.url -like '*/licenseDetails*') { return @{ status = 200; body = @{ value = @(@{ skuId = $global:skuId; skuPartNumber = 'ENTERPRISEPACK' }) } } }
                    @{ status = 200; body = @{ id = 'id1'; userPrincipalName = 'user1@contoso.com'; licenseAssignmentStates = $global:LicenseAssignmentStatesForUsers } }
                }
            }
            @(
                [pscustomobject]@{ UserPrincipalName = 'user1@contoso.com'; AddLicense = 'NEWPACK' }
                [pscustomobject]@{ UserPrincipalName = 'id1'; RemoveLicense = 'ENTERPRISEPACK' }
            ) | Set-UserMsolAccountSku -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            $writes[0].url | Should -Be '/users/id1/assignLicense'
            $writes[0].body.addLicenses.skuId | Should -Be $global:newSkuId
            $writes[0].body.removeLicenses | Should -Be @($global:skuId)
            @($global:SeenRequests | Where-Object url -like '*/licenseDetails*').Count | Should -Be 1
        }

        It 'continues independent users when another user cannot be resolved' {
            @('ghost@contoso.com', 'user1@contoso.com') | Set-UserMsolAccountSku -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false
            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 1
            $writes[0].url | Should -Be '/users/id1/assignLicense'
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*ghost@contoso.com*' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'INFO' -and $Message -like '*changed successfully: 1; failed: 1; unchanged: 0; skipped/not executed: 0*'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'ERROR' -and $Message -like '*Failed users:*ghost@contoso.com*'
            }
        }

        It 'counts and identifies users rejected by the license assignment request' {
            $global:LicenseChangesFailFor = 'id1'
            Set-LicenseGraphMock

            Set-UserMsolAccountSku user1@contoso.com -AddLicense NEWPACK -RemoveLicense ENTERPRISEPACK -Confirm:$false

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'INFO' -and $Message -like '*changed successfully: 0; failed: 1; unchanged: 0; skipped/not executed: 0*'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter {
                $Level -eq 'ERROR' -and $Message -like '*Failed users: user1@contoso.com (license conflict)*'
            }
        }
    }

    Context 'Swap user licenses without spare tenant seats' {
        BeforeEach {
            $global:swapFirstSkuId = [guid]'11111111-1111-1111-1111-111111111111'
            $global:swapSecondSkuId = [guid]'22222222-2222-2222-2222-222222222222'
            $global:SwapFailOn = @()
            $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
            Mock Get-LicenseCatalog { $null }
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($request.method -eq 'GET' -and $url -match '^/users/(user[12]%40contoso\.com)\?') {
                        $user = $Matches[1]
                        if ($user -like 'user1*') {
                            return @{ status = 200; body = @{ id = 'id1'; userPrincipalName = 'user1@contoso.com'; usageLocation = 'IT'; licenseAssignmentStates = @(@{ skuId = $global:swapFirstSkuId; assignedByGroup = $null; disabledPlans = @('first-disabled') }) } }
                        }
                        return @{ status = 200; body = @{ id = 'id2'; userPrincipalName = 'user2@contoso.com'; usageLocation = 'IT'; licenseAssignmentStates = @(@{ skuId = $global:swapSecondSkuId; assignedByGroup = $null; disabledPlans = @('second-disabled') }) } }
                    }
                    if ($request.method -eq 'GET' -and $url -match '^/users/(id[12])/licenseDetails$') {
                        if ($Matches[1] -eq 'id1') { return @{ status = 200; body = @{ value = @(@{ skuId = $global:swapFirstSkuId; skuPartNumber = 'FIRSTPACK' }) } } }
                        return @{ status = 200; body = @{ value = @(@{ skuId = $global:swapSecondSkuId; skuPartNumber = 'SECONDPACK' }) } }
                    }
                    if ($request.method -eq 'POST' -and $url -match '^/users/(id[12])/assignLicense$') {
                        $kind = if (@($request.body.addLicenses).Count -gt 0) { 'add' } else { 'remove' }
                        $failureKey = "${kind}:$($Matches[1])"
                        if ($kind -eq 'add') { $failureKey += ":$([string]$request.body.addLicenses[0].skuId)" }
                        if ($global:SwapFailOn -contains $failureKey) { return @{ status = 400; body = @{ error = @{ code = 'RequestFailed'; message = 'simulated failure' } } } }
                        return @{ status = 200; body = @{ id = $Matches[1] } }
                    }
                    return @{ status = 500; body = @{ error = @{ message = "Unexpected request: $($request.method) $url" } } }
                }
            }
        }

        It 'removes both selected licenses before assigning them reciprocally' {
            Swap-UserMsolAccountSku -FirstUserPrincipalName user1@contoso.com -FirstUserLicense FIRSTPACK -SecondUserPrincipalName user2@contoso.com -SecondUserLicense SECONDPACK -Confirm:$false

            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 4
            @($writes[0].body.removeLicenses) | Should -Be @($global:swapFirstSkuId)
            @($writes[1].body.removeLicenses) | Should -Be @($global:swapSecondSkuId)
            @($writes[0].body.addLicenses).Count | Should -Be 0
            @($writes[1].body.addLicenses).Count | Should -Be 0
            $writes[2].url | Should -Be '/users/id1/assignLicense'
            $writes[2].body.addLicenses.skuId | Should -Be $global:swapSecondSkuId
            $writes[2].body.addLicenses.disabledPlans | Should -Be @('second-disabled')
            $writes[3].url | Should -Be '/users/id2/assignLicense'
            $writes[3].body.addLicenses.skuId | Should -Be $global:swapFirstSkuId
            $writes[3].body.addLicenses.disabledPlans | Should -Be @('first-disabled')
        }

        It 'restores a successful removal and does not start additions if the other removal fails' {
            $global:SwapFailOn = @('remove:id2')

            Swap-UserMsolAccountSku -FirstUserPrincipalName user1@contoso.com -FirstUserLicense FIRSTPACK -SecondUserPrincipalName user2@contoso.com -SecondUserLicense SECONDPACK -Confirm:$false

            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 3
            @($writes[2].body.addLicenses.skuId) | Should -Be @($global:swapFirstSkuId)
            $writes[2].url | Should -Be '/users/id1/assignLicense'
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*simulated failure*' }
        }

        It 'removes partial reciprocal additions and restores both original licenses after an add failure' {
            $global:SwapFailOn = @("add:id2:$($global:swapFirstSkuId)")

            Swap-UserMsolAccountSku -FirstUserPrincipalName user1@contoso.com -FirstUserLicense FIRSTPACK -SecondUserPrincipalName user2@contoso.com -SecondUserLicense SECONDPACK -Confirm:$false

            $writes = @($global:SeenRequests | Where-Object method -eq POST)
            $writes.Count | Should -Be 7
            @($writes[4].body.removeLicenses) | Should -Be @($global:swapSecondSkuId)
            $writes[4].url | Should -Be '/users/id1/assignLicense'
            $writes[5].url | Should -Be '/users/id1/assignLicense'
            @($writes[5].body.addLicenses.skuId) | Should -Be @($global:swapFirstSkuId)
            $writes[6].url | Should -Be '/users/id2/assignLicense'
            @($writes[6].body.addLicenses.skuId) | Should -Be @($global:swapSecondSkuId)
            Should -Invoke Write-NCMessage -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*simulated failure*' }
        }
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
