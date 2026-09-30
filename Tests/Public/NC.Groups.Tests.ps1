BeforeAll {
    function Test-MgGraphConnection {
        param(
            [string[]]$Scopes,
            [bool]$EnsureExchangeOnline
        )
    }
    function Add-EmptyLine {}
    function Write-NCMessage {
        param(
            [string]$Message,
            [string]$Level
        )
    }
    function Get-MgGroup {
        param(
            [string]$GroupId,
            [string]$Filter,
            [switch]$All,
            [string[]]$Property
        )
    }
    function Get-MgUser {
        param(
            [string]$UserId,
            [string]$Filter,
            [switch]$All
        )
    }
    function Find-UserRecipient {
        param(
            [string]$UserPrincipalName,
            [switch]$PreferGraphIdentity
        )
    }
    function New-MgGroupMember {
        param(
            [string]$GroupId,
            [string]$DirectoryObjectId
        )
    }
    function Get-MgUserMemberOf {
        param(
            [string]$UserId,
            [switch]$All
        )
    }
    function Remove-MgGroupMemberByRef {
        param(
            [string]$GroupId,
            [string]$DirectoryObjectId
        )
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
    function Get-MgEnvironment {
        param([string]$Name)
    }

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
    . "$PSScriptRoot/../../Private/NC-Hlp.Groups.ps1"
    . "$PSScriptRoot/../../Public/NC.Groups.ps1"
}

Describe 'Entra group user identity resolution' {
    BeforeAll {
        $memberUpn = 'employee@contoso.com'
        $guestMail = 'consultant@external.example'
        $memberId = '11111111-1111-1111-1111-111111111111'
        $guestId = 'guest-id'
        $groupId = '33333333-3333-3333-3333-333333333333'

        $memberUser = [pscustomobject]@{
            Id                = $memberId
            UserPrincipalName = $memberUpn
            DisplayName       = 'Tenant Employee'
        }
        $guestUser = [pscustomobject]@{
            Id                = $guestId
            UserPrincipalName = 'consultant_external.example#EXT#@contoso.onmicrosoft.com'
            DisplayName       = 'External Consultant'
        }
        $group = [pscustomobject]@{
            Id                    = 'group-id'
            DisplayName           = 'Group'
            OnPremisesSyncEnabled = $false
        }
        $membership = [pscustomobject]@{
            Id                   = $groupId
            AdditionalProperties = @{
                displayName = 'Cloud Group'
                mail        = 'cloud-group@contoso.com'
            }
        }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Get-MgContext { [pscustomobject]@{ Environment = 'Global' } }
        Mock Get-MgGroup { $group }
        Mock Find-UserRecipient {
            if ($PreferGraphIdentity) {
                return $guestId
            }

            return $UserPrincipalName
        }
        Mock Get-MgUserMemberOf { @($membership) }
    }



    It 'reads memberships for a tenant member through the unchanged direct lookup' {
        Mock Get-MgUser {
            if ($UserId -eq $memberUpn) {
                return $memberUser
            }

            throw "Unexpected user lookup: $UserId"
        }

        $null = Get-EntraGroupUser -UserIdentifier $memberUpn

        Assert-MockCalled Find-UserRecipient -Times 0 -Scope It
        Assert-MockCalled Get-MgUserMemberOf -Times 1 -Scope It -ParameterFilter {
            $UserId -eq $memberId
        }
    }

    It 'reads memberships for a guest through a Graph-compatible fallback identity' {
        Mock Get-MgUser {
            if ($UserId -eq $guestId) {
                return $guestUser
            }

            throw "User not found: $UserId"
        }

        $null = Get-EntraGroupUser -UserIdentifier $guestMail

        Assert-MockCalled Find-UserRecipient -Times 1 -Scope It -ParameterFilter {
            $UserPrincipalName -eq $guestMail -and $PreferGraphIdentity
        }
        Assert-MockCalled Get-MgUserMemberOf -Times 1 -Scope It -ParameterFilter {
            $UserId -eq $guestId
        }
    }

    It 'adds a tenant member through the unchanged direct lookup' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.url -like '/users/member%40contoso.com*') { return @{ status = 200; body = @{ id = 'member-id'; userPrincipalName = 'member@contoso.com' } } }
                if ($request.method -eq 'POST' -and $request.url -eq '/groups/group-id/members/$ref') { return @{ status = 204 } }
                @{ status = 500 }
            }
        }

        Add-EntraGroupUser -GroupName 'Group' -UserIdentifier 'member@contoso.com' -Confirm:$false

        Should -Invoke Find-UserRecipient -Times 0 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Added user 'member@contoso.com' to group 'Group'." -and $Level -eq 'SUCCESS' }
    }

    It 'adds a guest through a Graph-compatible fallback identity' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.url -like '/users/guest%40external.com*') { return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } } }
                if ($request.url -like '/users/guest-id*') { return @{ status = 200; body = @{ id = 'guest-id'; userPrincipalName = 'guest_external.com#EXT#@contoso.onmicrosoft.com' } } }
                if ($request.url -eq '/groups/group-id/members/$ref') { return @{ status = 204 } }
                @{ status = 500 }
            }
        }

        Add-EntraGroupUser -GroupName 'Group' -UserIdentifier 'guest@external.com' -Confirm:$false

        Should -Invoke Find-UserRecipient -Times 1 -Exactly -Scope It -ParameterFilter { $UserPrincipalName -eq 'guest@external.com' -and $PreferGraphIdentity }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -like "Added user 'guest_external.com#EXT#@contoso.onmicrosoft.com'*" }
    }

    It 'reports existing members and keeps processing the others' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') {
                    $upn = [uri]::UnescapeDataString(($request.url -replace '^/users/([^?]+)\?.*$', '$1'))
                    return @{ status = 200; body = @{ id = "id-$upn"; userPrincipalName = $upn } }
                }
                if ($request.body.'@odata.id' -like '*id-old@contoso.com') {
                    return @{ status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'One or more added object references already exist for the following modified properties: ''members''.' } } }
                }
                @{ status = 204 }
            }
        }

        $result = @('old@contoso.com', 'new@contoso.com' | Add-EntraGroupUser -GroupName 'Group' -PassThru -Confirm:$false)

        $result.Count | Should -Be 2
        $result[0].MemberName | Should -Be 'old@contoso.com'
        $result[0].Status | Should -Be 'Exists'
        $result[1].Status | Should -Be 'Added'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "User 'old@contoso.com' is already a member of 'Group'." -and $Level -eq 'WARNING' }
    }

    It 'adds 14 users with one resolve batch and one add batch' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') { return @{ status = 200; body = @{ id = "id-$($request.id)"; userPrincipalName = "$($request.id)@contoso.com" } } }
                @{ status = 204 }
            }
        }
        $users = @(1..14 | ForEach-Object { "user$_@contoso.com" })

        $users | Add-EntraGroupUser -GroupName 'Group' -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
    }

    It 'does not send writes with -WhatIf' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder { param($request) @{ status = 200; body = @{ id = 'member-id'; userPrincipalName = 'member@contoso.com' } } }
        }

        Add-EntraGroupUser -GroupName 'Group' -UserIdentifier 'member@contoso.com' -WhatIf

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
    }

    It 'removes a tenant member through the unchanged direct lookup' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') { return @{ status = 200; body = @{ id = 'member-id'; userPrincipalName = 'member@contoso.com' } } }
                if ($request.method -eq 'DELETE' -and $request.url -eq '/groups/group-id/members/member-id/$ref') { return @{ status = 204 } }
                @{ status = 500 }
            }
        }

        Remove-EntraGroupUser -GroupName 'Group' -UserIdentifier 'member@contoso.com' -Confirm:$false

        Should -Invoke Find-UserRecipient -Times 0 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Removed user 'member@contoso.com' from group 'Group'." -and $Level -eq 'SUCCESS' }
    }

    It 'removes a guest through a Graph-compatible fallback identity' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.url -like '/users/guest%40external.com*') { return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } } }
                if ($request.url -like '/users/guest-id*') { return @{ status = 200; body = @{ id = 'guest-id'; userPrincipalName = 'guest_external.com#EXT#@contoso.onmicrosoft.com' } } }
                if ($request.url -eq '/groups/group-id/members/guest-id/$ref') { return @{ status = 204 } }
                @{ status = 500 }
            }
        }

        Remove-EntraGroupUser -GroupName 'Group' -UserIdentifier 'guest@external.com' -Confirm:$false

        Should -Invoke Find-UserRecipient -Times 1 -Exactly -Scope It -ParameterFilter { $UserPrincipalName -eq 'guest@external.com' -and $PreferGraphIdentity }
    }

    It 'reports users that are not members as NotFound' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') { return @{ status = 200; body = @{ id = 'member-id'; userPrincipalName = 'member@contoso.com' } } }
                @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = "Resource 'member-id' does not exist or one of its queried reference-property objects are not present." } } }
            }
        }

        $result = Remove-EntraGroupUser -GroupName 'Group' -UserIdentifier 'member@contoso.com' -PassThru -Confirm:$false

        $result.Status | Should -Be 'NotFound'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "User 'member@contoso.com' is not a member of 'Group'" -and $Level -eq 'WARNING' }
    }
}
