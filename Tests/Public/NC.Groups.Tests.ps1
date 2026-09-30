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
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.url -like "/users/$memberId/memberOf*") { return @{ status = 200; body = @{ value = @(@{ id = $groupId; displayName = 'Cloud Group'; mail = 'cloud-group@contoso.com' }) } } }
                if ($request.url -like '/users/employee%40contoso.com*') { return @{ status = 200; body = @{ id = $memberId; userPrincipalName = $memberUpn; displayName = 'Tenant Employee' } } }
                @{ status = 500 }
            }
        }

        $null = Get-EntraGroupUser -UserIdentifier $memberUpn

        Assert-MockCalled Find-UserRecipient -Times 0 -Scope It
        Assert-MockCalled Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter {
            $Body -like "*/users/$memberId/memberOf*"
        }
    }

    It 'reads memberships for a guest through a Graph-compatible fallback identity' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.url -like "/users/$guestId/memberOf*") { return @{ status = 200; body = @{ value = @(@{ id = $groupId; displayName = 'Cloud Group' }) } } }
                if ($request.url -like "/users/$guestId*") { return @{ status = 200; body = @{ id = $guestId; userPrincipalName = 'consultant_external.example#EXT#@contoso.onmicrosoft.com'; displayName = 'External Consultant' } } }
                if ($request.url -like '/users/consultant%40external.example*') { return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } } }
                @{ status = 500 }
            }
        }

        $null = Get-EntraGroupUser -UserIdentifier $guestMail

        Assert-MockCalled Find-UserRecipient -Times 1 -Scope It -ParameterFilter {
            $UserPrincipalName -eq $guestMail -and $PreferGraphIdentity
        }
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter {
            $Body -like "*/users/$guestId/memberOf*"
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

Describe 'Entra group device batching' {
    BeforeAll {
        $group = [pscustomobject]@{ Id = 'group-id'; DisplayName = 'Group'; OnPremisesSyncEnabled = $false }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Get-MgContext { [pscustomobject]@{ Environment = 'Global' } }
        Mock Get-MgGroup { $group }
    }

    It 'adds devices resolved by display name with one lookup batch and one add batch' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') {
                    $decoded = [uri]::UnescapeDataString($request.url)
                    $name = $decoded -replace "^/devices\?\`$filter=displayName eq '([^']+)'.*$", '$1'
                    return @{ status = 200; body = @{ value = @(@{ id = "id-$name"; displayName = $name }) } }
                }
                @{ status = 204 }
            }
        }
        $devices = @(1..14 | ForEach-Object { "PC$_" })

        $devices | Add-EntraGroupDevice -GroupName 'Group' -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
    }

    It 'reports Added with the device label and the original message' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') { return @{ status = 200; body = @{ value = @(@{ id = 'dev-id'; displayName = 'PC1' }) } } }
                if ($request.method -eq 'POST' -and $request.url -eq '/groups/group-id/members/$ref' -and $request.body.'@odata.id' -like '*dev-id') { return @{ status = 204 } }
                @{ status = 500 }
            }
        }

        $result = Add-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier 'PC1' -PassThru -Confirm:$false

        $result.Status | Should -Be 'Added'
        $result.MemberId | Should -Be 'dev-id'
        $result.MemberType | Should -Be 'Device'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Added device 'PC1' to group 'Group'" -and $Level -eq 'SUCCESS' }
    }

    It 'reports Exists when the device is already a member' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                @{ status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'One or more added object references already exist for the following modified properties: ''members''.' } } }
            }
        }

        $result = Add-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier '11111111-1111-1111-1111-111111111111' -PassThru -Confirm:$false

        $result.Status | Should -Be 'Exists'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Device '11111111-1111-1111-1111-111111111111' is already a member of 'Group'" -and $Level -eq 'WARNING' }
    }

    It 'reports Failed with the Graph error message' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
            }
        }

        $result = Add-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier '11111111-1111-1111-1111-111111111111' -PassThru -Confirm:$false

        $result.Status | Should -Be 'Failed'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Failed to add device '11111111-1111-1111-1111-111111111111' to 'Group': denied" -and $Level -eq 'ERROR' }
    }

    It 'warns when a device is not found and when several match' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') {
                    if ($request.url -like '*Ghost*') { return @{ status = 200; body = @{ value = @() } } }
                    return @{ status = 200; body = @{ value = @(@{ id = 'first-id'; displayName = 'Twin' }, @{ id = 'second-id'; displayName = 'Twin' }) } }
                }
                @{ status = 204 }
            }
        }

        $result = @('Ghost', 'Twin' | Add-EntraGroupDevice -GroupName 'Group' -PassThru -Confirm:$false)

        $result.Count | Should -Be 1
        $result[0].MemberId | Should -Be 'first-id'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Device 'Ghost' not found" -and $Level -eq 'WARNING' }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Multiple devices matched 'Twin'. Using the first result (Twin)" -and $Level -eq 'WARNING' }
    }

    It 'reports a lookup failure for a device name' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
            }
        }

        Add-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier 'PC1' -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Unable to resolve device 'PC1': denied" -and $Level -eq 'ERROR' }
    }

    It 'does not send writes with -WhatIf' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder { param($request) @{ status = 200; body = @{ value = @(@{ id = 'dev-id'; displayName = 'PC1' }) } } }
        }

        Add-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier 'PC1' -WhatIf

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
    }

    It 'removes devices with one batched DELETE per chunk' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                if ($request.method -eq 'GET') { return @{ status = 200; body = @{ value = @(@{ id = 'dev-id'; displayName = 'PC1' }) } } }
                if ($request.method -eq 'DELETE' -and $request.url -eq '/groups/group-id/members/dev-id/$ref') { return @{ status = 204 } }
                @{ status = 500 }
            }
        }

        $result = Remove-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier 'PC1' -PassThru -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        $result.Status | Should -Be 'Removed'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Removed device 'PC1' from group 'Group'" -and $Level -eq 'SUCCESS' }
    }

    It 'reports devices that are not members as NotFound' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = "Resource 'x' does not exist or one of its queried reference-property objects are not present." } } }
            }
        }

        $result = Remove-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier '11111111-1111-1111-1111-111111111111' -PassThru -Confirm:$false

        $result.Status | Should -Be 'NotFound'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Device '11111111-1111-1111-1111-111111111111' is not a member of 'Group'" -and $Level -eq 'WARNING' }
    }

    It 'reports a failed removal with the Graph error message' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($request)
                @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
            }
        }

        $result = Remove-EntraGroupDevice -GroupName 'Group' -DeviceIdentifier '11111111-1111-1111-1111-111111111111' -PassThru -Confirm:$false

        $result.Status | Should -Be 'Failed'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Failed to remove device '11111111-1111-1111-1111-111111111111' from 'Group': denied" -and $Level -eq 'ERROR' }
    }
}

Describe 'Entra group owner batching' {
    BeforeAll {
        function Get-MgGroupMember {
            param(
                [string]$GroupId,
                [switch]$All
            )
        }
        $group = [pscustomobject]@{ Id = 'group-id'; DisplayName = 'Group'; OnPremisesSyncEnabled = $false }
        $guid1 = '11111111-1111-1111-1111-111111111111'
        $script:ownerBatchSizes = [System.Collections.Generic.List[object]]::new()
        $script:ownerDirectCalls = [System.Collections.Generic.List[string]]::new()
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Get-MgContext { [pscustomobject]@{ Environment = 'Global' } }
        Mock Get-MgGroup { $group }
        $script:ownerBatchSizes.Clear()
        $script:ownerDirectCalls.Clear()
        $script:ownerWriteAnswer = { param($request) @{ status = 204 } }
        $script:ownerDirectAnswer = { param($uri) @{ value = @() } }
        Mock Invoke-MgGraphRequest {
            if ($Uri -like '*$batch') {
                $payload = $Body | ConvertFrom-Json
                $script:ownerBatchSizes.Add(@($payload.requests | ForEach-Object { "$($_.method) $($_.url)" }))
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    if ($request.method -eq 'GET') {
                        $name = [uri]::UnescapeDataString($request.url) -replace '^/users/([^?]+)\?.*$', '$1'
                        if ($name -like 'ghost*') {
                            return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'not found' } } }
                        }
                        return @{ status = 200; body = @{ id = "id-$name"; userPrincipalName = $name; displayName = "Name $name" } }
                    }
                    & $script:ownerWriteAnswer $request
                }
            }
            else {
                $script:ownerDirectCalls.Add("$Method $Uri")
                # The real cmdlet output is a dictionary; the functions read it as an object, so hand back an object.
                (& $script:ownerDirectAnswer $Uri) | ConvertTo-Json -Depth 6 | ConvertFrom-Json
            }
        }
    }

    Context 'Resolve-NCEntraOwnerBatch' {
        It 'resolves UPNs in one batch and keeps the Resolve-NCEntraOwner shape' {
            $result = @(Resolve-NCEntraOwnerBatch -OwnerIdentifier @('a@contoso.com', 'b@contoso.com'))

            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
            $result.Count | Should -Be 2
            $result[0].Id | Should -Be 'id-a@contoso.com'
            $result[0].Label | Should -Be 'a@contoso.com'
            ($result[0].PSObject.Properties.Name -join ',') | Should -Be 'Id,Label'
        }

        It 'passes unresolvable GUIDs through as placeholders like Resolve-NCEntraOwner' {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder { param($request) @{ status = 404; body = @{ error = @{ code = 'x'; message = 'nope' } } } }
            }

            $result = @(Resolve-NCEntraOwnerBatch -OwnerIdentifier @($guid1) -TreatInputAsId)

            $result[0].Id | Should -Be $guid1
            $result[0].Label | Should -Be $guid1
        }

        It 'hands inputs that fail the batch lookup to Resolve-NCEntraOwner and keeps input order' {
            Mock Resolve-NCEntraOwner { [pscustomobject]@{ Id = 'sp-id'; Label = 'App' } }

            $result = @(Resolve-NCEntraOwnerBatch -OwnerIdentifier @('a@contoso.com', 'ghost@contoso.com', 'b@contoso.com'))

            $result.Label | Should -Be @('a@contoso.com', 'App', 'b@contoso.com')
            Should -Invoke Resolve-NCEntraOwner -Times 1 -Exactly -Scope It -ParameterFilter { $OwnerIdentifier -eq 'ghost@contoso.com' }
        }

        It 'prints the original not-found warning exactly once' {
            Mock Get-MgUser { throw 'not found' }
            Mock Find-UserRecipient { $null }

            $result = @(Resolve-NCEntraOwnerBatch -OwnerIdentifier @('ghost@contoso.com'))

            $result.Count | Should -Be 0
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Owner 'ghost@contoso.com' not found." -and $Level -eq 'WARNING' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It
        }
    }

    Context 'Add-EntraGroupOwner' {
        It 'adds 3 UPN owners with one resolve batch and one add batch' {
            $result = @(Add-EntraGroupOwner -GroupName 'Group' -OwnerIdentifier 'a@contoso.com', 'b@contoso.com', 'c@contoso.com' -PassThru -Confirm:$false)

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            $script:ownerBatchSizes[1] | Should -Be @('POST /groups/group-id/owners/$ref', 'POST /groups/group-id/owners/$ref', 'POST /groups/group-id/owners/$ref')
            $result.Status | Should -Be @('Added', 'Added', 'Added')
            ($result[0].PSObject.Properties.Name -join ',') | Should -Be 'GroupName,GroupId,OwnerName,OwnerId,Status'
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Added owner 'a@contoso.com' to group 'Group'." -and $Level -eq 'SUCCESS' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing 3 owner(s) in Graph batches (20 per request) ...' -and $Level -eq 'INFO' }
        }

        It 'reports Exists and Failed with the original messages' {
            $script:ownerWriteAnswer = {
                param($request)
                if ($request.body.'@odata.id' -like '*id-old@contoso.com') {
                    return @{ status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'One or more added object references already exist for the following modified properties: owners.' } } }
                }
                @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
            }

            $result = @(Add-EntraGroupOwner -GroupName 'Group' -OwnerIdentifier 'old@contoso.com', 'bad@contoso.com' -PassThru -Confirm:$false)

            $result.Status | Should -Be @('Exists', 'Failed')
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Owner 'old@contoso.com' is already an owner of 'Group'." -and $Level -eq 'WARNING' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Failed to add owner 'bad@contoso.com' to 'Group': denied" -and $Level -eq 'ERROR' }
        }

        It 'does not send writes with -WhatIf' {
            Add-EntraGroupOwner -GroupName 'Group' -OwnerIdentifier 'a@contoso.com' -WhatIf

            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
        }
    }

    Context 'Remove-EntraGroupOwner' {
        It 'removes listed owners with one resolve batch and one DELETE batch' {
            $result = @(Remove-EntraGroupOwner -GroupName 'Group' -OwnerIdentifier 'a@contoso.com', 'b@contoso.com' -PassThru -Confirm:$false)

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            $script:ownerBatchSizes[1] | Should -Be @('DELETE /groups/group-id/owners/id-a%40contoso.com/$ref', 'DELETE /groups/group-id/owners/id-b%40contoso.com/$ref')
            $result.Status | Should -Be @('Removed', 'Removed')
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Removed owner 'a@contoso.com' from group 'Group'." -and $Level -eq 'SUCCESS' }
        }

        It 'reports NotFound and Failed with the original messages' {
            $script:ownerWriteAnswer = {
                param($request)
                if ($request.url -like '*id-gone%40contoso.com*') {
                    return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'Resource does not exist' } } }
                }
                @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
            }

            $result = @(Remove-EntraGroupOwner -GroupName 'Group' -OwnerIdentifier 'gone@contoso.com', 'bad@contoso.com' -PassThru -Confirm:$false)

            $result.Status | Should -Be @('NotFound', 'Failed')
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Owner 'gone@contoso.com' is not an owner of 'Group'." -and $Level -eq 'WARNING' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Failed to remove owner 'bad@contoso.com' from 'Group': denied" -and $Level -eq 'ERROR' }
        }
    }

    Context 'Copy-EntraGroupOwner' {
        BeforeEach {
            Mock Get-MgGroup { [pscustomobject]@{ Id = "id-$GroupId"; DisplayName = "Group $GroupId" } } -ParameterFilter { $GroupId }
            $script:ownerDirectAnswer = {
                param($uri)
                if ($uri -like '*/groups/id-src/owners*') {
                    return @{ value = @(@{ id = 'o1'; userPrincipalName = 'o1@contoso.com' }, @{ id = 'o2'; userPrincipalName = 'o2@contoso.com' }, @{ id = 'o3'; userPrincipalName = 'o3@contoso.com' }) }
                }
                @{ value = @(@{ id = 'o2' }) }
            }
        }

        It 'sends one batch with only the missing owners' {
            $result = @(Copy-EntraGroupOwner -SourceGroupId 'src' -DestinationGroupId 'dst' -PassThru -Confirm:$false)

            $script:ownerBatchSizes.Count | Should -Be 1
            $script:ownerBatchSizes[0] | Should -Be @('POST /groups/id-dst/owners/$ref', 'POST /groups/id-dst/owners/$ref')
            $result.Status | Should -Be @('Added', 'Exists', 'Added')
            ($result[0].PSObject.Properties.Name -join ',') | Should -Be 'SourceGroup,DestinationGroup,OwnerName,OwnerId,Status'
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Copied owner 'o1@contoso.com' to 'Group dst'." -and $Level -eq 'SUCCESS' }
        }

        It 'reports Exists and Failed returned by the batch' {
            $script:ownerWriteAnswer = {
                param($request)
                if ($request.body.'@odata.id' -like '*/o1') {
                    return @{ status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'object references already exist' } } }
                }
                @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
            }

            $result = @(Copy-EntraGroupOwner -SourceGroupId 'src' -DestinationGroupId 'dst' -PassThru -Confirm:$false)

            $result.Status | Should -Be @('Exists', 'Exists', 'Failed')
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Owner 'o1@contoso.com' is already an owner of 'Group dst'." -and $Level -eq 'WARNING' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Failed to copy owner 'o3@contoso.com' to 'Group dst': denied" -and $Level -eq 'ERROR' }
        }

        It 'does not send writes with -WhatIf' {
            Copy-EntraGroupOwner -SourceGroupId 'src' -DestinationGroupId 'dst' -WhatIf

            $script:ownerBatchSizes.Count | Should -Be 0
        }
    }

    Context 'Copy-EntraGroup' {
        BeforeEach {
            Mock Get-MgGroup {
                $gid = $GroupId -replace '^id-', ''
                [pscustomobject]@{ Id = "id-$gid"; DisplayName = "Group $gid"; Description = $null; GroupTypes = @(); MailEnabled = $false; SecurityEnabled = $true; OnPremisesSyncEnabled = $false; IsAssignableToRole = $false }
            } -ParameterFilter { $GroupId }
            Mock Get-MgGroupMember {
                if ($GroupId -eq 'id-src') {
                    1..25 | ForEach-Object { [pscustomobject]@{ Id = "m$_"; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user'; displayName = "User $_" } } }
                }
            }
            $script:ownerDirectAnswer = { param($uri) @{ value = @() } }
        }

        It 'sends 25 member adds as 2 batches of 20 and 5' {
            $result = Copy-EntraGroup -SourceGroupId 'src' -DestinationGroupId 'dst' -SkipOwners -PassThru -Confirm:$false

            $script:ownerBatchSizes.Count | Should -Be 2
            $script:ownerBatchSizes[0].Count | Should -Be 20
            $script:ownerBatchSizes[1].Count | Should -Be 5
            $script:ownerBatchSizes[0][0] | Should -Be 'POST /groups/id-dst/members/$ref'
            $result.MembersCopied | Should -Be 25
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Copied User 'User 1' to 'Group dst'." -and $Level -eq 'SUCCESS' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing 25 member(s) in Graph batches (20 per request) ...' -and $Level -eq 'INFO' }
        }

        It 'counts Exists as skipped and reports failures with the original messages' {
            $script:ownerWriteAnswer = {
                param($request)
                if ($request.body.'@odata.id' -like '*/m1') {
                    return @{ status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'object references already exist' } } }
                }
                if ($request.body.'@odata.id' -like '*/m2') {
                    return @{ status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'denied' } } }
                }
                @{ status = 204 }
            }

            $result = Copy-EntraGroup -SourceGroupId 'src' -DestinationGroupId 'dst' -SkipOwners -PassThru -Confirm:$false

            $result.MembersCopied | Should -Be 23
            $result.MembersSkipped | Should -Be 1
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "User 'User 1' is already a member of 'Group dst'." -and $Level -eq 'WARNING' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Failed to copy User 'User 2' to 'Group dst': denied" -and $Level -eq 'ERROR' }
        }

        It 'batches owner adds' {
            $script:ownerDirectAnswer = {
                param($uri)
                if ($uri -like '*/groups/id-src/owners*') { return @{ value = @(@{ id = 'o1'; userPrincipalName = 'o1@contoso.com' }, @{ id = 'o2'; userPrincipalName = 'o2@contoso.com' }) } }
                @{ value = @() }
            }

            $result = Copy-EntraGroup -SourceGroupId 'src' -DestinationGroupId 'dst' -SkipMembers -PassThru -Confirm:$false

            $script:ownerBatchSizes.Count | Should -Be 1
            $script:ownerBatchSizes[0] | Should -Be @('POST /groups/id-dst/owners/$ref', 'POST /groups/id-dst/owners/$ref')
            $result.OwnersCopied | Should -Be 2
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Copied owner 'o1@contoso.com' to 'Group dst'." -and $Level -eq 'SUCCESS' }
        }
    }
}

Describe 'Entra group read batching' {
    BeforeAll {
        function Set-ProgressAndInfoPreferences {}
        function Restore-ProgressAndInfoPreferences {}
        function Get-NCProgressPercent { param($Current, $Total) 0 }
        function Get-MgGroupMember {
            param(
                [string]$GroupId,
                [switch]$All
            )
        }
        function Out-GridView {
            param(
                [Parameter(ValueFromPipeline = $true)]
                [object]$InputObject,
                [string]$Title
            )
            process {}
        }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Get-MgContext { [pscustomobject]@{ Environment = 'Global' } }
        Mock Find-UserRecipient { $null }
    }

    Context 'Get-EntraGroupUser' {
        BeforeEach {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    if ($request.url -like '/users/*/memberOf*') {
                        return @{ status = 200; body = @{ value = @(
                                    @{ '@odata.type' = '#microsoft.graph.group'; id = 'g2'; displayName = 'Zulu'; mail = 'zulu@contoso.com'; groupTypes = @('Unified'); description = 'Z group'; mailNickname = 'zulu'; mailEnabled = $true }
                                    @{ '@odata.type' = '#microsoft.graph.group'; id = 'g1'; displayName = 'Alpha'; mail = $null; groupTypes = @(); description = 'A group'; mailNickname = 'alpha'; mailEnabled = $false }
                                ) } }
                    }
                    if ($request.url -like '/users/*') {
                        $upn = [uri]::UnescapeDataString(($request.url -replace '^/users/([^?]+).*$', '$1'))
                        return @{ status = 200; body = @{ id = "id-$upn"; userPrincipalName = $upn; displayName = "Name $upn" } }
                    }
                    @{ status = 500 }
                }
            }
        }

        It 'reads 14 users with one resolve batch and one memberOf batch' {
            $users = @(1..14 | ForEach-Object { "user$_@contoso.com" })

            $users | Get-EntraGroupUser | Out-Null

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing users in Graph batches (20 per request) ...' -and $Level -eq 'INFO' }
        }

        It 'emits the same rows, sorted by group name, for each user in input order' {
            $result = @('a@contoso.com', 'b@contoso.com' | Get-EntraGroupUser)

            $result.Count | Should -Be 4
            $result[0].PSObject.Properties.Name | Should -Be @('Group Name', 'Group Mail')
            $result[0].'Group Name' | Should -Be 'Alpha'
            $result[0].'Group Mail' | Should -BeNullOrEmpty
            $result[1].'Group Name' | Should -Be 'Zulu'
            $result[1].'Group Mail' | Should -Be 'zulu@contoso.com'
        }

        It 'emits the extra GridView columns in the original order' {
            $script:gridRows = @()
            Mock Out-GridView { $script:gridRows += $InputObject }

            Get-EntraGroupUser -UserIdentifier 'a@contoso.com' -GridView

            $script:gridRows[0].PSObject.Properties.Name | Should -Be @('Group Name', 'Group Mail', 'Group Description', 'Group Mail Nickname', 'Group Mail Enabled', 'Group Type', 'Group ID')
            $zulu = $script:gridRows | Where-Object { $_.'Group Name' -eq 'Zulu' }
            $zulu.'Group Description' | Should -Be 'Z group'
            $zulu.'Group Type' | Should -Be 'Unified'
            $zulu.'Group ID' | Should -Be 'g2'
        }

        It 'reports a single connection error and skips all work when Graph is unavailable' {
            Mock Test-MgGraphConnection { $false }

            @('a@contoso.com', 'b@contoso.com') | Get-EntraGroupUser

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "Can't connect or use Microsoft Graph modules. Please check logs." -and $Level -eq 'ERROR' }
            Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -Scope It
        }

        It 'falls back to the display-name filter once for an unresolved input' {
            Mock Get-MgUser { $null }
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } }
                }
            }

            Get-EntraGroupUser -UserIdentifier 'Nobody Here'

            Should -Invoke Get-MgUser -Times 1 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq "User 'Nobody Here' not found" -and $Level -eq 'WARNING' }
        }

        It 'reports a missing object ID with the original message' {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'gone' } } }
                }
            }

            Get-EntraGroupUser -UserIdentifier '11111111-1111-1111-1111-111111111111'

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -like "Entra user with ID '11111111-1111-1111-1111-111111111111' not found: *" -and $Level -eq 'ERROR' }
        }
    }

    Context 'Get-EntraGroupDevice' {
        BeforeEach {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    if ($request.url -like '/devices/*/memberOf*') {
                        return @{ status = 200; body = @{ value = @(@{ '@odata.type' = '#microsoft.graph.group'; id = 'g1'; displayName = 'Device Group'; mail = 'dg@contoso.com' }) } }
                    }
                    $decoded = [uri]::UnescapeDataString($request.url)
                    $name = $decoded -replace "^/devices\?\`$filter=displayName eq '([^']+)'.*$", '$1'
                    @{ status = 200; body = @{ value = @(@{ id = "id-$name"; displayName = $name }) } }
                }
            }
        }

        It 'reads 3 devices by name with one resolve batch and one memberOf batch' {
            $result = @('PC1', 'PC2', 'PC3' | Get-EntraGroupDevice)

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            $result.Count | Should -Be 3
            $result[0].PSObject.Properties.Name | Should -Be @('Group Name', 'Group Mail')
            $result[0].'Group Name' | Should -Be 'Device Group'
            $result[0].'Group Mail' | Should -Be 'dg@contoso.com'
        }

        It 'reports a missing device ID with the original message' {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'gone' } } }
                }
            }

            Get-EntraGroupDevice -DeviceIdentifier '22222222-2222-2222-2222-222222222222'

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -like "Entra device with ID '22222222-2222-2222-2222-222222222222' not found: *" -and $Level -eq 'ERROR' }
        }
    }

    Context 'Get-EntraGroupMembers' {
        It 'reads registered owners and users of 3 device members in one batch of 6 sub-requests' {
            Mock Get-MgGroup { [pscustomobject]@{ Id = 'group-id'; DisplayName = 'Group' } }
            Mock Get-MgGroupMember {
                1..3 | ForEach-Object {
                    [pscustomobject]@{ Id = "dev$_"; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.device'; displayName = "PC$_" } }
                }
            }
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    if ($request.url -like '*/registeredOwners') { return @{ status = 200; body = @{ value = @(@{ id = 'u1'; userPrincipalName = 'owner@contoso.com'; displayName = 'Owner' }) } } }
                    @{ status = 200; body = @{ value = @(@{ id = 'u2'; displayName = 'Only Name' }) } }
                }
            }

            $result = @(Get-EntraGroupMembers -GroupName 'Group' -IncludeDeviceUsers)

            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter { @(($Body | ConvertFrom-Json).requests).Count -eq 6 }
            $result.Count | Should -Be 3
            $result[0].PSObject.Properties.Name | Should -Be @('Member Name', 'Member Type', 'Member Id', 'Device Owners/Users')
            $result[0].'Device Owners/Users' | Should -Be 'Owners: owner@contoso.com | Users: Only Name'
        }
    }

    Context 'Export-EmptyEntraGroups' {
        It 'checks 45 groups with 3 batch calls and reports only the empty ones' {
            Mock Get-MgGroup {
                1..45 | ForEach-Object {
                    [pscustomobject]@{ Id = "gid$_"; DisplayName = "Group $_"; GroupTypes = @(); SecurityEnabled = $true; MailEnabled = $false }
                }
            }
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $number = [int]($request.url -replace '^/groups/gid(\d+)/.*$', '$1')
                    if ($number % 5 -eq 0) { return @{ status = 200; body = @{ value = @() } } }
                    @{ status = 200; body = @{ value = @(@{ id = 'member' }) } }
                }
            }

            $result = @(Export-EmptyEntraGroups -Csv $false)

            Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
            $result.Count | Should -Be 9
            $result[0].PSObject.Properties.Name | Should -Be @('DisplayName', 'Id', 'GroupType', 'MemberCount', 'MailEnabled', 'SecurityEnabled')
            $result.Id | Should -Contain 'gid45'
            $result.Id | Should -Not -Contain 'gid44'
            $result[0].GroupType | Should -Be 'Security'
        }
    }
}
