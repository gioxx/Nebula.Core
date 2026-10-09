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
    function Invoke-MgGraphRequest {
        param(
            [string]$Uri,
            [string]$Method,
            [object]$Body,
            [string]$ContentType
        )
    }
    function Invoke-NCGraphAllPagesCore {
        param(
            [string]$Uri,
            [int]$DelayMs
        )
    }
    function Get-NCGraphDirectoryObjectUri {
        param([string]$Id)
        "https://graph.example/v1.0/directoryObjects/$Id"
    }
    function Get-NCEnterpriseApplicationSnapshot {
        param(
            [string]$ApplicationName,
            [string]$ApplicationId,
            [switch]$IncludeAppRoleAssignments
        )
    }
    function Set-NCEnterpriseApplicationFromSnapshot {
        param(
            [object]$Snapshot,
            [string]$TargetDisplayName,
            [switch]$IncludeAppRoleAssignments
        )
    }
    function Compare-NCEnterpriseApplicationSnapshot {
        param(
            [object]$ReferenceSnapshot,
            [object]$DifferenceSnapshot,
            [switch]$IncludeAppRoleAssignments
        )
    }

    . "$PSScriptRoot/../../Private/NC-Hlp.EnterpriseApplications.ps1"
    . "$PSScriptRoot/../../Public/NC.EnterpriseApplications.ps1"
}

Describe 'Get-NCEnterpriseApplicationSnapshot' {
    BeforeAll {
        $app = [pscustomobject]@{
            id                 = 'app-id-1'
            appId              = 'client-id-1'
            displayName        = 'Contoso Test App'
            signInAudience     = 'AzureADMyOrg'
            identifierUris     = @('api://contoso-test-app')
            notes              = 'test notes'
            tags               = @('tag1')
            web                = [pscustomobject]@{ redirectUris = @('https://localhost/callback') }
            spa                = [pscustomobject]@{ redirectUris = @() }
            publicClient       = [pscustomobject]@{ redirectUris = @() }
            requiredResourceAccess = @()
            appRoles           = @()
            api                = [pscustomobject]@{ oauth2PermissionScopes = @() }
            passwordCredentials = @(
                [pscustomobject]@{ displayName = 'secret1'; keyId = 'kid-1'; endDateTime = '2027-01-01T00:00:00Z' }
            )
            keyCredentials      = @()
        }
        $sp = [pscustomobject]@{
            id          = 'sp-id-1'
            appId       = 'client-id-1'
            displayName = 'Contoso Test App'
            tags        = @()
            homepage    = $null
            logoUrl     = $null
        }
        $owner = [pscustomobject]@{ id = 'owner-1'; displayName = 'Jane Doe'; userPrincipalName = 'jane@contoso.com' }
        $assignment = [pscustomobject]@{ principalId = 'principal-1'; principalDisplayName = 'Some Group'; principalType = 'Group'; appRoleId = 'role-1' }
    }

    BeforeEach {
        Mock Write-NCMessage {}
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '/applications\?') {
                return [pscustomobject]@{ value = @($app) }
            }
            if ($Uri -match '/servicePrincipals\?') {
                return [pscustomobject]@{ value = @($sp) }
            }
            throw "Unexpected Uri: $Uri"
        }
        Mock Invoke-NCGraphAllPagesCore {
            if ($Uri -match '/owners') { return @($owner) }
            if ($Uri -match '/appRoleAssignedTo') { return @($assignment) }
            return @()
        }
    }

    It 'builds a normalized snapshot from an application looked up by name' {
        $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App'

        $snapshot.SchemaVersion | Should -Be 1
        $snapshot.Application.DisplayName | Should -Be 'Contoso Test App'
        $snapshot.Application.Owners.Count | Should -Be 1
        $snapshot.Application.Owners[0].UserPrincipalName | Should -Be 'jane@contoso.com'
        $snapshot.ServicePrincipal.AppId | Should -Be 'client-id-1'
        $snapshot.CredentialsMetadata.PasswordCredentials[0].KeyId | Should -Be 'kid-1'
        $snapshot.AppRoleAssignments.Count | Should -Be 0
    }

    It 'captures the writable api settings besides the permission scopes' {
        $originalApi = $app.api
        $app.api = [pscustomobject]@{
            oauth2PermissionScopes      = @([pscustomobject]@{ id = 'scope-1'; value = 'access_as_user' })
            acceptMappedClaims          = $true
            knownClientApplications     = @('client-app-1')
            preAuthorizedApplications   = @([pscustomobject]@{ appId = 'client-app-1'; delegatedPermissionIds = @('scope-1') })
            requestedAccessTokenVersion = 2
        }

        try {
            $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App'
        }
        finally {
            $app.api = $originalApi
        }

        $snapshot.Application.Oauth2PermissionScopes[0].value | Should -Be 'access_as_user'
        $snapshot.Application.Api.AcceptMappedClaims | Should -BeTrue
        $snapshot.Application.Api.KnownClientApplications | Should -Be @('client-app-1')
        $snapshot.Application.Api.PreAuthorizedApplications[0].appId | Should -Be 'client-app-1'
        $snapshot.Application.Api.RequestedAccessTokenVersion | Should -Be 2
    }

    It 'returns no snapshot when the owners cannot be read completely' {
        Mock Invoke-NCGraphAllPagesCore {
            if ($Uri -match '/owners') { throw 'Forbidden' }
            return @()
        }

        $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App'

        $snapshot | Should -BeNullOrEmpty
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*owners*Forbidden*' }
    }

    It 'returns no snapshot when the App Role Assignments cannot be read completely' {
        Mock Invoke-NCGraphAllPagesCore {
            if ($Uri -match '/appRoleAssignedTo') { throw 'Forbidden' }
            return @()
        }

        $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App' -IncludeAppRoleAssignments

        $snapshot | Should -BeNullOrEmpty
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*App Role Assignments*Forbidden*' }
    }

    It 'captures token claim settings and the assignment requirement' {
        $app | Add-Member -NotePropertyName groupMembershipClaims -NotePropertyValue 'SecurityGroup' -Force
        $app | Add-Member -NotePropertyName optionalClaims -NotePropertyValue ([pscustomobject]@{ idToken = @([pscustomobject]@{ name = 'email' }) }) -Force
        $sp | Add-Member -NotePropertyName appRoleAssignmentRequired -NotePropertyValue $true -Force
        try {
            $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App'
        }
        finally {
            $app.PSObject.Properties.Remove('groupMembershipClaims'); $app.PSObject.Properties.Remove('optionalClaims'); $sp.PSObject.Properties.Remove('appRoleAssignmentRequired')
        }

        $snapshot.Application.GroupMembershipClaims | Should -Be 'SecurityGroup'
        $snapshot.Application.OptionalClaims.idToken[0].name | Should -Be 'email'
        $snapshot.ServicePrincipal.AppRoleAssignmentRequired | Should -BeTrue
    }
    It 'captures the Service Principal owners and enabled state' {
        $sp | Add-Member -NotePropertyName accountEnabled -NotePropertyValue $false -Force
        try {
            $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App'
        }
        finally {
            $sp.PSObject.Properties.Remove('accountEnabled')
        }

        $snapshot.ServicePrincipal.AccountEnabled | Should -BeFalse
        $snapshot.ServicePrincipal.Owners[0].Id | Should -Be 'owner-1'
        Assert-MockCalled Invoke-NCGraphAllPagesCore -Times 1 -Exactly -Scope It -ParameterFilter { $Uri -like 'v1.0/servicePrincipals/sp-id-1/owners*' }
    }

    It 'returns no snapshot when the Service Principal owners cannot be read' {
        Mock Invoke-NCGraphAllPagesCore {
            if ($Uri -match '/servicePrincipals/.+/owners') { throw 'Forbidden' }
            return @()
        }

        Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App' | Should -BeNullOrEmpty
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*Service Principal owners*' }
    }
    It 'captures the fallback public client flag and the SSO mode' {
        $app | Add-Member -NotePropertyName isFallbackPublicClient -NotePropertyValue $true -Force
        $sp | Add-Member -NotePropertyName preferredSingleSignOnMode -NotePropertyValue 'saml' -Force
        try {
            $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App'
        }
        finally {
            $app.PSObject.Properties.Remove('isFallbackPublicClient'); $sp.PSObject.Properties.Remove('preferredSingleSignOnMode')
        }

        $snapshot.Application.IsFallbackPublicClient | Should -BeTrue
        $snapshot.ServicePrincipal.PreferredSingleSignOnMode | Should -Be 'saml'
    }
    It 'includes App Role Assignments only when requested' {
        $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App' -IncludeAppRoleAssignments

        $snapshot.AppRoleAssignments.Count | Should -Be 1
        $snapshot.AppRoleAssignments[0].PrincipalDisplayName | Should -Be 'Some Group'
    }

    It 'URI-encodes the display name filter so names with & or # resolve' {
        $null = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'R&D Portal #2'

        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter {
            $Uri -like 'v1.0/applications?$filter=*' -and $Uri -like '*R%26D%20Portal%20%232*' -and $Uri -notlike '*R&D*'
        }
    }
    It 'returns nothing and logs an error when the application is not found' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '/applications\?') {
                return [pscustomobject]@{ value = @() }
            }
            throw "Unexpected Uri: $Uri"
        }

        $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Missing App'

        $snapshot | Should -BeNullOrEmpty
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }
}

Describe 'Set-NCEnterpriseApplicationFromSnapshot' {
    BeforeAll {
        $snapshot = [pscustomobject]@{
            Application        = [pscustomobject]@{
                DisplayName            = 'Source App'
                SignInAudience         = 'AzureADMyOrg'
                IdentifierUris         = @()
                Notes                  = $null
                Tags                   = @()
                Web                    = [pscustomobject]@{ redirectUris = @('https://localhost/callback') }
                Spa                    = [pscustomobject]@{ redirectUris = @() }
                PublicClient           = [pscustomobject]@{ redirectUris = @() }
                RequiredResourceAccess = @()
                AppRoles               = @()
                Oauth2PermissionScopes = @()
                Owners                 = @([pscustomobject]@{ Id = 'owner-1'; DisplayName = 'Jane Doe'; UserPrincipalName = 'jane@contoso.com' })
            }
            ServicePrincipal   = [pscustomobject]@{ Tags = @('WindowsAzureActiveDirectoryIntegratedApp'); Homepage = $null; LogoUrl = $null }
            AppRoleAssignments = @([pscustomobject]@{ PrincipalId = 'principal-1'; PrincipalDisplayName = 'Some Group'; PrincipalType = 'Group'; AppRoleId = 'role-1' })
        }
    }

    BeforeEach {
        Mock Write-NCMessage {}
    }

    It 'creates the destination application and service principal when none exists' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') {
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') {
                return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' }
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.Created | Should -Be $true
        $result.TargetApplicationId | Should -Be 'new-app-id'
        $result.OwnersAdded | Should -Be 1
        $result.AssignmentsAdded | Should -Be 0
        $result.AssignmentsFailed | Should -Be 0
        $result.Error | Should -BeNullOrEmpty

        Assert-MockCalled Invoke-MgGraphRequest -Scope It -ParameterFilter {
            $Method -eq 'POST' -and $Uri -eq 'v1.0/applications'
        }
        Assert-MockCalled Invoke-MgGraphRequest -Scope It -ParameterFilter {
            $Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals' -and
            $Body -match 'WindowsAzureActiveDirectoryIntegratedApp'
        }
        # Owner references use the active cloud's directoryObjects URI
        Assert-MockCalled Invoke-MgGraphRequest -Scope It -ParameterFilter {
            $Method -eq 'POST' -and $Uri -eq 'v1.0/applications/new-app-id/owners/$ref' -and
            $Body -match 'https://graph\.example/v1\.0/directoryObjects/'
        }
    }

    It 'writes the snapshot api settings to the destination application' {
        $apiSnapshot = $snapshot.PSObject.Copy()
        $apiSnapshot.Application = $snapshot.Application.PSObject.Copy()
        $apiSnapshot.Application | Add-Member -NotePropertyName Api -NotePropertyValue ([pscustomobject]@{
                AcceptMappedClaims          = $true
                KnownClientApplications     = @('client-app-1')
                PreAuthorizedApplications   = @([pscustomobject]@{ appId = 'client-app-1'; delegatedPermissionIds = @('scope-1') })
                RequestedAccessTokenVersion = 2
            }) -Force
        $script:appPost = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') {
                $script:appPost = $Body | ConvertFrom-Json
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $apiSnapshot -TargetDisplayName 'Target App' -Confirm:$false

        $script:appPost.api.acceptMappedClaims | Should -BeTrue
        $script:appPost.api.knownClientApplications | Should -Be @('client-app-1')
        $script:appPost.api.preAuthorizedApplications[0].appId | Should -Be 'client-app-1'
        $script:appPost.api.requestedAccessTokenVersion | Should -Be 2
    }

    It 'sends null api settings so an update clears the destination values' {
        $apiSnapshot = $snapshot.PSObject.Copy()
        $apiSnapshot.Application = $snapshot.Application.PSObject.Copy()
        $apiSnapshot.Application | Add-Member -NotePropertyName Api -NotePropertyValue ([pscustomobject]@{
                AcceptMappedClaims          = $null
                KnownClientApplications     = @()
                PreAuthorizedApplications   = @()
                RequestedAccessTokenVersion = $null
            }) -Force
        $script:appPatch = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App'; appRoles = @(); api = [pscustomobject]@{ oauth2PermissionScopes = @(); requestedAccessTokenVersion = 2; acceptMappedClaims = $true } }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -eq 'v1.0/applications/existing-app-id') { $script:appPatch = $Body | ConvertFrom-Json; return $null }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $apiSnapshot -TargetDisplayName 'Target App' -Confirm:$false

        $apiNames = @($script:appPatch.api.PSObject.Properties.Name)
        $apiNames | Should -Contain 'acceptMappedClaims'
        $apiNames | Should -Contain 'requestedAccessTokenVersion'
        $script:appPatch.api.acceptMappedClaims | Should -BeNullOrEmpty
        $script:appPatch.api.requestedAccessTokenVersion | Should -BeNullOrEmpty
    }
    It 'restores token claim settings and the assignment requirement' {
        $claimSnapshot = $snapshot.PSObject.Copy()
        $claimSnapshot.Application = $snapshot.Application.PSObject.Copy()
        $claimSnapshot.ServicePrincipal = $snapshot.ServicePrincipal.PSObject.Copy()
        $claimSnapshot.Application | Add-Member -NotePropertyName GroupMembershipClaims -NotePropertyValue 'SecurityGroup' -Force
        $claimSnapshot.Application | Add-Member -NotePropertyName OptionalClaims -NotePropertyValue ([pscustomobject]@{ idToken = @([pscustomobject]@{ name = 'email' }) }) -Force
        $claimSnapshot.ServicePrincipal | Add-Member -NotePropertyName AppRoleAssignmentRequired -NotePropertyValue $true -Force
        $script:appPost = $null; $script:spPost = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { $script:appPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $claimSnapshot -TargetDisplayName 'Target App' -Confirm:$false

        $script:appPost.groupMembershipClaims | Should -Be 'SecurityGroup'
        $script:appPost.optionalClaims.idToken[0].name | Should -Be 'email'
        $script:spPost.appRoleAssignmentRequired | Should -BeTrue
    }

    It 'restores the Service Principal enabled state and adds its owners' {
        $spSnapshot = $snapshot.PSObject.Copy()
        $spSnapshot.ServicePrincipal = $snapshot.ServicePrincipal.PSObject.Copy()
        $spSnapshot.ServicePrincipal | Add-Member -NotePropertyName AccountEnabled -NotePropertyValue $false -Force
        $spSnapshot.ServicePrincipal | Add-Member -NotePropertyName Owners -NotePropertyValue @([pscustomobject]@{ Id = 'sp-owner-1'; DisplayName = 'SP Owner'; UserPrincipalName = 'spowner@contoso.com' }) -Force
        $script:spPost = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $spSnapshot -TargetDisplayName 'Target App' -Confirm:$false

        $script:spPost.accountEnabled | Should -BeFalse
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter { $Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals/new-sp-id/owners/$ref' -and $Body -match 'sp-owner-1' }
        $result.OwnersAdded | Should -Be 2
    }
    It 'restores the fallback public client flag' {
        $flagSnapshot = $snapshot.PSObject.Copy()
        $flagSnapshot.Application = $snapshot.Application.PSObject.Copy()
        $flagSnapshot.Application | Add-Member -NotePropertyName IsFallbackPublicClient -NotePropertyValue $true -Force
        $script:appPost = $null; $script:spPost = $null; $script:ownerPostFails = $false
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { $script:appPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            if ($Method -eq 'POST' -and $Uri -like 'v1.0/*/owners/$ref') { if ($script:ownerPostFails) { throw 'Request_ResourceNotFound: the object does not exist' } ; return $null }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $flagSnapshot -TargetDisplayName 'Target App' -Confirm:$false

        $script:appPost.isFallbackPublicClient | Should -BeTrue
    }

    It 'warns that SAML or password SSO is not cloned and does not write the SSO mode' {
        $ssoSnapshot = $snapshot.PSObject.Copy()
        $ssoSnapshot.ServicePrincipal = $snapshot.ServicePrincipal.PSObject.Copy()
        $ssoSnapshot.ServicePrincipal | Add-Member -NotePropertyName PreferredSingleSignOnMode -NotePropertyValue 'saml' -Force
        $script:appPost = $null; $script:spPost = $null; $script:ownerPostFails = $false
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { $script:appPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            if ($Method -eq 'POST' -and $Uri -like 'v1.0/*/owners/$ref') { if ($script:ownerPostFails) { throw 'Request_ResourceNotFound: the object does not exist' } ; return $null }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $ssoSnapshot -TargetDisplayName 'Target App' -Confirm:$false

        @($script:spPost.PSObject.Properties.Name) | Should -Not -Contain 'preferredSingleSignOnMode'
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -like '*saml*not copied*' }
    }

    It 'sends empty web URLs and implicit grant settings so an update can clear them' {
        $script:appPost = $null; $script:spPost = $null; $script:ownerPostFails = $false
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { $script:appPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            if ($Method -eq 'POST' -and $Uri -like 'v1.0/*/owners/$ref') { if ($script:ownerPostFails) { throw 'Request_ResourceNotFound: the object does not exist' } ; return $null }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $webNames = @($script:appPost.web.PSObject.Properties.Name)
        $webNames | Should -Contain 'homePageUrl'
        $webNames | Should -Contain 'logoutUrl'
        $script:appPost.web.homePageUrl | Should -BeNullOrEmpty
        $script:appPost.web.implicitGrantSettings.enableAccessTokenIssuance | Should -BeFalse
        $script:appPost.web.implicitGrantSettings.enableIdTokenIssuance | Should -BeFalse
    }

    It 'counts owners that could not be added' {
        $script:appPost = $null; $script:spPost = $null; $script:ownerPostFails = $true
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { $script:appPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            if ($Method -eq 'POST' -and $Uri -like 'v1.0/*/owners/$ref') { if ($script:ownerPostFails) { throw 'Request_ResourceNotFound: the object does not exist' } ; return $null }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.OwnersFailed | Should -Be 1
        $result.OwnersAdded | Should -Be 0
    }
    It 'leaves token claims and the assignment requirement alone for older snapshots' {
        $script:appPost = $null; $script:spPost = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { $script:appPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { $script:spPost = $Body | ConvertFrom-Json; return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        @($script:appPost.PSObject.Properties.Name) | Should -Not -Contain 'groupMembershipClaims'
        @($script:appPost.PSObject.Properties.Name) | Should -Not -Contain 'optionalClaims'
        @($script:spPost.PSObject.Properties.Name) | Should -Not -Contain 'appRoleAssignmentRequired'
        @($script:spPost.PSObject.Properties.Name) | Should -Not -Contain 'accountEnabled'
        @($script:appPost.PSObject.Properties.Name) | Should -Not -Contain 'isFallbackPublicClient'
        Assert-MockCalled Invoke-MgGraphRequest -Times 0 -Scope It -ParameterFilter { $Uri -like 'v1.0/servicePrincipals/*/owners/*' }
    }
    It 'still accepts snapshots saved without api settings' {
        $script:appPost = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') {
                $script:appPost = $Body | ConvertFrom-Json
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.Error | Should -BeNullOrEmpty
        @($script:appPost.api.PSObject.Properties.Name) | Should -Be @('oauth2PermissionScopes')
    }

    It 'updates an existing destination application instead of creating a new one' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App' }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -match '/applications/existing-app-id$') { return $null }
            if ($Uri -match '/servicePrincipals\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -match '/servicePrincipals/existing-sp-id$') { return $null }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.Created | Should -Be $false
        $result.TargetApplicationId | Should -Be 'existing-app-id'
        $result.AssignmentsFailed | Should -Be 0
        $result.Error | Should -BeNullOrEmpty
        Assert-MockCalled Invoke-MgGraphRequest -Scope It -ParameterFilter {
            $Method -eq 'PATCH' -and $Uri -eq 'v1.0/applications/existing-app-id'
        }
        # The snapshot has no homepage: the update clears any value left on the destination
        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter {
            $Method -eq 'PATCH' -and $Uri -eq 'v1.0/servicePrincipals/existing-sp-id' -and
            ($Body | ConvertFrom-Json).PSObject.Properties.Name -contains 'homepage' -and $null -eq ($Body | ConvertFrom-Json).homepage
        }
    }

    It 'disables obsolete app roles and scopes in a first update before replacing them' {
        $script:appPatches = [System.Collections.Generic.List[object]]::new()
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{
                            id       = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App'
                            appRoles = @([pscustomobject]@{ id = 'role-old'; value = 'Old.Role'; isEnabled = $true })
                            api      = [pscustomobject]@{
                                requestedAccessTokenVersion = 2
                                oauth2PermissionScopes      = @([pscustomobject]@{ id = 'scope-old'; value = 'old_scope'; isEnabled = $true })
                                preAuthorizedApplications   = @([pscustomobject]@{ appId = 'client-app-1'; delegatedPermissionIds = @('scope-old') })
                            }
                        }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -eq 'v1.0/applications/existing-app-id') { $script:appPatches.Add(($Body | ConvertFrom-Json)); return $null }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.Error | Should -BeNullOrEmpty
        $script:appPatches.Count | Should -Be 2
        $staging = $script:appPatches[0]
        ($staging.appRoles | Where-Object { $_.id -eq 'role-old' }).isEnabled | Should -BeFalse
        ($staging.api.oauth2PermissionScopes | Where-Object { $_.id -eq 'scope-old' }).isEnabled | Should -BeFalse
        @($staging.api.preAuthorizedApplications).Count | Should -Be 0
        $staging.api.requestedAccessTokenVersion | Should -Be 2
        @($script:appPatches[1].appRoles).Count | Should -Be 0
    }

    It 'sends a single update when no enabled role or scope would be removed' {
        $script:appPatches = [System.Collections.Generic.List[object]]::new()
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{
                            id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App'
                            appRoles = @([pscustomobject]@{ id = 'role-off'; isEnabled = $false })
                            api = [pscustomobject]@{ oauth2PermissionScopes = @() }
                        }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -eq 'v1.0/applications/existing-app-id') { $script:appPatches.Add(($Body | ConvertFrom-Json)); return $null }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $script:appPatches.Count | Should -Be 1
    }
    It 'stops and reports an error when the Service Principal update fails' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App'; appRoles = @(); api = [pscustomobject]@{ oauth2PermissionScopes = @() } }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -eq 'v1.0/applications/existing-app-id') { return $null }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) } }
            if ($Method -eq 'PATCH' -and $Uri -eq 'v1.0/servicePrincipals/existing-sp-id') { throw 'ServiceUnavailable' }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.Error | Should -BeLike '*Service Principal*ServiceUnavailable*'
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*Service Principal*' }
        Assert-MockCalled Invoke-NCGraphAllPagesCore -Times 0 -Scope It
    }
    It 'does not send an empty homepage when creating a Service Principal' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') { return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' } }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') { return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' } }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $null = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Invoke-MgGraphRequest -Times 1 -Exactly -Scope It -ParameterFilter {
            $Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals' -and ($Body | ConvertFrom-Json).PSObject.Properties.Name -notcontains 'homepage'
        }
    }

    It 'updates tags/homepage on an existing Service Principal so the app stays visible under Enterprise applications' {
        $spPatchCalled = $false
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App' }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -match '/applications/existing-app-id$') { return $null }
            if ($Uri -match '/servicePrincipals\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -match '/servicePrincipals/existing-sp-id$') {
                $script:spPatchCalled = $true
                $script:spPatchBody = $Body
                return $null
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false | Out-Null

        $script:spPatchCalled | Should -Be $true
        $script:spPatchBody | Should -Match 'WindowsAzureActiveDirectoryIntegratedApp'
    }

    It 'applies App Role Assignments only when requested' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App' }) }
            }
            if ($Uri -match '/servicePrincipals\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) }
            }
            if ($Method -eq 'POST' -and $Uri -match '/appRoleAssignedTo$') { return $null }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -IncludeAppRoleAssignments -Confirm:$false

        $result.AssignmentsAdded | Should -Be 1
        $result.AssignmentsFailed | Should -Be 0
        $result.Error | Should -BeNullOrEmpty
        Assert-MockCalled Invoke-MgGraphRequest -Scope It -ParameterFilter {
            $Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals/existing-sp-id/appRoleAssignedTo'
        }
    }

    It 'returns an error object without crashing when the target name is ambiguous' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{
                    value = @(
                        [pscustomobject]@{ id = 'dup-app-id-1'; appId = 'dup-client-id-1'; displayName = 'Target App' }
                        [pscustomobject]@{ id = 'dup-app-id-2'; appId = 'dup-client-id-2'; displayName = 'Target App' }
                    )
                }
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result | Should -Not -BeNullOrEmpty
        $result.Error | Should -Not -BeNullOrEmpty
        $result.Created | Should -Be $false
        $result.TargetApplicationId | Should -BeNullOrEmpty
    }

    It 'never forwards read-only redirectUriSettings alongside redirectUris' {
        $snapshotWithRedirectUriSettings = [pscustomobject]@{
            Application        = [pscustomobject]@{
                DisplayName            = 'Source App'
                SignInAudience         = 'AzureADMyOrg'
                IdentifierUris         = @()
                Notes                  = $null
                Tags                   = @()
                Web                    = [pscustomobject]@{
                    redirectUris        = @('https://localhost/callback')
                    redirectUriSettings = @([pscustomobject]@{ uri = 'https://localhost/callback'; index = $null })
                }
                Spa                    = [pscustomobject]@{ redirectUris = @() }
                PublicClient           = [pscustomobject]@{ redirectUris = @() }
                RequiredResourceAccess = @()
                AppRoles               = @()
                Oauth2PermissionScopes = @()
                Owners                 = @()
            }
            AppRoleAssignments = @()
        }
        $capturedBody = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') {
                $script:capturedBody = $Body
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') {
                return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' }
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshotWithRedirectUriSettings -TargetDisplayName 'Target App' -Confirm:$false

        $result.Created | Should -Be $true
        $script:capturedBody | Should -Not -Match 'redirectUriSettings'
        $script:capturedBody | Should -Match 'https://localhost/callback'
    }

    It 'never writes identifierUris and logs a warning when the source has them' {
        $snapshotWithUri = [pscustomobject]@{
            Application        = [pscustomobject]@{
                DisplayName            = 'Source App'
                SignInAudience         = 'AzureADMyOrg'
                IdentifierUris         = @('http://localhost:8080/saml2/service-provider-metadata/contract-manager')
                Notes                  = $null
                Tags                   = @()
                Web                    = [pscustomobject]@{ redirectUris = @() }
                Spa                    = [pscustomobject]@{ redirectUris = @() }
                PublicClient           = [pscustomobject]@{ redirectUris = @() }
                RequiredResourceAccess = @()
                AppRoles               = @()
                Oauth2PermissionScopes = @()
                Owners                 = @()
            }
            AppRoleAssignments = @()
        }
        $capturedBody = $null
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/applications') {
                $script:capturedBody = $Body
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'v1.0/servicePrincipals') {
                return [pscustomobject]@{ id = 'new-sp-id'; appId = 'new-client-id' }
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshotWithUri -TargetDisplayName 'Target App' -Confirm:$false

        $result.Created | Should -Be $true
        $script:capturedBody | Should -Not -Match 'identifierUris'
        Assert-MockCalled Write-NCMessage -Scope It -ParameterFilter {
            $Level -eq 'WARNING' -and $Message -match 'identifierUris'
        }
    }

    It 'counts a non-duplicate App Role Assignment failure as failed, not skipped, and logs at ERROR' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App' }) }
            }
            if ($Uri -match '/servicePrincipals\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) }
            }
            if ($Method -eq 'POST' -and $Uri -match '/appRoleAssignedTo$') {
                throw 'Insufficient privileges to complete the operation.'
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -IncludeAppRoleAssignments -Confirm:$false

        $result.AssignmentsFailed | Should -Be 1
        $result.AssignmentsSkipped | Should -Be 0
        $result.AssignmentsAdded | Should -Be 0

        Assert-MockCalled Write-NCMessage -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -match 'Failed to assign'
        }
    }
}

Describe 'Compare-NCEnterpriseApplicationSnapshot' {
    BeforeAll {
        function New-TestSnapshot {
            param([string]$DisplayName, [string[]]$RedirectUris, [object[]]$Owners = @())
            [pscustomobject]@{
                Application         = [pscustomobject]@{
                    DisplayName            = $DisplayName
                    SignInAudience         = 'AzureADMyOrg'
                    IdentifierUris         = @()
                    Notes                  = $null
                    Tags                   = @()
                    Owners                 = $Owners
                    Web                    = [pscustomobject]@{ redirectUris = $RedirectUris }
                    Spa                    = [pscustomobject]@{ redirectUris = @() }
                    PublicClient           = [pscustomobject]@{ redirectUris = @() }
                    RequiredResourceAccess = @()
                    AppRoles               = @()
                    Oauth2PermissionScopes = @()
                }
                ServicePrincipal    = [pscustomobject]@{ Tags = @(); Homepage = $null; LogoUrl = $null }
                AppRoleAssignments  = @()
                CredentialsMetadata = [pscustomobject]@{ PasswordCredentials = @(); KeyCredentials = @() }
            }
        }
    }

    It 'reports no differences for identical snapshots' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @('https://localhost/callback')
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @('https://localhost/callback')

        $rows = Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b

        $rows.Count | Should -Be 0
    }

    It 'reports a row for a changed redirect URI' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @('https://localhost/callback')
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @('https://prod.contoso.com/callback')

        $rows = Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b

        ($rows | Where-Object { $_.Property -eq 'Application.Web' }).Count | Should -Be 1
    }

    It 'ignores App Role Assignments unless requested' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $a.AppRoleAssignments = @([pscustomobject]@{ PrincipalId = 'p1'; AppRoleId = 'r1' })

        $rowsWithoutAssignments = Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b
        $rowsWithAssignments = Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b -IncludeAppRoleAssignments

        ($rowsWithoutAssignments | Where-Object { $_.Property -eq 'AppRoleAssignments' }).Count | Should -Be 0
        ($rowsWithAssignments | Where-Object { $_.Property -eq 'AppRoleAssignments' }).Count | Should -Be 1
    }

    It 'reports a row for a changed owner' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @() -Owners @([pscustomobject]@{ Id = 'owner-1'; DisplayName = 'Jane Doe'; UserPrincipalName = 'jane@contoso.com' })
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @() -Owners @([pscustomobject]@{ Id = 'owner-2'; DisplayName = 'John Smith'; UserPrincipalName = 'john@contoso.com' })

        $rows = Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b

        ($rows | Where-Object { $_.Property -eq 'Application.Owners' }).Count | Should -Be 1
    }

    It 'ignores the order of permissions, redirect URIs, app roles, scopes and tags' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @('https://a.contoso.com/cb', 'https://b.contoso.com/cb')
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @('https://b.contoso.com/cb', 'https://a.contoso.com/cb')
        $graphPermissions = [pscustomobject]@{ resourceAppId = 'res-graph'; resourceAccess = @([pscustomobject]@{ id = 'p1'; type = 'Scope' }, [pscustomobject]@{ id = 'p2'; type = 'Role' }) }
        $graphPermissionsReordered = [pscustomobject]@{ resourceAppId = 'res-graph'; resourceAccess = @([pscustomobject]@{ id = 'p2'; type = 'Role' }, [pscustomobject]@{ id = 'p1'; type = 'Scope' }) }
        $otherPermissions = [pscustomobject]@{ resourceAppId = 'res-other'; resourceAccess = @([pscustomobject]@{ id = 'p3'; type = 'Scope' }) }
        $a.Application.RequiredResourceAccess = @($graphPermissions, $otherPermissions)
        $b.Application.RequiredResourceAccess = @($otherPermissions, $graphPermissionsReordered)
        $a.Application.AppRoles = @([pscustomobject]@{ id = 'r1'; value = 'Read' }, [pscustomobject]@{ id = 'r2'; value = 'Write' })
        $b.Application.AppRoles = @([pscustomobject]@{ id = 'r2'; value = 'Write' }, [pscustomobject]@{ id = 'r1'; value = 'Read' })
        $a.Application.Oauth2PermissionScopes = @([pscustomobject]@{ id = 's1' }, [pscustomobject]@{ id = 's2' })
        $b.Application.Oauth2PermissionScopes = @([pscustomobject]@{ id = 's2' }, [pscustomobject]@{ id = 's1' })
        $a.Application.Tags = @('t1', 't2')
        $b.Application.Tags = @('t2', 't1')
        $a.ServicePrincipal.Tags = @('WindowsAzureActiveDirectoryIntegratedApp', 'HideApp')
        $b.ServicePrincipal.Tags = @('HideApp', 'WindowsAzureActiveDirectoryIntegratedApp')

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        $rows.Count | Should -Be 0
    }

    It 'ignores the order of allowed member types inside an app role' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $a.Application.AppRoles = @([pscustomobject]@{ id = 'r1'; value = 'Read'; allowedMemberTypes = @('User', 'Application') })
        $b.Application.AppRoles = @([pscustomobject]@{ id = 'r1'; value = 'Read'; allowedMemberTypes = @('Application', 'User') })

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        $rows.Count | Should -Be 0
    }
    It 'ignores the order of optional claims within each token type' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $email = [pscustomobject]@{ name = 'email'; source = $null; essential = $false; additionalProperties = @() }
        $upn = [pscustomobject]@{ name = 'upn'; source = $null; essential = $false; additionalProperties = @() }
        $groups = [pscustomobject]@{ name = 'groups'; source = $null; essential = $false; additionalProperties = @('sam_account_name', 'emit_as_roles') }
        $groupsReordered = [pscustomobject]@{ name = 'groups'; source = $null; essential = $false; additionalProperties = @('emit_as_roles', 'sam_account_name') }
        $a.Application | Add-Member -NotePropertyName OptionalClaims -NotePropertyValue ([pscustomobject]@{ idToken = @($email, $upn); accessToken = @($groups); saml2Token = @() })
        $b.Application | Add-Member -NotePropertyName OptionalClaims -NotePropertyValue ([pscustomobject]@{ idToken = @($upn, $email); accessToken = @($groupsReordered); saml2Token = @() })

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        $rows.Count | Should -Be 0
    }
    It 'still reports a real permission difference' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $a.Application.RequiredResourceAccess = @([pscustomobject]@{ resourceAppId = 'res-graph'; resourceAccess = @([pscustomobject]@{ id = 'p1'; type = 'Scope' }) })
        $b.Application.RequiredResourceAccess = @([pscustomobject]@{ resourceAppId = 'res-graph'; resourceAccess = @([pscustomobject]@{ id = 'p1'; type = 'Role' }) })

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        @($rows.Property) | Should -Contain 'Application.RequiredResourceAccess'
    }
    It 'reports changed token claims and assignment requirement' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $a.Application | Add-Member -NotePropertyName GroupMembershipClaims -NotePropertyValue 'SecurityGroup'
        $b.Application | Add-Member -NotePropertyName GroupMembershipClaims -NotePropertyValue 'None'
        $a.Application | Add-Member -NotePropertyName OptionalClaims -NotePropertyValue ([pscustomobject]@{ idToken = @([pscustomobject]@{ name = 'email' }) })
        $a.ServicePrincipal | Add-Member -NotePropertyName AppRoleAssignmentRequired -NotePropertyValue $true
        $b.ServicePrincipal | Add-Member -NotePropertyName AppRoleAssignmentRequired -NotePropertyValue $false

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        @($rows.Property) | Should -Contain 'Application.GroupMembershipClaims'
        @($rows.Property) | Should -Contain 'Application.OptionalClaims'
        @($rows.Property) | Should -Contain 'ServicePrincipal.AppRoleAssignmentRequired'
    }
    It 'reports changed Service Principal owners and enabled state' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $a.ServicePrincipal | Add-Member -NotePropertyName AccountEnabled -NotePropertyValue $true
        $b.ServicePrincipal | Add-Member -NotePropertyName AccountEnabled -NotePropertyValue $false
        $a.ServicePrincipal | Add-Member -NotePropertyName Owners -NotePropertyValue @([pscustomobject]@{ Id = 'o1' }, [pscustomobject]@{ Id = 'o2' })
        $b.ServicePrincipal | Add-Member -NotePropertyName Owners -NotePropertyValue @([pscustomobject]@{ Id = 'o2' }, [pscustomobject]@{ Id = 'o3' })

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        @($rows.Property) | Should -Contain 'ServicePrincipal.AccountEnabled'
        @($rows.Property) | Should -Contain 'ServicePrincipal.Owners'
    }
    It 'reports a changed fallback public client flag and SSO mode' {
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @()
        $a.Application | Add-Member -NotePropertyName IsFallbackPublicClient -NotePropertyValue $true
        $b.Application | Add-Member -NotePropertyName IsFallbackPublicClient -NotePropertyValue $false
        $a.ServicePrincipal | Add-Member -NotePropertyName PreferredSingleSignOnMode -NotePropertyValue 'saml'

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b)

        @($rows.Property) | Should -Contain 'Application.IsFallbackPublicClient'
        @($rows.Property) | Should -Contain 'ServicePrincipal.PreferredSingleSignOnMode'
    }
    It 'ignores the order of owners, App Role Assignments and credentials' {
        $jane = [pscustomobject]@{ Id = 'owner-1'; DisplayName = 'Jane Doe'; UserPrincipalName = 'jane@contoso.com' }
        $john = [pscustomobject]@{ Id = 'owner-2'; DisplayName = 'John Smith'; UserPrincipalName = 'john@contoso.com' }
        $a = New-TestSnapshot -DisplayName 'App' -RedirectUris @() -Owners @($jane, $john)
        $b = New-TestSnapshot -DisplayName 'App' -RedirectUris @() -Owners @($john, $jane)
        $assignment1 = [pscustomobject]@{ PrincipalId = 'p1'; PrincipalDisplayName = 'Group 1'; PrincipalType = 'Group'; AppRoleId = 'r1' }
        $assignment2 = [pscustomobject]@{ PrincipalId = 'p2'; PrincipalDisplayName = 'Group 2'; PrincipalType = 'Group'; AppRoleId = 'r1' }
        $a.AppRoleAssignments = @($assignment1, $assignment2)
        $b.AppRoleAssignments = @($assignment2, $assignment1)
        $secret1 = [pscustomobject]@{ DisplayName = 's1'; KeyId = 'k1'; EndDateTime = '2027-01-01' }
        $secret2 = [pscustomobject]@{ DisplayName = 's2'; KeyId = 'k2'; EndDateTime = '2027-06-01' }
        $a.CredentialsMetadata.PasswordCredentials = @($secret1, $secret2)
        $b.CredentialsMetadata.PasswordCredentials = @($secret2, $secret1)

        $rows = @(Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $a -DifferenceSnapshot $b -IncludeAppRoleAssignments)

        $rows.Count | Should -Be 0
    }
}

Describe 'Export-EnterpriseApplication' {
    BeforeAll {
        $snapshot = [pscustomobject]@{
            SchemaVersion = 1
            Application   = [pscustomobject]@{ DisplayName = 'Contoso Test App' }
        }
        $outputPath = Join-Path $TestDrive 'export.json'
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Get-NCEnterpriseApplicationSnapshot { $snapshot }
    }

    It 'writes the snapshot as JSON to the output path' {
        Export-EnterpriseApplication -ApplicationName 'Contoso Test App' -OutputPath $outputPath

        Test-Path -LiteralPath $outputPath | Should -Be $true
        $written = Get-Content -LiteralPath $outputPath -Raw | ConvertFrom-Json
        $written.Application.DisplayName | Should -Be 'Contoso Test App'
    }

    It 'refuses to overwrite an existing file without -Force' {
        Set-Content -LiteralPath $outputPath -Value '{}'

        Export-EnterpriseApplication -ApplicationName 'Contoso Test App' -OutputPath $outputPath

        Assert-MockCalled Get-NCEnterpriseApplicationSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }

    It 'overwrites an existing file when -Force is used' {
        Set-Content -LiteralPath $outputPath -Value '{}'

        Export-EnterpriseApplication -ApplicationName 'Contoso Test App' -OutputPath $outputPath -Force

        $written = Get-Content -LiteralPath $outputPath -Raw | ConvertFrom-Json
        $written.Application.DisplayName | Should -Be 'Contoso Test App'
    }

    It 'stops early when Microsoft Graph is not connected' {
        Mock Test-MgGraphConnection { $false }

        Export-EnterpriseApplication -ApplicationName 'Contoso Test App' -OutputPath $outputPath

        Assert-MockCalled Get-NCEnterpriseApplicationSnapshot -Times 0 -Scope It
    }
}

Describe 'Import-EnterpriseApplication' {
    BeforeAll {
        $inputPath = Join-Path $TestDrive 'import.json'
        $applyResult = [pscustomobject]@{ TargetDisplayName = 'Target App'; TargetApplicationId = 'app-1'; Created = $true; OwnersAdded = 0; OwnersSkipped = 0; AssignmentsAdded = 0; AssignmentsSkipped = 0 }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Set-NCEnterpriseApplicationFromSnapshot { $applyResult }
        $script:validSnapshot = @{
            SchemaVersion    = 1
            Application      = @{ DisplayName = 'Source App'; SignInAudience = 'AzureADMyOrg'; Notes = $null; Tags = @(); Web = @{ redirectUris = @() }; Spa = @{ redirectUris = @() }; PublicClient = @{ redirectUris = @() }; RequiredResourceAccess = @(); AppRoles = @(); Oauth2PermissionScopes = @(); Owners = @() }
            ServicePrincipal = @{ Tags = @(); Homepage = $null }
        }
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath
    }

    It 'reads the snapshot file and applies it to the target' {
        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 1 -Scope It -ParameterFilter {
            $TargetDisplayName -eq 'Target App' -and $Snapshot.Application.DisplayName -eq 'Source App'
        }
    }

    It 'emits the result object only with -PassThru' {
        $result = Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false
        $result | Should -BeNullOrEmpty

        $result = Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -PassThru -Confirm:$false
        $result.TargetApplicationId | Should -Be 'app-1'
    }

    It 'errors when the input file does not exist' {
        Import-EnterpriseApplication -InputPath (Join-Path $TestDrive 'missing.json') -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }

    It 'does not apply anything under -WhatIf' {
        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -WhatIf

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
    }

    It 'stops early when Microsoft Graph is not connected' {
        Mock Test-MgGraphConnection { $false }

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
    }

    It 'requests AppRoleAssignment.ReadWrite.All only when copying App Role Assignments' {
        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false
        Assert-MockCalled Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $Scopes -notcontains 'AppRoleAssignment.ReadWrite.All' }

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -IncludeAppRoleAssignments -Confirm:$false
        Assert-MockCalled Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter {
            $Scopes -contains 'AppRoleAssignment.ReadWrite.All' -and $Scopes -contains 'Application.ReadWrite.All'
        }
    }

    It 'refuses a snapshot with an unsupported schema version' {
        $script:validSnapshot.SchemaVersion = 2
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*schema version*' }
    }

    It 'refuses a snapshot that misses required application properties' {
        $script:validSnapshot.Application.Remove('RequiredResourceAccess')
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*RequiredResourceAccess*' }
    }

    It 'refuses a snapshot without Notes or Service Principal Tags/Homepage' {
        $script:validSnapshot.Application.Remove('Notes')
        $script:validSnapshot.ServicePrincipal.Remove('Tags')
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*Application.Notes*' -and $Message -like '*ServicePrincipal.Tags*' }
    }
    It 'refuses a snapshot whose redirect URI containers miss redirectUris' {
        $script:validSnapshot.Application.Web = @{ homePageUrl = 'https://contoso.com' }
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*Application.Web.redirectUris*' }
    }
    It 'refuses a snapshot whose Api misses some of its settings' {
        $script:validSnapshot.Application.Api = @{ AcceptMappedClaims = $null; KnownClientApplications = @() }
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter {
            $Level -eq 'ERROR' -and $Message -like '*Application.Api.PreAuthorizedApplications*' -and $Message -like '*Application.Api.RequestedAccessTokenVersion*'
        }
    }

    It 'accepts a snapshot with a complete Api' {
        $script:validSnapshot.Application.Api = @{ AcceptMappedClaims = $null; KnownClientApplications = @(); PreAuthorizedApplications = @(); RequestedAccessTokenVersion = 2 }
        $script:validSnapshot | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 1 -Scope It
    }
    It 'refuses a JSON file that is not an Enterprise Application snapshot' {
        '{"name":"something else"}' | Set-Content -LiteralPath $inputPath

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }
    It 'errors when the input file contains invalid JSON' {
        Set-Content -LiteralPath $inputPath -Value 'not valid json {{{'

        Import-EnterpriseApplication -InputPath $inputPath -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }
}

Describe 'Copy-EnterpriseApplication' {
    BeforeAll {
        $snapshot = [pscustomobject]@{ Application = [pscustomobject]@{ DisplayName = 'Source App' } }
        $applyResult = [pscustomobject]@{ TargetDisplayName = 'Target App'; TargetApplicationId = 'app-1'; Created = $true; OwnersAdded = 0; OwnersSkipped = 0; AssignmentsAdded = 0; AssignmentsSkipped = 0 }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Get-NCEnterpriseApplicationSnapshot { $snapshot }
        Mock Set-NCEnterpriseApplicationFromSnapshot { $applyResult }
    }

    It 'reads the source snapshot and applies it to the target' {
        Copy-EnterpriseApplication -SourceApplicationName 'Source App' -TargetDisplayName 'Target App' -Confirm:$false

        Assert-MockCalled Get-NCEnterpriseApplicationSnapshot -Times 1 -Scope It -ParameterFilter { $ApplicationName -eq 'Source App' }
        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 1 -Scope It -ParameterFilter { $TargetDisplayName -eq 'Target App' }
    }

    It 'refuses to clone an application onto itself' {
        Mock Get-NCEnterpriseApplicationSnapshot { [pscustomobject]@{ Application = [pscustomobject]@{ DisplayName = 'Same App' } } }

        Copy-EnterpriseApplication -SourceApplicationName 'Same App' -TargetDisplayName 'Same App' -Confirm:$false

        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 0 -Scope It
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }

    It 'requests AppRoleAssignment.ReadWrite.All only when copying App Role Assignments' {
        Copy-EnterpriseApplication -SourceApplicationName 'Source App' -TargetDisplayName 'Target App' -Confirm:$false
        Assert-MockCalled Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $Scopes -notcontains 'AppRoleAssignment.ReadWrite.All' }

        Copy-EnterpriseApplication -SourceApplicationName 'Source App' -TargetDisplayName 'Target App' -IncludeAppRoleAssignments -Confirm:$false
        Assert-MockCalled Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter {
            $Scopes -contains 'AppRoleAssignment.ReadWrite.All' -and $Scopes -contains 'Application.ReadWrite.All'
        }
    }

    It 'passes -IncludeAppRoleAssignments through to both helpers' {
        Copy-EnterpriseApplication -SourceApplicationName 'Source App' -TargetDisplayName 'Target App' -IncludeAppRoleAssignments -Confirm:$false

        Assert-MockCalled Get-NCEnterpriseApplicationSnapshot -Times 1 -Scope It -ParameterFilter { $IncludeAppRoleAssignments }
        Assert-MockCalled Set-NCEnterpriseApplicationFromSnapshot -Times 1 -Scope It -ParameterFilter { $IncludeAppRoleAssignments }
    }
}

Describe 'Compare-EnterpriseApplication' {
    BeforeAll {
        $referencePath = Join-Path $TestDrive 'reference.json'
        $differencePath = Join-Path $TestDrive 'difference.json'
        $diffRows = @([pscustomobject]@{ Property = 'Application.Web'; ReferenceValue = 'a'; DifferenceValue = 'b' })
    }

    BeforeEach {
        $script:NCVars = @{ CSV_Encoding = 'UTF-8'; CSV_DefaultLimiter = ',' }
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Compare-NCEnterpriseApplicationSnapshot { $diffRows }
        '{"Application":{"DisplayName":"Reference App"}}' | Set-Content -LiteralPath $referencePath
        '{"Application":{"DisplayName":"Difference App"}}' | Set-Content -LiteralPath $differencePath
    }

    It 'compares two files without needing a Graph connection' {
        $rows = Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath

        $rows[0].Property | Should -Be 'Application.Web'
        Assert-MockCalled Test-MgGraphConnection -Times 0 -Scope It
    }

    It 'compares a file against a live application' {
        Mock Get-NCEnterpriseApplicationSnapshot { [pscustomobject]@{ Application = [pscustomobject]@{ DisplayName = 'Live App' } } }

        $rows = Compare-EnterpriseApplication -ReferencePath $referencePath -DifferenceApplicationName 'Live App'

        $rows[0].Property | Should -Be 'Application.Web'
        Assert-MockCalled Test-MgGraphConnection -Times 1 -Scope It
        Assert-MockCalled Get-NCEnterpriseApplicationSnapshot -Times 1 -Scope It -ParameterFilter { $ApplicationName -eq 'Live App' }
    }

    It 'errors when more than one reference source is given' {
        $rows = Compare-EnterpriseApplication -ReferencePath $referencePath -ReferenceApplicationName 'X' -DifferencePath $differencePath

        $rows | Should -BeNullOrEmpty
        Assert-MockCalled Write-NCMessage -Times 1 -Scope It -ParameterFilter { $Level -eq 'ERROR' }
    }

    It 'writes a JSON report when -OutputReportPath is given' {
        $reportPath = Join-Path $TestDrive 'report.json'

        Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath -OutputReportPath $reportPath

        Test-Path -LiteralPath $reportPath | Should -Be $true
        $written = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
        $written[0].Property | Should -Be 'Application.Web'
    }

    It 'writes a JSON array (not a bare object or null) when there is exactly one difference' {
        Mock Compare-NCEnterpriseApplicationSnapshot { @([pscustomobject]@{ Property = 'Application.Web'; ReferenceValue = 'a'; DifferenceValue = 'b' }) }
        $reportPath = Join-Path $TestDrive 'single-diff-report.json'

        Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath -OutputReportPath $reportPath

        $rawContent = Get-Content -LiteralPath $reportPath -Raw
        $rawContent.TrimStart() | Should -Match '^\['
        $written = $rawContent | ConvertFrom-Json
        @($written).Count | Should -Be 1
    }

    It 'writes an empty JSON array (not null) when there are no differences' {
        Mock Compare-NCEnterpriseApplicationSnapshot { @() }
        $reportPath = Join-Path $TestDrive 'no-diff-report.json'

        Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath -OutputReportPath $reportPath

        $rawContent = Get-Content -LiteralPath $reportPath -Raw
        # ConvertTo-Json pretty-prints an empty array as "[\r\n\r\n]" (internal whitespace), so compare
        # with all whitespace stripped rather than a strict string match.
        ($rawContent -replace '\s', '') | Should -Be '[]'
        $parsed = $rawContent | ConvertFrom-Json
        @($parsed).Count | Should -Be 0
    }

    It 'writes JSON-projected, distinguishable values for non-scalar diff rows in the CSV report' {
        Mock Compare-NCEnterpriseApplicationSnapshot {
            @([pscustomobject]@{
                Property        = 'Application.Web'
                ReferenceValue  = [pscustomobject]@{ redirectUris = @('https://ref.contoso.com/callback') }
                DifferenceValue = [pscustomobject]@{ redirectUris = @('https://diff.contoso.com/callback') }
            })
        }
        $reportPath = Join-Path $TestDrive 'nonscalar-report.csv'

        Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath -OutputReportPath $reportPath

        $written = Import-Csv -LiteralPath $reportPath
        $written[0].ReferenceValue | Should -Not -Be $written[0].DifferenceValue
        $written[0].ReferenceValue | Should -Match 'ref\.contoso\.com'
        $written[0].DifferenceValue | Should -Match 'diff\.contoso\.com'
    }

    It 'renders a null diff value as an empty CSV cell instead of erroring' {
        Mock Compare-NCEnterpriseApplicationSnapshot {
            @([pscustomobject]@{
                Property        = 'Application.Notes'
                ReferenceValue  = $null
                DifferenceValue = 'Updated notes'
            })
        }
        $reportPath = Join-Path $TestDrive 'null-value-report.csv'

        Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath -OutputReportPath $reportPath

        $written = Import-Csv -LiteralPath $reportPath
        $written[0].ReferenceValue | Should -Be ''
        $written[0].DifferenceValue | Should -Be '"Updated notes"'
    }

    It 'writes the CSV using the configured CSV encoding and delimiter defaults' {
        $script:NCVars.CSV_DefaultLimiter = ';'
        $reportPath = Join-Path $TestDrive 'delimiter-report.csv'

        Compare-EnterpriseApplication -ReferencePath $referencePath -DifferencePath $differencePath -OutputReportPath $reportPath

        $rawLines = Get-Content -LiteralPath $reportPath
        $rawLines[0] | Should -Match ';'
        $rawLines[0] | Should -Not -Match ','
        $written = Import-Csv -LiteralPath $reportPath -Delimiter ';'
        $written[0].Property | Should -Be 'Application.Web'
    }
}
