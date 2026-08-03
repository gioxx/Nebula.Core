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

. "$PSScriptRoot/../../Private/NC-Hlp.EnterpriseApplications.ps1"

Describe 'Get-NCEnterpriseApplicationSnapshot' {
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

    It 'includes App Role Assignments only when requested' {
        $snapshot = Get-NCEnterpriseApplicationSnapshot -ApplicationName 'Contoso Test App' -IncludeAppRoleAssignments

        $snapshot.AppRoleAssignments.Count | Should -Be 1
        $snapshot.AppRoleAssignments[0].PrincipalDisplayName | Should -Be 'Some Group'
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
        AppRoleAssignments = @([pscustomobject]@{ PrincipalId = 'principal-1'; PrincipalDisplayName = 'Some Group'; PrincipalType = 'Group'; AppRoleId = 'role-1' })
    }

    BeforeEach {
        Mock Write-NCMessage {}
    }

    It 'creates the destination application and service principal when none exists' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/applications') {
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/servicePrincipals') {
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
            $Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/applications'
        }
    }

    It 'updates an existing destination application instead of creating a new one' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-app-id'; appId = 'existing-client-id'; displayName = 'Target App' }) }
            }
            if ($Method -eq 'PATCH' -and $Uri -match '/applications/existing-app-id$') { return $null }
            if ($Uri -match '/servicePrincipals\?') {
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'existing-sp-id'; appId = 'existing-client-id' }) }
            }
            return $null
        }
        Mock Invoke-NCGraphAllPagesCore { return @() }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName 'Target App' -Confirm:$false

        $result.Created | Should -Be $false
        $result.TargetApplicationId | Should -Be 'existing-app-id'
        $result.AssignmentsFailed | Should -Be 0
        $result.Error | Should -BeNullOrEmpty
        Assert-MockCalled Invoke-MgGraphRequest -Scope It -ParameterFilter {
            $Method -eq 'PATCH' -and $Uri -eq 'https://graph.microsoft.com/v1.0/applications/existing-app-id'
        }
    }

    It 'applies App Role Assignments only when requested' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') {
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
            $Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/servicePrincipals/existing-sp-id/appRoleAssignedTo'
        }
    }

    It 'returns an error object without crashing when the target name is ambiguous' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') {
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

    It 'counts a non-duplicate App Role Assignment failure as failed, not skipped, and logs at ERROR' {
        Mock Invoke-MgGraphRequest {
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') {
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
