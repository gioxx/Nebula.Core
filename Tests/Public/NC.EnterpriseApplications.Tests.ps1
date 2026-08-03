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
