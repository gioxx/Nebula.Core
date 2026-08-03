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
            AppRoleAssignments = @([pscustomobject]@{ PrincipalId = 'principal-1'; PrincipalDisplayName = 'Some Group'; PrincipalType = 'Group'; AppRoleId = 'role-1' })
        }
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
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/applications') {
                $script:capturedBody = $Body
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/servicePrincipals') {
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
            if ($Uri -match '^https://graph\.microsoft\.com/v1\.0/applications\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/applications') {
                $script:capturedBody = $Body
                return [pscustomobject]@{ id = 'new-app-id'; appId = 'new-client-id'; displayName = 'Target App' }
            }
            if ($Uri -match '/servicePrincipals\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'POST' -and $Uri -eq 'https://graph.microsoft.com/v1.0/servicePrincipals') {
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
        '{"Application":{"DisplayName":"Source App"}}' | Set-Content -LiteralPath $inputPath
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
