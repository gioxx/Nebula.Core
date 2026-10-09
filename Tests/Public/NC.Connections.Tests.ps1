BeforeAll {
    $global:NCConnectionOrder = @()
    $global:NCVars = @{ CheckUpdatesOnConnect = $false }

    function Write-NCMessage {
        param(
            [string]$Message,
            [string]$Level
        )
    }

    function Test-NebulaModuleUpdates {}

    function Test-EOLConnection {
        param(
            [string]$UserPrincipalName,
            [switch]$AutoInstall,
            [switch]$ForceReconnect,
            [switch]$DisableWAM
        )

        $global:NCConnectionOrder += if ($DisableWAM) { 'ExchangeOnlineWithoutWam' } else { 'ExchangeOnline' }
        return $true
    }

    function Test-MgGraphConnection {
        param(
            [string[]]$Scopes,
            [string]$TenantId,
            [switch]$UseDeviceCode,
            [switch]$AutoInstall,
            [switch]$ForceReconnect,
            [bool]$EnsureExchangeOnline,
            [string]$LoginHint
        )

        $global:NCConnectionOrder += 'MicrosoftGraph'
        return $true
    }

    function Find-UserConnected {}
    function Get-MgContext {}
    function Get-ConnectionInformation {}
    function Get-NebulaConnections {}

    . "$PSScriptRoot/../../Public/NC.Connections.ps1"
}

Describe 'Connect-Nebula' {
    BeforeEach {
        $global:NCConnectionOrder = @()
        Mock Find-UserConnected { 'workstation@contoso.com' }
        Mock Get-MgContext { $null }
    }

    It 'connects to Microsoft Graph before Exchange Online' {
        $result = Connect-Nebula

        if ($result.ExchangeOnline -ne $true) { throw 'Exchange Online connection was not reported as successful.' }
        if ($result.MicrosoftGraph -ne $true) { throw 'Microsoft Graph connection was not reported as successful.' }
        if (($global:NCConnectionOrder -join ',') -ne 'MicrosoftGraph,ExchangeOnlineWithoutWam') {
            throw "Unexpected connection order: $($global:NCConnectionOrder -join ',')"
        }
    }

    It 'keeps Exchange Online-only behavior when Graph is skipped' {
        $result = Connect-Nebula -SkipGraph

        if ($result.ExchangeOnline -ne $true) { throw 'Exchange Online connection was not reported as successful.' }
        if ($result.MicrosoftGraph -ne $false) { throw 'Microsoft Graph should be skipped.' }
        if (($global:NCConnectionOrder -join ',') -ne 'ExchangeOnline') {
            throw "Unexpected connection order with Graph skipped: $($global:NCConnectionOrder -join ',')"
        }
    }

    It 'keeps the account of an active Graph session when no identity is given' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com' } }
        Mock Test-MgGraphConnection { $true }

        $null = Connect-Nebula

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $LoginHint -eq 'admin@contoso.com' }
        Should -Invoke Find-UserConnected -Times 0 -Exactly -Scope It
    }

    It 'falls back to the workstation identity when no Graph session is active' {
        Mock Test-MgGraphConnection { $true }

        $null = Connect-Nebula

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $LoginHint -eq 'workstation@contoso.com' }
    }

    It 'keeps the tenant of an active Graph session when no tenant is given' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com'; TenantId = 'tenant-guest' } }
        Mock Test-MgGraphConnection { $true }

        $null = Connect-Nebula

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $TenantId -eq 'tenant-guest' }
    }
    It 'prefers an explicit -GraphLoginHint over the active Graph account' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com' } }
        Mock Test-MgGraphConnection { $true }

        $null = Connect-Nebula -GraphLoginHint 'other@contoso.com'

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $LoginHint -eq 'other@contoso.com' }
    }
}

Describe 'Update-NebulaConnections' {
    BeforeEach {
        Mock Test-EOLConnection { $true }
        Mock Test-MgGraphConnection { $true }
        Mock Get-ConnectionInformation { $null }
        Mock Get-NebulaConnections {}
        Mock Find-UserConnected { 'workstation@contoso.com' }
    }

    It 'repairs Graph with the account of the active session' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @('User.Read.All') } }

        Update-NebulaConnections

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $LoginHint -eq 'admin@contoso.com' }
        Should -Invoke Find-UserConnected -Times 0 -Exactly -Scope It
    }

    It 'repairs Microsoft Graph before Exchange Online, like Connect-Nebula' {
        $script:repairOrder = [System.Collections.Generic.List[string]]::new()
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @('User.Read.All') } }
        Mock Test-MgGraphConnection { $script:repairOrder.Add('Graph'); $true }
        Mock Test-EOLConnection { $script:repairOrder.Add('EXO'); $true }

        Update-NebulaConnections

        ($script:repairOrder -join ',') | Should -Be 'Graph,EXO'
        Should -Invoke Test-EOLConnection -Times 1 -Exactly -Scope It -ParameterFilter { $DisableWAM }
    }

    It 'forces a Graph reconnect when the session is present but fails the health probe' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @('User.Read.All') } }
        Mock Get-NebulaConnections { [pscustomobject]@{ MicrosoftGraphConnected = $true; MicrosoftGraphHealthy = $false; ExchangeOnlineConnected = $true; ExchangeOnlineHealthy = $true } }

        Update-NebulaConnections

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $ForceReconnect }
    }

    It 'repairs Graph against the tenant of the current session' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @('User.Read.All'); TenantId = 'tenant-guest' } }

        Update-NebulaConnections

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $TenantId -eq 'tenant-guest' }
    }
    It 'does not force a Graph reconnect when the session is healthy' {
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @('User.Read.All') } }
        Mock Get-NebulaConnections { [pscustomobject]@{ MicrosoftGraphConnected = $true; MicrosoftGraphHealthy = $true; ExchangeOnlineConnected = $true; ExchangeOnlineHealthy = $true } }

        Update-NebulaConnections

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { -not $ForceReconnect }
    }
    It 'falls back to the workstation identity when no Graph session is active' {
        Mock Get-MgContext { $null }

        Update-NebulaConnections

        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It -ParameterFilter { $LoginHint -eq 'workstation@contoso.com' }
    }
}
