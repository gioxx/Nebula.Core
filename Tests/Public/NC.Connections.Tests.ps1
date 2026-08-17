$global:NCConnectionOrder = @()

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
        [bool]$EnsureExchangeOnline
    )

    $global:NCConnectionOrder += 'MicrosoftGraph'
    return $true
}

. "$PSScriptRoot/../../Public/NC.Connections.ps1"

Describe 'Connect-Nebula' {
    BeforeEach {
        $global:NCConnectionOrder = @()
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
}
