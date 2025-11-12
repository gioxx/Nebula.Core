#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Connections ==========================================================================================================================

function Connect-EOL {
    <#
    .SYNOPSIS
        Establishes a connection to Exchange Online using EXO V3 module.
    .DESCRIPTION
        Imports the ExchangeOnlineManagement module (if needed) and invokes Connect-ExchangeOnline
        with opinionated defaults. When no explicit UserPrincipalName is supplied, the helper
        Find-UserConnected is used to auto-detect the interactive user.
    .PARAMETER UserPrincipalName
        UPN/e-mail used for the authentication prompt. Defaults to the current user.
    .PARAMETER DelegatedOrganization
        Optional customer tenant to target when running in delegated admin scenarios.
    .PARAMETER PassThru
        Return the Connect-ExchangeOnline result (session info) to the caller.
    .EXAMPLE
        Connect-EOL -UserPrincipalName 'admin@tenant.onmicrosoft.com'
    .EXAMPLE
        'admin@tenant.onmicrosoft.com' | Connect-EOL -DelegatedOrganization 'customer.onmicrosoft.com' -PassThru
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('UPN', 'User')]
        [string]$UserPrincipalName,
        [string]$DelegatedOrganization,
        [switch]$PassThru
    )

    begin {
        $module = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Select-Object -First 1
        if (-not $module) {
            throw "ExchangeOnlineManagement module not found. Install it with Install-Module ExchangeOnlineManagement."
        }

        if (-not (Get-Module -Name ExchangeOnlineManagement)) {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
        }

        $baseConnectParams = @{
            ShowBanner            = $false
            SkipLoadingCmdletHelp = $true
        }
    }

    process {
        if (-not $PSBoundParameters.ContainsKey('UserPrincipalName') -or [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            $UserPrincipalName = Find-UserConnected
        }

        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            throw "Unable to determine the UserPrincipalName. Provide -UserPrincipalName explicitly."
        }

        $connectParams = $baseConnectParams.Clone()
        $connectParams.UserPrincipalName = $UserPrincipalName

        if ($PSBoundParameters.ContainsKey('DelegatedOrganization') -and $DelegatedOrganization) {
            $connectParams.DelegatedOrganization = $DelegatedOrganization
        }

        Write-NCMessage "Connecting to Exchange Online as $UserPrincipalName ..." -Level INFO
        $session = Connect-ExchangeOnline @connectParams

        if ($PassThru.IsPresent) {
            return $session
        }
    }
}

function Connect-Nebula {
    <#
    .SYNOPSIS
        Entry point to establish both Exchange Online and Microsoft Graph sessions.
    .DESCRIPTION
        Uses the private Test-* helpers to ensure Exchange Online and Microsoft Graph are connected,
        optionally forcing reconnects, installing modules, or skipping the Graph portion.
    .PARAMETER UserPrincipalName
        Optional explicit UPN for the Exchange Online connection.
    .PARAMETER GraphScopes
        Microsoft Graph delegated scopes to request (default User.Read.All).
    .PARAMETER GraphTenantId
        Optional tenant ID/domain for the Graph connection.
    .PARAMETER GraphDeviceCode
        Use device-code auth for Graph instead of launching a browser.
    .PARAMETER AutoInstall
        Install missing modules automatically without prompting.
    .PARAMETER ForceReconnect
        Force reconnect (skip health checks) for both services.
    .PARAMETER SkipGraph
        Only establish Exchange Online; skip Microsoft Graph.
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,
        [string[]]$GraphScopes = @('User.Read.All'),
        [string]$GraphTenantId,
        [switch]$GraphDeviceCode,
        [switch]$AutoInstall,
        [switch]$ForceReconnect,
        [switch]$SkipGraph
    )

    Write-NCMessage "Welcome to Nebula." -Level INFO

    $exoConnected = Test-EOLConnection -UserPrincipalName $UserPrincipalName `
        -AutoInstall:$AutoInstall.IsPresent `
        -ForceReconnect:$ForceReconnect.IsPresent

    if (-not $exoConnected) {
        throw "Failed to establish Exchange Online session."
    }

    if ($SkipGraph) {
        return [pscustomobject]@{
            ExchangeOnline = $true
            MicrosoftGraph = $false
        }
    }

    $graphConnected = Test-MgGraphConnection `
        -Scopes $GraphScopes `
        -TenantId $GraphTenantId `
        -UseDeviceCode:$GraphDeviceCode.IsPresent `
        -AutoInstall:$AutoInstall.IsPresent `
        -ForceReconnect:$ForceReconnect.IsPresent `
        -EnsureExchangeOnline:$false

    if (-not $graphConnected) {
        throw "Failed to establish Microsoft Graph session."
    }

    return [pscustomobject]@{
        ExchangeOnline = $true
        MicrosoftGraph = $true
    }
}

function Disconnect-Nebula {
    <#
    .SYNOPSIS
        Disconnects from Exchange Online and Microsoft Graph sessions.
    .DESCRIPTION
        Calls Disconnect-ExchangeOnline (if available) and Disconnect-MgGraph,
        suppressing common errors and allowing targeted disconnects.
    .PARAMETER ExchangeOnly
        Disconnect only Exchange Online.
    .PARAMETER GraphOnly
        Disconnect only Microsoft Graph.
    #>
    [CmdletBinding()]
    param(
        [switch]$ExchangeOnly,
        [switch]$GraphOnly
    )

    $disconnectExo = if ($GraphOnly) { $false } else { $true }
    $disconnectGraph = if ($ExchangeOnly) { $false } else { $true }

    if ($disconnectExo) {
        try {
            if (Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                Write-NCMessage "Exchange Online session disconnected." -Level INFO
            }
        }
        catch {
            Write-NCMessage "Failed to disconnect Exchange Online: $($_.Exception.Message)" -Level WARNING
        }
    }

    if ($disconnectGraph) {
        try {
            if (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue) {
                Disconnect-MgGraph -ErrorAction Stop | Out-Null
                Write-NCMessage "Microsoft Graph session disconnected." -Level INFO
            }
        }
        catch {
            Write-NCMessage "Failed to disconnect Microsoft Graph: $($_.Exception.Message)" -Level WARNING
        }
    }
}

Set-Alias -Name Leave-Nebula -Value Disconnect-Nebula
