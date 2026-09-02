#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Connections helpers ========================================================================================================

function Test-EOLConnection {
    <#
    .SYNOPSIS
        Makes sure the Exchange Online session is ready to use.
    .DESCRIPTION
        Verifies that the ExchangeOnlineManagement module is present (installing it if asked),
        checks the current session and, if necessary, reconnects with the detected user. The
        combined Connect-Nebula flow can disable WAM after Graph initializes its dependencies.
    .PARAMETER UserPrincipalName
        Optional explicit UPN to use when connecting. Defaults to Find-UserConnected.
    .PARAMETER AutoInstall
        Install the module automatically instead of prompting.
    .PARAMETER ForceReconnect
        Skip the quick connectivity probe and always reconnect.
    .PARAMETER DisableWAM
        Disable Web Account Manager for the EXO sign-in flow.
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,
        [switch]$AutoInstall,
        [switch]$ForceReconnect,
        [switch]$DisableWAM
    )

    $moduleAvailable = @(Get-Module -Name ExchangeOnlineManagement -ListAvailable).Count -gt 0
    if (-not $moduleAvailable) {
        Write-NCMessage "Microsoft Exchange Online Management module is not available." -Level WARNING

        $installModule = $AutoInstall.IsPresent
        if (-not $installModule) {
            $confirmation = Read-Host "Install Microsoft Exchange Online Management module now? [Y] Yes [N] No"
            $installModule = $confirmation -match '^[yY]'
        }

        if (-not $installModule) {
            Add-EmptyLine
            Write-NCMessage "Microsoft Exchange Online Management module is required. Install it with Install-Module ExchangeOnlineManagement." -Level ERROR
            return $false
        }

        try {
            Write-Verbose "Installing Microsoft Exchange Online Management PowerShell module ..."
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop
        }
        catch {
            Add-EmptyLine
            Write-NCMessage "Can't install Exchange Online Management module. $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }

    if (-not $PSBoundParameters.ContainsKey('UserPrincipalName') -or [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $resolvedUpn = Find-UserConnected
    }
    else {
        $resolvedUpn = $UserPrincipalName
    }

    if (-not $ForceReconnect.IsPresent) {
        try {
            Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null

            $currentUpn = $null
            try {
                $currentUpn = (Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1).UserPrincipalName
            }
            catch {
                # Get-ConnectionInformation can throw (e.g. HttpResponseMessage/GetResponseHeader clash when
                # Microsoft.Graph is also loaded) even though the EXO session above is perfectly healthy.
                # Skip the UPN comparison rather than forcing an unnecessary reconnect.
            }

            if ([string]::IsNullOrWhiteSpace($resolvedUpn) -or [string]::IsNullOrWhiteSpace($currentUpn) -or $currentUpn -eq $resolvedUpn) {
                return $true
            }

            Write-NCMessage "Existing Exchange Online session is for $currentUpn, not $resolvedUpn. Reconnecting ..." -Level WARNING
        }
        catch {
            Write-NCMessage "Existing Exchange Online session not detected. Reconnecting ..." -Level WARNING
        }
    }

    $resolvedUpn = if ([string]::IsNullOrWhiteSpace($resolvedUpn)) { $null } else { $resolvedUpn }
    $message = if ($resolvedUpn) {
        "Connecting to Microsoft Exchange Online Management.`nUsing $resolvedUpn ..."
    }
    else {
        "Connecting to Microsoft Exchange Online Management ..."
    }
    Write-Verbose $message

    try {
        if ($resolvedUpn) {
            Connect-EOL -UserPrincipalName $resolvedUpn -DisableWAM:$DisableWAM.IsPresent -ErrorAction Stop
        }
        else {
            Connect-EOL -DisableWAM:$DisableWAM.IsPresent -ErrorAction Stop
        }

        # Run a lightweight cmdlet to be sure the session is alive.
        Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        Add-EmptyLine
        Write-NCMessage "Unable to establish Exchange Online session. $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Test-MgGraphConnection {
    <#
    .SYNOPSIS
        Ensures a Microsoft Graph session is available with the requested scopes.
    .DESCRIPTION
        Verifies that the Microsoft Graph module is present (installs it if requested),
        checks the current MgGraph context for the required scopes, and reconnects if needed.
        Optionally runs the Exchange Online prerequisite first to preserve legacy behaviour.
    .PARAMETER Scopes
        Delegated scopes to request during Connect-MgGraph. Defaults to User.Read.All.
    .PARAMETER TenantId
        Optional tenant to target explicitly.
    .PARAMETER UseDeviceCode
        Use device code authentication instead of opening a browser window.
    .PARAMETER AutoInstall
        Install Microsoft.Graph automatically if missing.
    .PARAMETER ForceReconnect
        Skip context validation and force a new Connect-MgGraph call.
    .PARAMETER EnsureExchangeOnline
        Run Test-EOLConnection before attempting Graph (defaults to true for backward compatibility).
    .PARAMETER LoginHint
        UPN to pass as -LoginHint to Connect-MgGraph, helping the WAM broker resolve the target
        account without prompting when it differs from the signed-in Windows account.
    #>
    [CmdletBinding()]
    param(
        [string[]]$Scopes = @('User.Read.All'),
        [string]$TenantId,
        [switch]$UseDeviceCode,
        [switch]$AutoInstall,
        [switch]$ForceReconnect,
        [bool]$EnsureExchangeOnline = $true,
        [string]$LoginHint
    )

    if ($EnsureExchangeOnline -and -not (Test-EOLConnection -AutoInstall:$AutoInstall.IsPresent)) {
        Write-NCMessage "Exchange Online prerequisite check failed. Skipping Microsoft Graph connection." -Level ERROR
        return $false
    }

    $requestedScopes = @($Scopes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not $requestedScopes) {
        $requestedScopes = @('User.Read.All')
    }

    $graphModuleAvailable = @(Get-Module -Name Microsoft.Graph -ListAvailable).Count -gt 0
    if (-not $graphModuleAvailable) {
        Write-NCMessage "Microsoft Graph PowerShell module is not available." -Level WARNING

        $installModule = $AutoInstall.IsPresent
        if (-not $installModule) {
            $confirmation = Read-Host "Install Microsoft Graph PowerShell module now? [Y] Yes [N] No"
            $installModule = $confirmation -match '^[yY]'
        }

        if (-not $installModule) {
            Add-EmptyLine
            Write-NCMessage "Microsoft Graph PowerShell module is required. Install it with Install-Module Microsoft.Graph." -Level ERROR
            return $false
        }

        try {
            Write-Verbose "Installing Microsoft Graph PowerShell module ..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop
        }
        catch {
            Add-EmptyLine
            Write-NCMessage "Can't install Microsoft Graph module. $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null
    }
    catch {
        Add-EmptyLine
        Write-NCMessage "Unable to import Microsoft.Graph.Authentication. $($_.Exception.Message)" -Level ERROR
        return $false
    }

    $missingScopeMessage = {
        param($currentScopes, $requiredScopes)
        if (-not $requiredScopes -or $requiredScopes.Count -eq 0) { return $null }
        if (-not $currentScopes -or $currentScopes.Count -eq 0) { return $requiredScopes }
        $requiredScopes | Where-Object { $currentScopes -notcontains $_ }
    }

    if (-not $ForceReconnect.IsPresent) {
        try {
            $ctx = Get-MgContext -ErrorAction Stop
            if ($ctx -and $ctx.Account) {
                if ((-not [string]::IsNullOrWhiteSpace($LoginHint)) -and $ctx.Account -ne $LoginHint) {
                    Write-NCMessage "Existing Microsoft Graph session is for $($ctx.Account), not $LoginHint. Reconnecting ..." -Level WARNING
                }
                else {
                    $missingScopes = &$missingScopeMessage $ctx.Scopes $requestedScopes
                    if (-not $missingScopes -or $missingScopes.Count -eq 0) {
                        return $true
                    }

                    Write-NCMessage "Existing Microsoft Graph session missing required scopes ($($missingScopes -join ', ')). Reconnecting ..." -Level WARNING
                }
            }
            else {
                Write-Verbose "Microsoft Graph context not found. Establishing connection ..."
            }
        }
        catch {
            Write-NCMessage "Unable to retrieve current Microsoft Graph context. Reconnecting ..." -Level WARNING
        }
    }

    $scopeLabel = $requestedScopes -join ', '
    Write-Verbose "Connecting to Microsoft Graph requesting scopes: $scopeLabel"

    $connectParams = @{
        Scopes    = $requestedScopes
        NoWelcome = $true
    }

    if ($TenantId) {
        $connectParams.TenantId = $TenantId
    }

    if ($UseDeviceCode.IsPresent) {
        $connectParams.UseDeviceCode = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($LoginHint)) {
        $connectParams.LoginHint = $LoginHint
    }

    try {
        Connect-MgGraph @connectParams | Out-Null

        $ctx = Get-MgContext -ErrorAction Stop
        if (-not $ctx -or -not $ctx.Account) {
            throw "Connect-MgGraph did not return an authenticated context."
        }

        $missingScopes = &$missingScopeMessage $ctx.Scopes $requestedScopes
        if ($missingScopes -and $missingScopes.Count -gt 0) {
            throw "Connected session is missing scopes: $($missingScopes -join ', ')"
        }

        return $true
    }
    catch {
        Add-EmptyLine
        Write-NCMessage "Unable to establish Microsoft Graph session. $($_.Exception.Message)" -Level ERROR
        return $false
    }
}
