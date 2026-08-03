#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Enterprise Application helpers =============================================================================================

function Get-NCEnterpriseApplicationSnapshot {
    <#
    .SYNOPSIS
        Reads an Enterprise Application (Application + Service Principal) into a normalized snapshot object.
    .PARAMETER ApplicationName
        Display name of the source Enterprise Application.
    .PARAMETER ApplicationId
        Object ID of the source Application.
    .PARAMETER IncludeAppRoleAssignments
        Also read App Role Assignments (users/groups assigned to the app).
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName')]
        [string]$ApplicationName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
        [string]$ApplicationId,

        [switch]$IncludeAppRoleAssignments
    )

    $selectProps = 'id,appId,displayName,signInAudience,identifierUris,notes,tags,web,spa,publicClient,requiredResourceAccess,appRoles,api,passwordCredentials,keyCredentials'

    if ($PSCmdlet.ParameterSetName -eq 'ById') {
        try {
            $app = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications/$ApplicationId`?`$select=$selectProps" -Method GET -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Enterprise Application with ID '$ApplicationId' not found: $($_.Exception.Message)" -Level ERROR
            return
        }
    }
    else {
        $escapedName = $ApplicationName.Replace("'", "''")
        try {
            $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$escapedName'&`$select=$selectProps" -Method GET -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to resolve Enterprise Application '$ApplicationName': $($_.Exception.Message)" -Level ERROR
            return
        }

        $foundApps = @($response.value)
        if ($foundApps.Count -eq 0) {
            Write-NCMessage "Enterprise Application '$ApplicationName' not found." -Level ERROR
            return
        }
        if ($foundApps.Count -gt 1) {
            Write-NCMessage "Multiple Enterprise Applications named '$ApplicationName' found. Use -ApplicationId instead." -Level ERROR
            return
        }

        $app = $foundApps[0]
    }

    try {
        $spResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id,appId,displayName,tags,homepage,logoUrl" -Method GET -ErrorAction Stop
    }
    catch {
        Write-NCMessage "Unable to read Service Principal for '$($app.displayName)': $($_.Exception.Message)" -Level ERROR
        return
    }

    $sp = @($spResponse.value) | Select-Object -First 1
    if (-not $sp) {
        Write-NCMessage "No Service Principal found for Application '$($app.displayName)' (appId: $($app.appId))." -Level ERROR
        return
    }

    try {
        $owners = @(Invoke-NCGraphAllPagesCore -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/owners?`$select=id,displayName,userPrincipalName")
    }
    catch {
        Write-NCMessage "Unable to read owners for '$($app.displayName)': $($_.Exception.Message)" -Level WARNING
        $owners = @()
    }

    $appRoleAssignments = @()
    if ($IncludeAppRoleAssignments.IsPresent) {
        try {
            $appRoleAssignments = @(Invoke-NCGraphAllPagesCore -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignedTo")
        }
        catch {
            Write-NCMessage "Unable to read App Role Assignments for '$($app.displayName)': $($_.Exception.Message)" -Level WARNING
            $appRoleAssignments = @()
        }
    }

    [pscustomobject][ordered]@{
        SchemaVersion       = 1
        ExportedAt          = (Get-Date).ToUniversalTime().ToString('o')
        Application         = [pscustomobject][ordered]@{
            DisplayName            = $app.displayName
            SignInAudience         = $app.signInAudience
            IdentifierUris         = @($app.identifierUris)
            Notes                  = $app.notes
            Tags                   = @($app.tags)
            Web                    = $app.web
            Spa                    = $app.spa
            PublicClient           = $app.publicClient
            RequiredResourceAccess = @($app.requiredResourceAccess)
            AppRoles               = @($app.appRoles)
            Oauth2PermissionScopes = @($app.api.oauth2PermissionScopes)
            Owners                 = @($owners | ForEach-Object {
                    [pscustomobject][ordered]@{
                        Id                = $_.id
                        DisplayName       = $_.displayName
                        UserPrincipalName = $_.userPrincipalName
                    }
                })
        }
        ServicePrincipal    = [pscustomobject][ordered]@{
            AppId       = $sp.appId
            DisplayName = $sp.displayName
            Tags        = @($sp.tags)
            Homepage    = $sp.homepage
            LogoUrl     = $sp.logoUrl
        }
        AppRoleAssignments  = @($appRoleAssignments | ForEach-Object {
                [pscustomobject][ordered]@{
                    PrincipalId          = $_.principalId
                    PrincipalDisplayName = $_.principalDisplayName
                    PrincipalType        = $_.principalType
                    AppRoleId            = $_.appRoleId
                }
            })
        CredentialsMetadata = [pscustomobject][ordered]@{
            PasswordCredentials = @($app.passwordCredentials | ForEach-Object {
                    [pscustomobject][ordered]@{
                        DisplayName = $_.displayName
                        KeyId       = $_.keyId
                        EndDateTime = $_.endDateTime
                    }
                })
            KeyCredentials      = @($app.keyCredentials | ForEach-Object {
                    [pscustomobject][ordered]@{
                        DisplayName = $_.displayName
                        KeyId       = $_.keyId
                        EndDateTime = $_.endDateTime
                    }
                })
        }
    }
}
