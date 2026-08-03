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

function Set-NCEnterpriseApplicationFromSnapshot {
    <#
    .SYNOPSIS
        Creates or updates an Enterprise Application (Application + Service Principal) from a snapshot.
    .PARAMETER Snapshot
        Snapshot object produced by Get-NCEnterpriseApplicationSnapshot.
    .PARAMETER TargetDisplayName
        Display name of the destination Enterprise Application. Created if missing, updated if it exists.
    .PARAMETER IncludeAppRoleAssignments
        Also apply App Role Assignments from the snapshot.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Snapshot,

        [Parameter(Mandatory = $true)]
        [string]$TargetDisplayName,

        [switch]$IncludeAppRoleAssignments
    )

    $escapedName = $TargetDisplayName.Replace("'", "''")
    try {
        $existingResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$escapedName'&`$select=id,appId,displayName" -Method GET -ErrorAction Stop
    }
    catch {
        Write-NCMessage "Unable to resolve target Enterprise Application '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
        return
    }

    $existingMatches = @($existingResponse.value)
    if ($existingMatches.Count -gt 1) {
        Write-NCMessage "Multiple Enterprise Applications named '$TargetDisplayName' found. Aborting to avoid ambiguity." -Level ERROR
        return
    }

    $targetApp = $existingMatches | Select-Object -First 1
    $created = $false

    $appBody = [ordered]@{
        displayName            = $TargetDisplayName
        signInAudience         = $Snapshot.Application.SignInAudience
        notes                  = $Snapshot.Application.Notes
        tags                   = @($Snapshot.Application.Tags)
        web                    = $Snapshot.Application.Web
        spa                    = $Snapshot.Application.Spa
        publicClient           = $Snapshot.Application.PublicClient
        requiredResourceAccess = @($Snapshot.Application.RequiredResourceAccess)
        appRoles               = @($Snapshot.Application.AppRoles)
        api                    = @{ oauth2PermissionScopes = @($Snapshot.Application.Oauth2PermissionScopes) }
    }

    if ($Snapshot.Application.IdentifierUris -and @($Snapshot.Application.IdentifierUris).Count -gt 0) {
        $appBody.identifierUris = @($Snapshot.Application.IdentifierUris)
    }

    if (-not $targetApp) {
        if (-not $PSCmdlet.ShouldProcess($TargetDisplayName, "Create Enterprise Application '$TargetDisplayName'")) {
            return
        }

        try {
            $targetApp = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/applications' -Method POST -Body ($appBody | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop
            $created = $true
            Write-NCMessage "Created Enterprise Application '$TargetDisplayName'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to create Enterprise Application '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
            return
        }
    }
    else {
        if (-not $PSCmdlet.ShouldProcess($TargetDisplayName, "Update Enterprise Application '$TargetDisplayName'")) {
            return
        }

        $patchBody = [ordered]@{}
        foreach ($key in $appBody.Keys) {
            if ($key -eq 'displayName') { continue }
            $patchBody[$key] = $appBody[$key]
        }

        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications/$($targetApp.id)" -Method PATCH -Body ($patchBody | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop | Out-Null
            Write-NCMessage "Updated Enterprise Application '$TargetDisplayName'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to update Enterprise Application '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
            return
        }
    }

    try {
        $spResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$($targetApp.appId)'&`$select=id,appId,displayName" -Method GET -ErrorAction Stop
    }
    catch {
        Write-NCMessage "Unable to check Service Principal for '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
        return
    }

    $targetSp = @($spResponse.value) | Select-Object -First 1
    if (-not $targetSp) {
        try {
            $targetSp = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Method POST -Body (@{ appId = $targetApp.appId } | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop
            Write-NCMessage "Created Service Principal for '$TargetDisplayName'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to create Service Principal for '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
            return
        }
    }

    $ownersAdded = 0
    $ownersSkipped = 0
    if ($Snapshot.Application.Owners -and @($Snapshot.Application.Owners).Count -gt 0) {
        try {
            $destinationOwners = @(Invoke-NCGraphAllPagesCore -Uri "https://graph.microsoft.com/v1.0/applications/$($targetApp.id)/owners?`$select=id")
        }
        catch {
            Write-NCMessage "Unable to read existing owners for '$TargetDisplayName': $($_.Exception.Message)" -Level WARNING
            $destinationOwners = @()
        }
        $destinationOwnerIds = @($destinationOwners | ForEach-Object { [string]$_.id })

        foreach ($owner in @($Snapshot.Application.Owners)) {
            if ($destinationOwnerIds -contains $owner.Id) {
                $ownersSkipped++
                continue
            }

            try {
                $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)" } | ConvertTo-Json -Depth 3
                Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications/$($targetApp.id)/owners/`$ref" -Method POST -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $ownersAdded++
                Write-NCMessage "Copied owner '$($owner.DisplayName)' to '$TargetDisplayName'." -Level SUCCESS
            }
            catch {
                if ($_.Exception.Message -match 'already exist' -or $_.Exception.Message -match 'exists') {
                    $ownersSkipped++
                }
                else {
                    Write-NCMessage "Failed to copy owner '$($owner.DisplayName)' to '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
                }
            }
        }
    }

    $assignmentsAdded = 0
    $assignmentsSkipped = 0
    if ($IncludeAppRoleAssignments.IsPresent -and $Snapshot.AppRoleAssignments -and @($Snapshot.AppRoleAssignments).Count -gt 0) {
        try {
            $destinationAssignments = @(Invoke-NCGraphAllPagesCore -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($targetSp.id)/appRoleAssignedTo")
        }
        catch {
            Write-NCMessage "Unable to read existing App Role Assignments for '$TargetDisplayName': $($_.Exception.Message)" -Level WARNING
            $destinationAssignments = @()
        }

        foreach ($assignment in @($Snapshot.AppRoleAssignments)) {
            $alreadyAssigned = $destinationAssignments | Where-Object {
                $_.principalId -eq $assignment.PrincipalId -and $_.appRoleId -eq $assignment.AppRoleId
            }
            if ($alreadyAssigned) {
                $assignmentsSkipped++
                continue
            }

            try {
                $body = @{
                    principalId = $assignment.PrincipalId
                    resourceId  = $targetSp.id
                    appRoleId   = $assignment.AppRoleId
                } | ConvertTo-Json -Depth 3
                Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($targetSp.id)/appRoleAssignedTo" -Method POST -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $assignmentsAdded++
                Write-NCMessage "Assigned '$($assignment.PrincipalDisplayName)' to '$TargetDisplayName'." -Level SUCCESS
            }
            catch {
                Write-NCMessage "Failed to assign '$($assignment.PrincipalDisplayName)' to '$TargetDisplayName': $($_.Exception.Message)" -Level WARNING
                $assignmentsSkipped++
            }
        }
    }

    [pscustomobject][ordered]@{
        TargetDisplayName   = $TargetDisplayName
        TargetApplicationId = $targetApp.id
        Created             = $created
        OwnersAdded         = $ownersAdded
        OwnersSkipped       = $ownersSkipped
        AssignmentsAdded    = $assignmentsAdded
        AssignmentsSkipped  = $assignmentsSkipped
    }
}
