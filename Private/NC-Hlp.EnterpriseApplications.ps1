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

    $selectProps = 'id,appId,displayName,signInAudience,identifierUris,notes,tags,web,spa,publicClient,requiredResourceAccess,appRoles,api,groupMembershipClaims,optionalClaims,passwordCredentials,keyCredentials'

    if ($PSCmdlet.ParameterSetName -eq 'ById') {
        try {
            $app = Invoke-MgGraphRequest -Uri "v1.0/applications/$ApplicationId`?`$select=$selectProps" -Method GET -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Enterprise Application with ID '$ApplicationId' not found: $($_.Exception.Message)" -Level ERROR
            return
        }
    }
    else {
        $escapedName = $ApplicationName.Replace("'", "''")
        try {
            $response = Invoke-MgGraphRequest -Uri "v1.0/applications?`$filter=displayName eq '$escapedName'&`$select=$selectProps" -Method GET -ErrorAction Stop
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
        $spResponse = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id,appId,displayName,tags,homepage,logoUrl,appRoleAssignmentRequired" -Method GET -ErrorAction Stop
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
        $owners = @(Invoke-NCGraphAllPagesCore -Uri "v1.0/applications/$($app.id)/owners?`$select=id,displayName,userPrincipalName")
    }
    catch {
        # An incomplete owner list would make the snapshot (and any copy from it) silently drop owners
        Write-NCMessage "Unable to read owners for '$($app.displayName)', snapshot not created: $($_.Exception.Message)" -Level ERROR
        return
    }

    $appRoleAssignments = @()
    if ($IncludeAppRoleAssignments.IsPresent) {
        try {
            $appRoleAssignments = @(Invoke-NCGraphAllPagesCore -Uri "v1.0/servicePrincipals/$($sp.id)/appRoleAssignedTo")
        }
        catch {
            Write-NCMessage "Unable to read App Role Assignments for '$($app.displayName)', snapshot not created: $($_.Exception.Message)" -Level ERROR
            return
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
            # Writable api settings besides the scopes, so a clone keeps pre-authorizations and token version
            Api                    = [pscustomobject][ordered]@{
                AcceptMappedClaims          = $app.api.acceptMappedClaims
                KnownClientApplications     = @($app.api.knownClientApplications)
                PreAuthorizedApplications   = @($app.api.preAuthorizedApplications)
                RequestedAccessTokenVersion = $app.api.requestedAccessTokenVersion
            }
            # Token claim settings: without them a clone issues different ID/access/SAML tokens
            GroupMembershipClaims  = $app.groupMembershipClaims
            OptionalClaims         = $app.optionalClaims
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
            # When true, only users and groups with an App Role Assignment can sign in
            AppRoleAssignmentRequired = [bool]$sp.appRoleAssignmentRequired
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
        $existingResponse = Invoke-MgGraphRequest -Uri "v1.0/applications?`$filter=displayName eq '$escapedName'&`$select=id,appId,displayName,appRoles,api" -Method GET -ErrorAction Stop
    }
    catch {
        Write-NCMessage "Unable to resolve target Enterprise Application '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
        [pscustomobject][ordered]@{
            TargetDisplayName   = $TargetDisplayName
            TargetApplicationId = $null
            Created             = $false
            OwnersAdded         = 0
            OwnersSkipped       = 0
            AssignmentsAdded    = 0
            AssignmentsSkipped  = 0
            AssignmentsFailed   = 0
            Error               = "Unable to resolve target Enterprise Application '$TargetDisplayName': $($_.Exception.Message)"
        }
        return
    }

    $existingMatches = @($existingResponse.value)
    if ($existingMatches.Count -gt 1) {
        Write-NCMessage "Multiple Enterprise Applications named '$TargetDisplayName' found. Aborting to avoid ambiguity." -Level ERROR
        [pscustomobject][ordered]@{
            TargetDisplayName   = $TargetDisplayName
            TargetApplicationId = $null
            Created             = $false
            OwnersAdded         = 0
            OwnersSkipped       = 0
            AssignmentsAdded    = 0
            AssignmentsSkipped  = 0
            AssignmentsFailed   = 0
            Error               = "Multiple Enterprise Applications named '$TargetDisplayName' found. Aborting to avoid ambiguity."
        }
        return
    }

    $targetApp = $existingMatches | Select-Object -First 1
    $created = $false

    # Graph returns read-only companion fields (e.g. web.redirectUriSettings) alongside the writable
    # ones when reading an application; sending both back on create/update is rejected outright
    # ("Can't modify redirectUris and redirectUriSettings within the same request."). Rebuild each
    # sub-object from only the fields Graph accepts on write, instead of forwarding the raw payload.
    $webSource = $Snapshot.Application.Web
    $cleanWeb = [ordered]@{
        redirectUris = @($webSource.redirectUris)
    }
    if ($webSource.homePageUrl) { $cleanWeb.homePageUrl = $webSource.homePageUrl }
    if ($webSource.logoutUrl) { $cleanWeb.logoutUrl = $webSource.logoutUrl }
    if ($webSource.implicitGrantSettings) { $cleanWeb.implicitGrantSettings = $webSource.implicitGrantSettings }

    $cleanApi = [ordered]@{ oauth2PermissionScopes = @($Snapshot.Application.Oauth2PermissionScopes) }
    # Snapshots saved before the Api property existed only carry the scopes
    $apiSource = $Snapshot.Application.Api
    if ($apiSource) {
        # Nulls are sent too: both settings are nullable, and omitting them would keep the destination's values
        $cleanApi.acceptMappedClaims = $apiSource.AcceptMappedClaims
        $cleanApi.knownClientApplications = @($apiSource.KnownClientApplications | Where-Object { $_ })
        $cleanApi.preAuthorizedApplications = @($apiSource.PreAuthorizedApplications | Where-Object { $_ })
        $cleanApi.requestedAccessTokenVersion = $apiSource.RequestedAccessTokenVersion
    }

    $appBody = [ordered]@{
        displayName            = $TargetDisplayName
        signInAudience         = $Snapshot.Application.SignInAudience
        notes                  = $Snapshot.Application.Notes
        tags                   = @($Snapshot.Application.Tags)
        web                    = $cleanWeb
        spa                    = @{ redirectUris = @($Snapshot.Application.Spa.redirectUris) }
        publicClient           = @{ redirectUris = @($Snapshot.Application.PublicClient.redirectUris) }
        requiredResourceAccess = @($Snapshot.Application.RequiredResourceAccess)
        appRoles               = @($Snapshot.Application.AppRoles)
        api                    = $cleanApi
    }
    # Snapshots saved before these properties existed leave the destination's values untouched
    $snapshotApplicationProperties = @($Snapshot.Application.PSObject.Properties.Name)
    if ($snapshotApplicationProperties -contains 'GroupMembershipClaims') { $appBody.groupMembershipClaims = $Snapshot.Application.GroupMembershipClaims }
    if ($snapshotApplicationProperties -contains 'OptionalClaims') { $appBody.optionalClaims = $Snapshot.Application.OptionalClaims }

    if ($Snapshot.Application.IdentifierUris -and @($Snapshot.Application.IdentifierUris).Count -gt 0) {
        Write-NCMessage "Source Enterprise Application '$($Snapshot.Application.DisplayName)' has identifierUris ($($Snapshot.Application.IdentifierUris -join ', ')). Identifier URIs are unique per tenant and are never copied to the destination app; set them manually if the destination needs to expose an API." -Level WARNING
    }

    if (-not $targetApp) {
        if (-not $PSCmdlet.ShouldProcess($TargetDisplayName, "Create Enterprise Application '$TargetDisplayName'")) {
            return
        }

        try {
            $targetApp = Invoke-MgGraphRequest -Uri 'v1.0/applications' -Method POST -Body ($appBody | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop
            $created = $true
            Write-NCMessage "Created Enterprise Application '$TargetDisplayName'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to create Enterprise Application '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
            [pscustomobject][ordered]@{
                TargetDisplayName   = $TargetDisplayName
                TargetApplicationId = $null
                Created             = $false
                OwnersAdded         = 0
                OwnersSkipped       = 0
                AssignmentsAdded    = 0
                AssignmentsSkipped  = 0
                AssignmentsFailed   = 0
                Error               = "Failed to create Enterprise Application '$TargetDisplayName': $($_.Exception.Message)"
            }
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

        # Graph only removes an app role or permission scope that is already disabled: disable the ones the
        # snapshot drops in a first update, then send the snapshot collections
        $toTable = {
            param($Item)
            $table = [ordered]@{}
            if ($Item -is [System.Collections.IDictionary]) { foreach ($k in $Item.Keys) { $table[$k] = $Item[$k] } }
            elseif ($null -ne $Item) { foreach ($property in $Item.PSObject.Properties) { $table[$property.Name] = $property.Value } }
            $table
        }
        $keptRoleIds = @($Snapshot.Application.AppRoles | ForEach-Object { [string]$_.id })
        $keptScopeIds = @($Snapshot.Application.Oauth2PermissionScopes | ForEach-Object { [string]$_.id })
        $obsoleteRoleIds = @($targetApp.appRoles | Where-Object { $_.isEnabled -and $keptRoleIds -notcontains [string]$_.id } | ForEach-Object { [string]$_.id })
        $obsoleteScopeIds = @($targetApp.api.oauth2PermissionScopes | Where-Object { $_.isEnabled -and $keptScopeIds -notcontains [string]$_.id } | ForEach-Object { [string]$_.id })

        if ($obsoleteRoleIds.Count -gt 0 -or $obsoleteScopeIds.Count -gt 0) {
            $stagingBody = [ordered]@{}
            if ($obsoleteRoleIds.Count -gt 0) {
                $stagingBody.appRoles = @($targetApp.appRoles | ForEach-Object {
                        $role = & $toTable $_
                        if ($obsoleteRoleIds -contains [string]$role.id) { $role.isEnabled = $false }
                        $role
                    })
            }
            if ($obsoleteScopeIds.Count -gt 0) {
                $stagingApi = & $toTable $targetApp.api
                $stagingApi.oauth2PermissionScopes = @($targetApp.api.oauth2PermissionScopes | ForEach-Object {
                        $scope = & $toTable $_
                        if ($obsoleteScopeIds -contains [string]$scope.id) { $scope.isEnabled = $false }
                        $scope
                    })
                # A pre-authorization can't point at a scope that is being disabled
                $stagingApi.preAuthorizedApplications = @($targetApp.api.preAuthorizedApplications | ForEach-Object {
                        $preAuthorized = & $toTable $_
                        $preAuthorized.delegatedPermissionIds = @($preAuthorized.delegatedPermissionIds | Where-Object { $obsoleteScopeIds -notcontains [string]$_ })
                        if ($preAuthorized.delegatedPermissionIds.Count -gt 0) { $preAuthorized }
                    })
                $stagingBody.api = $stagingApi
            }

            try {
                Invoke-MgGraphRequest -Uri "v1.0/applications/$($targetApp.id)" -Method PATCH -Body ($stagingBody | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop | Out-Null
                Write-Verbose "Disabled $($obsoleteRoleIds.Count) app role(s) and $($obsoleteScopeIds.Count) permission scope(s) before removing them from '$TargetDisplayName'."
            }
            catch {
                Write-NCMessage "Failed to disable the app roles or permission scopes removed from '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
                [pscustomobject][ordered]@{
                    TargetDisplayName   = $TargetDisplayName
                    TargetApplicationId = $null
                    Created             = $false
                    OwnersAdded         = 0
                    OwnersSkipped       = 0
                    AssignmentsAdded    = 0
                    AssignmentsSkipped  = 0
                    AssignmentsFailed   = 0
                    Error               = "Failed to disable the app roles or permission scopes removed from '$TargetDisplayName': $($_.Exception.Message)"
                }
                return
            }
        }

        try {
            Invoke-MgGraphRequest -Uri "v1.0/applications/$($targetApp.id)" -Method PATCH -Body ($patchBody | ConvertTo-Json -Depth 10) -ContentType 'application/json' -ErrorAction Stop | Out-Null
            Write-NCMessage "Updated Enterprise Application '$TargetDisplayName'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to update Enterprise Application '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
            [pscustomobject][ordered]@{
                TargetDisplayName   = $TargetDisplayName
                TargetApplicationId = $null
                Created             = $false
                OwnersAdded         = 0
                OwnersSkipped       = 0
                AssignmentsAdded    = 0
                AssignmentsSkipped  = 0
                AssignmentsFailed   = 0
                Error               = "Failed to update Enterprise Application '$TargetDisplayName': $($_.Exception.Message)"
            }
            return
        }
    }

    try {
        $spResponse = Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals?`$filter=appId eq '$($targetApp.appId)'&`$select=id,appId,displayName" -Method GET -ErrorAction Stop
    }
    catch {
        Write-NCMessage "Unable to check Service Principal for '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
        [pscustomobject][ordered]@{
            TargetDisplayName   = $TargetDisplayName
            TargetApplicationId = $null
            Created             = $false
            OwnersAdded         = 0
            OwnersSkipped       = 0
            AssignmentsAdded    = 0
            AssignmentsSkipped  = 0
            AssignmentsFailed   = 0
            Error               = "Unable to check Service Principal for '$TargetDisplayName': $($_.Exception.Message)"
        }
        return
    }

    # Tags (e.g. WindowsAzureActiveDirectoryIntegratedApp) are what makes the Entra Portal list a
    # Service Principal under "Enterprise applications" at all; a bare `appId`-only create leaves
    # the app visible in "App registrations" but absent from "Enterprise applications". homepageUrl
    # is applied too; logoUrl is read-only in Graph and is never written.
    # homepage is always written so an update also clears a value the snapshot doesn't have
    $spWriteBody = [ordered]@{
        tags     = @($Snapshot.ServicePrincipal.Tags)
        homepage = if ($Snapshot.ServicePrincipal.Homepage) { $Snapshot.ServicePrincipal.Homepage } else { $null }
    }
    if (@($Snapshot.ServicePrincipal.PSObject.Properties.Name) -contains 'AppRoleAssignmentRequired') {
        $spWriteBody.appRoleAssignmentRequired = [bool]$Snapshot.ServicePrincipal.AppRoleAssignmentRequired
    }

    $targetSp = @($spResponse.value) | Select-Object -First 1
    if (-not $targetSp) {
        $spCreateBody = [ordered]@{ appId = $targetApp.appId }
        foreach ($key in $spWriteBody.Keys) {
            if ($key -eq 'homepage' -and $null -eq $spWriteBody[$key]) { continue }
            $spCreateBody[$key] = $spWriteBody[$key]
        }

        try {
            $targetSp = Invoke-MgGraphRequest -Uri 'v1.0/servicePrincipals' -Method POST -Body ($spCreateBody | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop
            Write-NCMessage "Created Service Principal for '$TargetDisplayName'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to create Service Principal for '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
            [pscustomobject][ordered]@{
                TargetDisplayName   = $TargetDisplayName
                TargetApplicationId = $null
                Created             = $false
                OwnersAdded         = 0
                OwnersSkipped       = 0
                AssignmentsAdded    = 0
                AssignmentsSkipped  = 0
                AssignmentsFailed   = 0
                Error               = "Failed to create Service Principal for '$TargetDisplayName': $($_.Exception.Message)"
            }
            return
        }
    }
    else {
        try {
            Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals/$($targetSp.id)" -Method PATCH -Body ($spWriteBody | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop | Out-Null
        }
        catch {
            Write-NCMessage "Unable to update Service Principal tags/homepage for '$TargetDisplayName': $($_.Exception.Message)" -Level WARNING
        }
    }

    $ownersAdded = 0
    $ownersSkipped = 0
    if ($Snapshot.Application.Owners -and @($Snapshot.Application.Owners).Count -gt 0) {
        try {
            $destinationOwners = @(Invoke-NCGraphAllPagesCore -Uri "v1.0/applications/$($targetApp.id)/owners?`$select=id")
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
                $body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $owner.Id) } | ConvertTo-Json -Depth 3
                Invoke-MgGraphRequest -Uri "v1.0/applications/$($targetApp.id)/owners/`$ref" -Method POST -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
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
    $assignmentsFailed = 0
    if ($IncludeAppRoleAssignments.IsPresent -and $Snapshot.AppRoleAssignments -and @($Snapshot.AppRoleAssignments).Count -gt 0) {
        try {
            $destinationAssignments = @(Invoke-NCGraphAllPagesCore -Uri "v1.0/servicePrincipals/$($targetSp.id)/appRoleAssignedTo")
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
                Invoke-MgGraphRequest -Uri "v1.0/servicePrincipals/$($targetSp.id)/appRoleAssignedTo" -Method POST -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $assignmentsAdded++
                Write-NCMessage "Assigned '$($assignment.PrincipalDisplayName)' to '$TargetDisplayName'." -Level SUCCESS
            }
            catch {
                if ($_.Exception.Message -match 'already exist' -or $_.Exception.Message -match 'exists') {
                    $assignmentsSkipped++
                }
                else {
                    Write-NCMessage "Failed to assign '$($assignment.PrincipalDisplayName)' to '$TargetDisplayName': $($_.Exception.Message)" -Level ERROR
                    $assignmentsFailed++
                }
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
        AssignmentsFailed   = $assignmentsFailed
        Error               = $null
    }
}

function Compare-NCEnterpriseApplicationSnapshot {
    <#
    .SYNOPSIS
        Diffs two Enterprise Application snapshots.
    .PARAMETER ReferenceSnapshot
        Baseline snapshot (the "A" side).
    .PARAMETER DifferenceSnapshot
        Snapshot to compare against the baseline (the "B" side).
    .PARAMETER IncludeAppRoleAssignments
        Also compare App Role Assignments.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ReferenceSnapshot,

        [Parameter(Mandatory = $true)]
        [object]$DifferenceSnapshot,

        [switch]$IncludeAppRoleAssignments
    )

    $rows = [System.Collections.Generic.List[object]]::new()

    $addIfDifferent = {
        param($name, $refValue, $diffValue)
        $refJson = ($refValue | ConvertTo-Json -Depth 10 -Compress)
        $diffJson = ($diffValue | ConvertTo-Json -Depth 10 -Compress)
        if ($refJson -ne $diffJson) {
            $rows.Add([pscustomobject][ordered]@{
                    Property        = $name
                    ReferenceValue  = $refValue
                    DifferenceValue = $diffValue
                })
        }
    }

    # Graph returns these collections in no stable order: sort them so the same content compares equal
    $sortedBy = {
        param($Items, [string[]]$Keys)
        @(@($Items | Where-Object { $null -ne $_ }) | Sort-Object -Property $Keys)
    }

    & $addIfDifferent 'Application.DisplayName' $ReferenceSnapshot.Application.DisplayName $DifferenceSnapshot.Application.DisplayName
    & $addIfDifferent 'Application.SignInAudience' $ReferenceSnapshot.Application.SignInAudience $DifferenceSnapshot.Application.SignInAudience
    & $addIfDifferent 'Application.IdentifierUris' $ReferenceSnapshot.Application.IdentifierUris $DifferenceSnapshot.Application.IdentifierUris
    & $addIfDifferent 'Application.Notes' $ReferenceSnapshot.Application.Notes $DifferenceSnapshot.Application.Notes
    & $addIfDifferent 'Application.Tags' $ReferenceSnapshot.Application.Tags $DifferenceSnapshot.Application.Tags
    & $addIfDifferent 'Application.Owners' (& $sortedBy $ReferenceSnapshot.Application.Owners 'Id') (& $sortedBy $DifferenceSnapshot.Application.Owners 'Id')
    & $addIfDifferent 'Application.Web' $ReferenceSnapshot.Application.Web $DifferenceSnapshot.Application.Web
    & $addIfDifferent 'Application.Spa' $ReferenceSnapshot.Application.Spa $DifferenceSnapshot.Application.Spa
    & $addIfDifferent 'Application.PublicClient' $ReferenceSnapshot.Application.PublicClient $DifferenceSnapshot.Application.PublicClient
    & $addIfDifferent 'Application.RequiredResourceAccess' $ReferenceSnapshot.Application.RequiredResourceAccess $DifferenceSnapshot.Application.RequiredResourceAccess
    & $addIfDifferent 'Application.AppRoles' $ReferenceSnapshot.Application.AppRoles $DifferenceSnapshot.Application.AppRoles
    & $addIfDifferent 'Application.Oauth2PermissionScopes' $ReferenceSnapshot.Application.Oauth2PermissionScopes $DifferenceSnapshot.Application.Oauth2PermissionScopes
    & $addIfDifferent 'Application.Api' $ReferenceSnapshot.Application.Api $DifferenceSnapshot.Application.Api
    & $addIfDifferent 'Application.GroupMembershipClaims' $ReferenceSnapshot.Application.GroupMembershipClaims $DifferenceSnapshot.Application.GroupMembershipClaims
    & $addIfDifferent 'Application.OptionalClaims' $ReferenceSnapshot.Application.OptionalClaims $DifferenceSnapshot.Application.OptionalClaims
    & $addIfDifferent 'ServicePrincipal.Tags' $ReferenceSnapshot.ServicePrincipal.Tags $DifferenceSnapshot.ServicePrincipal.Tags
    & $addIfDifferent 'ServicePrincipal.Homepage' $ReferenceSnapshot.ServicePrincipal.Homepage $DifferenceSnapshot.ServicePrincipal.Homepage
    & $addIfDifferent 'ServicePrincipal.LogoUrl' $ReferenceSnapshot.ServicePrincipal.LogoUrl $DifferenceSnapshot.ServicePrincipal.LogoUrl
    & $addIfDifferent 'ServicePrincipal.AppRoleAssignmentRequired' $ReferenceSnapshot.ServicePrincipal.AppRoleAssignmentRequired $DifferenceSnapshot.ServicePrincipal.AppRoleAssignmentRequired

    if ($IncludeAppRoleAssignments.IsPresent) {
        & $addIfDifferent 'AppRoleAssignments' (& $sortedBy $ReferenceSnapshot.AppRoleAssignments 'PrincipalId', 'AppRoleId') (& $sortedBy $DifferenceSnapshot.AppRoleAssignments 'PrincipalId', 'AppRoleId')
    }

    & $addIfDifferent 'CredentialsMetadata.PasswordCredentials' (& $sortedBy $ReferenceSnapshot.CredentialsMetadata.PasswordCredentials 'KeyId') (& $sortedBy $DifferenceSnapshot.CredentialsMetadata.PasswordCredentials 'KeyId')
    & $addIfDifferent 'CredentialsMetadata.KeyCredentials' (& $sortedBy $ReferenceSnapshot.CredentialsMetadata.KeyCredentials 'KeyId') (& $sortedBy $DifferenceSnapshot.CredentialsMetadata.KeyCredentials 'KeyId')

    return $rows.ToArray()
}
