#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Group helpers ==============================================================================================================

function Get-NCGraphObjectLabel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $props = $InputObject.PSObject.Properties
    if ($props['userPrincipalName'] -and $InputObject.userPrincipalName) { return [string]$InputObject.userPrincipalName }
    if ($props['displayName'] -and $InputObject.displayName) { return [string]$InputObject.displayName }
    if ($props['appDisplayName'] -and $InputObject.appDisplayName) { return [string]$InputObject.appDisplayName }
    if ($props['id'] -and $InputObject.id) { return [string]$InputObject.id }
    return [string]$InputObject
}

function Resolve-NCEntraGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [string]$GroupId
    )

    if (-not [string]::IsNullOrWhiteSpace($GroupId)) {
        try {
            return Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Entra group with ID '$GroupId' not found: $($_.Exception.Message)" -Level ERROR
            return
        }
    }

    $escapedName = $GroupName.Replace("'", "''")
    try {
        $resolvedGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        Write-NCMessage "Unable to resolve group '$GroupName': $($_.Exception.Message)" -Level ERROR
        return
    }

    if (-not $resolvedGroup) {
        try {
            $resolvedGroup = Get-MgGroup -GroupId $GroupName -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Entra group '$GroupName' not found by name or ID" -Level ERROR
            return
        }
    }

    return $resolvedGroup
}

function Resolve-NCEntraOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OwnerIdentifier,

        [switch]$TreatInputAsId
    )

    if ([string]::IsNullOrWhiteSpace($OwnerIdentifier)) {
        return
    }

    $owner = $null
    $trimmed = $OwnerIdentifier.Trim()
    $looksLikeGuid = $trimmed -match '^[0-9a-fA-F-]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

    if ($TreatInputAsId.IsPresent -or $looksLikeGuid) {
        try {
            $owner = Get-MgUser -UserId $trimmed -Property Id,UserPrincipalName,DisplayName -ErrorAction Stop
        }
        catch {
            $owner = [pscustomobject]@{
                Id = $trimmed
                UserPrincipalName = $null
                DisplayName = $trimmed
            }
        }
    }
    else {
        try {
            $owner = Get-MgUser -UserId $trimmed -Property Id,UserPrincipalName,DisplayName -ErrorAction Stop
        }
        catch {
            $resolvedIdentifier = Find-UserRecipient -UserPrincipalName $trimmed -PreferGraphIdentity
            if ($resolvedIdentifier) {
                try {
                    $owner = Get-MgUser -UserId $resolvedIdentifier -Property Id,UserPrincipalName,DisplayName -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to resolve owner '$OwnerIdentifier': $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            else {
                Write-NCMessage "Owner '$OwnerIdentifier' not found." -Level WARNING
                return
            }
        }
    }

    if (-not $owner) {
        Write-NCMessage "Unable to determine object ID for owner '$OwnerIdentifier'." -Level ERROR
        return
    }

    $ownerLabel = if ($owner.UserPrincipalName) { $owner.UserPrincipalName } elseif ($owner.DisplayName) { $owner.DisplayName } else { $owner.Id }

    return [pscustomobject]@{
        Id    = [string]$owner.Id
        Label = [string]$ownerLabel
    }
}

function Resolve-NCEntraGroupUserTarget {
    <#
    .SYNOPSIS
        Resolves group-membership user inputs to object IDs with batched Graph lookups.
    .DESCRIPTION
        Object IDs (or every input when -TreatInputAsId is set) pass through unchanged. Other inputs are
        resolved through Resolve-NCGraphUserBatch; inputs that cannot be resolved are skipped (the resolver
        already reported them).
    .PARAMETER UserIdentifier
        UPNs, mail addresses, display names or object IDs.
    .PARAMETER TreatInputAsId
        Treat every input as an object ID.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$UserIdentifier,
        [switch]$TreatInputAsId
    )

    $guidPattern = '^[0-9a-fA-F-]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    $isId = { param($value) $TreatInputAsId.IsPresent -or $value -match $guidPattern }

    $lookup = @($UserIdentifier | Where-Object { -not (& $isId $_) })
    $resolved = if ($lookup.Count -gt 0) {
        Resolve-NCGraphUserBatch -Identifier $lookup -Property @('id', 'userPrincipalName', 'displayName')
    }
    else {
        @{}
    }

    foreach ($user in $UserIdentifier) {
        if (& $isId $user) {
            [pscustomobject]@{ Input = $user; Id = $user; Label = $user }
            continue
        }

        $match = $resolved[$user]
        if (-not $match) {
            continue
        }
        if (-not $match.id) {
            Write-NCMessage "Unable to determine object ID for user '$user'." -Level ERROR
            continue
        }

        $label = if ($match.userPrincipalName) { $match.userPrincipalName } else { $match.displayName }
        [pscustomobject]@{ Input = $user; Id = $match.id; Label = $label }
    }
}

function Resolve-NCEntraDeviceTargetBatch {
    <#
    .SYNOPSIS
        Resolves device inputs (object IDs or display names) with batched Graph lookups.
    .DESCRIPTION
        Object IDs (or every input when -TreatInputAsId is set) pass through unchanged. Display names are
        looked up with batched GET /devices requests; inputs that cannot be resolved are reported and skipped.
        When several devices match a name the first one is used (with a warning).
    .PARAMETER DeviceIdentifier
        Device object IDs or display names.
    .PARAMETER TreatInputAsId
        Treat every input as an object ID.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$DeviceIdentifier,
        [switch]$TreatInputAsId
    )

    $guidPattern = '^[0-9a-fA-F-]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    $requests = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $DeviceIdentifier.Count; $i++) {
        $value = $DeviceIdentifier[$i]
        if ($TreatInputAsId.IsPresent -or $value -match $guidPattern) { continue }
        $filter = "displayName eq '$($value.Replace("'", "''"))'"
        $requests.Add(@{ Id = "d$i"; Method = 'GET'; Url = "/devices?`$filter=$([uri]::EscapeDataString($filter))&`$select=id,displayName,deviceId" })
    }

    $lookup = @{}
    if ($requests.Count -gt 0) {
        foreach ($result in @(Invoke-NCGraphBatchCollection -Requests @($requests) -Activity 'Resolving devices')) {
            $lookup[$result.Id] = $result
        }
    }

    for ($i = 0; $i -lt $DeviceIdentifier.Count; $i++) {
        $value = $DeviceIdentifier[$i]
        if ($TreatInputAsId.IsPresent -or $value -match $guidPattern) {
            [pscustomobject]@{ Input = $value; Id = $value; Label = $value }
            continue
        }

        $result = $lookup["d$i"]
        if (-not $result.Success) {
            Write-NCMessage "Unable to resolve device '$value': $($result.ErrorMessage)" -Level ERROR
            continue
        }

        $found = @($result.Items)
        if ($found.Count -eq 0) {
            Write-NCMessage "Device '$value' not found" -Level WARNING
            continue
        }

        if ($found.Count -gt 1) {
            Write-NCMessage "Multiple devices matched '$value'. Using the first result ($($found[0].displayName))" -Level WARNING
        }

        $device = $found[0]
        if (-not $device.id) {
            Write-NCMessage "Unable to determine object ID for device '$value'." -Level ERROR
            continue
        }

        [pscustomobject]@{ Input = $value; Id = [string]$device.id; Label = $device.displayName }
    }
}
