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
