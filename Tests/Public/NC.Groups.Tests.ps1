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
function Get-MgGroup {
    param(
        [string]$GroupId,
        [string]$Filter,
        [switch]$All,
        [string[]]$Property
    )
}
function Get-MgUser {
    param(
        [string]$UserId,
        [string]$Filter,
        [switch]$All
    )
}
function Find-UserRecipient {
    param(
        [string]$UserPrincipalName,
        [switch]$PreferGraphIdentity
    )
}
function New-MgGroupMember {
    param(
        [string]$GroupId,
        [string]$DirectoryObjectId
    )
}
function Get-MgUserMemberOf {
    param(
        [string]$UserId,
        [switch]$All
    )
}
function Remove-MgGroupMemberByRef {
    param(
        [string]$GroupId,
        [string]$DirectoryObjectId
    )
}

. "$PSScriptRoot/../../Public/NC.Groups.ps1"

Describe 'Entra group user identity resolution' {
    $memberUpn = 'employee@contoso.com'
    $guestMail = 'consultant@external.example'
    $memberId = '11111111-1111-1111-1111-111111111111'
    $guestId = '22222222-2222-2222-2222-222222222222'
    $groupId = '33333333-3333-3333-3333-333333333333'

    $memberUser = [pscustomobject]@{
        Id                = $memberId
        UserPrincipalName = $memberUpn
        DisplayName       = 'Tenant Employee'
    }
    $guestUser = [pscustomobject]@{
        Id                = $guestId
        UserPrincipalName = 'consultant_external.example#EXT#@contoso.onmicrosoft.com'
        DisplayName       = 'External Consultant'
    }
    $group = [pscustomobject]@{
        Id                    = $groupId
        DisplayName           = 'Cloud Group'
        OnPremisesSyncEnabled = $false
    }
    $membership = [pscustomobject]@{
        Id                   = $groupId
        AdditionalProperties = @{
            displayName = 'Cloud Group'
            mail        = 'cloud-group@contoso.com'
        }
    }

    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Get-MgGroup { $group }
        Mock Find-UserRecipient {
            if ($PreferGraphIdentity) {
                return $guestId
            }

            return $UserPrincipalName
        }
        Mock New-MgGroupMember {}
        Mock Get-MgUserMemberOf { @($membership) }
        Mock Remove-MgGroupMemberByRef {}
    }

    It 'adds a tenant member through the unchanged direct lookup' {
        Mock Get-MgUser {
            if ($UserId -eq $memberUpn) {
                return $memberUser
            }

            throw "Unexpected user lookup: $UserId"
        }

        Add-EntraGroupUser -GroupName $group.DisplayName -UserIdentifier $memberUpn -Confirm:$false

        Assert-MockCalled Find-UserRecipient -Times 0 -Scope It
        Assert-MockCalled New-MgGroupMember -Times 1 -Scope It -ParameterFilter {
            $GroupId -eq $groupId -and $DirectoryObjectId -eq $memberId
        }
    }

    It 'adds a guest through a Graph-compatible fallback identity' {
        Mock Get-MgUser {
            if ($UserId -eq $guestId) {
                return $guestUser
            }

            throw "User not found: $UserId"
        }

        Add-EntraGroupUser -GroupName $group.DisplayName -UserIdentifier $guestMail -Confirm:$false

        Assert-MockCalled Find-UserRecipient -Times 1 -Scope It -ParameterFilter {
            $UserPrincipalName -eq $guestMail -and $PreferGraphIdentity
        }
        Assert-MockCalled New-MgGroupMember -Times 1 -Scope It -ParameterFilter {
            $GroupId -eq $groupId -and $DirectoryObjectId -eq $guestId
        }
    }

    It 'reads memberships for a tenant member through the unchanged direct lookup' {
        Mock Get-MgUser {
            if ($UserId -eq $memberUpn) {
                return $memberUser
            }

            throw "Unexpected user lookup: $UserId"
        }

        $null = Get-EntraGroupUser -UserIdentifier $memberUpn

        Assert-MockCalled Find-UserRecipient -Times 0 -Scope It
        Assert-MockCalled Get-MgUserMemberOf -Times 1 -Scope It -ParameterFilter {
            $UserId -eq $memberId
        }
    }

    It 'reads memberships for a guest through a Graph-compatible fallback identity' {
        Mock Get-MgUser {
            if ($UserId -eq $guestId) {
                return $guestUser
            }

            throw "User not found: $UserId"
        }

        $null = Get-EntraGroupUser -UserIdentifier $guestMail

        Assert-MockCalled Find-UserRecipient -Times 1 -Scope It -ParameterFilter {
            $UserPrincipalName -eq $guestMail -and $PreferGraphIdentity
        }
        Assert-MockCalled Get-MgUserMemberOf -Times 1 -Scope It -ParameterFilter {
            $UserId -eq $guestId
        }
    }

    It 'removes a tenant member through the unchanged direct lookup' {
        Mock Get-MgUser {
            if ($UserId -eq $memberUpn) {
                return $memberUser
            }

            throw "Unexpected user lookup: $UserId"
        }

        Remove-EntraGroupUser -GroupName $group.DisplayName -UserIdentifier $memberUpn -Confirm:$false

        Assert-MockCalled Find-UserRecipient -Times 0 -Scope It
        Assert-MockCalled Remove-MgGroupMemberByRef -Times 1 -Scope It -ParameterFilter {
            $GroupId -eq $groupId -and $DirectoryObjectId -eq $memberId
        }
    }

    It 'removes a guest through a Graph-compatible fallback identity' {
        Mock Get-MgUser {
            if ($UserId -eq $guestId) {
                return $guestUser
            }

            throw "User not found: $UserId"
        }

        Remove-EntraGroupUser -GroupName $group.DisplayName -UserIdentifier $guestMail -Confirm:$false

        Assert-MockCalled Find-UserRecipient -Times 1 -Scope It -ParameterFilter {
            $UserPrincipalName -eq $guestMail -and $PreferGraphIdentity
        }
        Assert-MockCalled Remove-MgGroupMemberByRef -Times 1 -Scope It -ParameterFilter {
            $GroupId -eq $groupId -and $DirectoryObjectId -eq $guestId
        }
    }
}
