@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.1.1'
    GUID                 = '07acc3c0-14dc-4c1d-a1d0-6140e83c2a41'
    Author               = 'Giovanni Solone'
    Description          = 'A PowerShell module that go beyond your workstations. It will make your Microsoft 365 life easier!'

    # Minimum required PowerShell (PS 5.1 works; better with PS 7+)
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredAssemblies   = @()
    FunctionsToExport    = @(
        'Add-EntraGroupDevice',
        'Add-EntraGroupUser',
        'Add-MboxAlias',
        'Add-MboxPermission',
        'Add-UserMsolAccountSku',
        'Connect-EOL',
        'Connect-Nebula',
        'Copy-OoOMessage',
        'Disable-UserDevices',
        'Disable-UserSignIn',
        'Disconnect-Nebula',
        'Export-CalendarPermission',
        'Export-DistributionGroups',
        'Export-DynamicDistributionGroups',
        'Export-M365Group',
        'Export-MboxAlias',
        'Export-MboxPermission',
        'Export-MboxStatistics',
        'Export-MsolAccountSku',
        'Export-QuarantineEml',
        'Format-MessageIDsFromClipboard',
        'Format-SortedEmailsFromClipboard',
        'Get-DynamicDistributionGroupFilter',
        'Get-MboxAlias',
        'Get-MboxPermission',
        'Get-NebulaConfig',
        'Get-QuarantineFrom',
        'Get-QuarantineFromDomain',
        'Get-QuarantineToRelease',
        'Get-RoleGroupsMembers',
        'Get-RoomDetails',
        'Get-TenantMsolAccountSku',
        'Get-UserGroups',
        'Get-UserLastSeen',
        'Get-UserMsolAccountSku',
        'Move-UserMsolAccountSku',
        'New-SharedMailbox',
        'Remove-MboxAlias',
        'Remove-MboxPermission',
        'Remove-UserMsolAccountSku',
        'Revoke-UserSessions',
        'Set-MboxLanguage',
        'Set-MboxRulesQuota',
        'Set-OoO',
        'Set-SharedMboxCopyForSent',
        'Sync-NebulaConfig',
        'Test-SharedMailboxCompliance',
        'Unlock-QuarantineFrom',
        'Unlock-QuarantineMessageId',
        'Update-LicenseCatalog'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @(
        'Export-DDG',
        'Export-DG',
        'fse',
        'Get-DDGRecipientFilter',
        'Leave-Nebula',
        'mids',
        'qrel',
        'rqf'
    )

    PrivateData          = @{
        PSData = @{
            Tags         = @('Microsoft', 'PowerShell', 'Microsoft365', 'Office365', 'Exchange', 'Entra')
            ProjectUri   = 'https://github.com/gioxx/Nebula.Core'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/Assets/icon.png'
            ReleaseNotes = @'
- Added Add-EntraGroupDevice and Add-EntraGroupUser cmdlets to manage Entra groups.
- Renamed Add-MsolAccountSku and Move-MsolAccountSku to Add-UserMsolAccountSku and Move-UserMsolAccountSku for clarity.
- Added Remove-UserMsolAccountSku cmdlet to remove licenses from users.
- Added Copy-OoOMessage cmdlet to duplicate Out of Office messages between mailboxes, Export-CalendarPermission to export calendar permissions, and Set-OoO to set Out of Office messages.
- Added Disable-UserDevices cmdlet to disable user devices, Disable-UserSignIn cmdlet to disable sign-in for users and Revoke-UserSessions cmdlet to revoke user sessions.
- Added Get-RoomDetails cmdlet to retrieve detailed information about meeting rooms.
'@
        }
    }
}
