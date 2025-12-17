@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.1.2'
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
        'Copy-UserMsolAccountSku',
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
        'Remove-EntraGroupDevice',
        'Remove-EntraGroupUser',
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
            Tags         = @(
                'Entra', 
                'Exchange', 
                'Groups', 
                'Licenses', 
                'M365', 
                'Mailboxes', 
                'Microsoft', 
                'Microsoft-365', 
                'Office-365', 
                'PowerShell', 
                'Quarantine', 
                'Rooms'
            )
            ProjectUri   = 'https://github.com/gioxx/Nebula.Core'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/Assets/icon.png'
            ReleaseNotes = @'
- Fix: Set-UsageLocation now correctly updates the usage location for users without one set (licenses functions).
- Fix: Added fallback on attempts/max also in other functions that used Invoke-NCRetry logic.
- Improve: Get-UserMsolAccountSku now accepts also pipeline input.
- Improve: Remove-UserMsolAccountSku now uses two parameter sets. With -All, it removes all licenses assigned to the user, displaying the names (resolved via catalog if available). The -License parameter remains for selective removal.
'@
        }
    }
}
