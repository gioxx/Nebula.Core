@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.2.0'
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
        'Get-EntraGroupDevice',
        'Get-EntraGroupMembers',
        'Get-EntraGroupUser',
        'Get-MboxAlias',
        'Get-MboxLastMessageTrace',
        'Get-MboxPermission',
        'Get-MboxPrimarySmtpAddress',
        'Get-MboxStatistics',
        'Get-NebulaConfig',
        'Get-NebulaConnections',
        'Get-NebulaModuleUpdates',
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
        'Search-EntraGroup',
        'Set-MboxLanguage',
        'Set-MboxRulesQuota',
        'Set-OoO',
        'Set-SharedMboxCopyForSent',
        'Sync-NebulaConfig',
        'Test-SharedMailboxCompliance',
        'Unlock-QuarantineFrom',
        'Unlock-QuarantineMessageId',
        'Update-LicenseCatalog',
        'Update-NebulaConnections'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @(
        'Export-DDG',
        'Export-DG',
        'fse',
        'Get-DDGRecipientFilter',
        'gpa',
        'Leave-Nebula',
        'mids',
        'qrel',
        'rqf'
    )

    PrivateData          = @{
        PSData = @{
            Tags         = @(
                'Administration',
                'Automation',
                'Calendar',
                'Configuration',
                'Entra', 
                'Exchange', 
                'Exchange-Online',
                'Groups', 
                'Licenses', 
                'M365', 
                'Mailboxes', 
                'Microsoft', 
                'Microsoft-365', 
                'Microsoft-Graph',
                'Office-365', 
                'PowerShell', 
                'Quarantine', 
                'Reporting',
                'Rooms',
                'Security'
            )
            ProjectUri   = 'https://github.com/gioxx/Nebula.Core'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/Assets/icon.png'
ReleaseNotes = @'
- Change: Remove-MboxPermission now uses -ClearAll (renamed from -RemoveAllAdditionalPermissions).
- Improve: Get-EntraGroupMembers can resolve registered owners/users for device members via -IncludeDeviceUsers.
- Improve: Get-EntraGroupMembers reports device owners/users in a single column when resolved.
- Improve: Get-NebulaModuleUpdates now also checks ExchangeOnlineManagement and Microsoft.Graph meta modules.
- Improve: Get-TenantMsolAccountSku now reports Available net of suspended seats and shows Total with enabled/suspended breakdown.
- Improve: Get-UserMsolAccountSku can show tenant availability for assigned SKUs via -CheckAvailability.
- Improve: Remove-EntraGroupDevice/Remove-EntraGroupUser can clear all group members via -ClearAll (with stronger confirmation).
- New: Get-EntraGroupMembers lists all members of an Entra group (users, devices, ...).
- New: Get-MboxStatistics returns a simplified mailbox statistics view.
- New: Get-NebulaModuleUpdates runs an on-demand update check for Nebula.* modules.
- New: Get-TenantMsolAccountSku adds TotalCount with the numeric total for scripting.
- New: Search-EntraGroup searches Entra groups by display name or description.
- New: Update checks during Connect-Nebula can be throttled via CheckUpdatesIntervalHours.
- New: Update-NebulaConnections adds an explicit refresh entry point for connection status checks.
'@
        }
    }
}
