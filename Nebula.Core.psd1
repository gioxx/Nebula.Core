@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.2.2'
    GUID                 = '07acc3c0-14dc-4c1d-a1d0-6140e83c2a41'
    Author               = 'Giovanni Solone'
    Description          = 'A PowerShell module that go beyond your workstations. It will make your Microsoft 365 life easier!'

    # Minimum required PowerShell (PS 5.1 works; better with PS 7+)
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredAssemblies   = @()
    FormatsToProcess     = @(
        'Formats\Nebula.Core.Format.ps1xml'
    )
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
        'Export-EmptyEntraGroups',
        'Export-IntuneAppInventory',
        'Export-M365Group',
        'Export-MboxDeletedItemSize',
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
        'Get-IntuneProfileAssignmentsByGroup',
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
        'New-IntuneAppBasedGroup',
        'Remove-EntraGroupDevice',
        'Remove-EntraGroupUser',
        'Remove-MboxAlias',
        'Remove-MboxPermission',
        'Remove-UserMsolAccountSku',
        'Revoke-UserSessions',
        'Search-EntraGroup',
        'Search-IntuneProfileLocation',
        'Search-MboxCutoffWindow',
        'Set-MboxLanguage',
        'Set-MboxMrmCleanup',
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
                'Intune',
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
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/icon.png'
            ReleaseNotes = @'
- Improve: `Get-TenantMsolAccountSku` now supports optional `-Domain` filtering for sample users, so license sample output can be scoped to a specific mail domain.
- Change: Removed `Export-MboxAlias`. Use `Get-MboxAlias` for single mailbox queries and CSV reports, including positional single-mailbox input.
- Fix: Consolidated alias export behavior under `Get-MboxAlias` and fixed CSV filtering so primary-only recipients are excluded after MOERA filtering.
- Improve: `Get-MboxAlias` can now include DisplayName and Name in CSV exports, optionally include primary-only recipients with `-IncludePrimaryOnly`, and opt into MOERA rows with `-IncludeMoera`.
- Improve: CSV export cmdlets in Calendar, Groups, Licenses and Statistics now consistently finish with a success message that includes the generated CSV path instead of echoing the path as a second pipeline line.
'@
        }
    }
}
