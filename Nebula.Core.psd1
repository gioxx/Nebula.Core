@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.2.3'
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
        'Add-EntraGroupOwner',
        'Add-EntraGroupUser',
        'Add-MboxAlias',
        'Add-MboxPermission',
        'Add-UserMsolAccountSku',
        'Compare-EnterpriseApplication',
        'Connect-EOL',
        'Connect-Nebula',
        'Copy-EnterpriseApplication',
        'Copy-EntraGroup',
        'Copy-EntraGroupOwner',
        'Copy-OoOMessage',
        'Copy-UserMsolAccountSku',
        'Disable-UserDevices',
        'Disable-UserSignIn',
        'Disconnect-Nebula',
        'Edit-ContentFilterPolicy',
        'Export-CalendarPermission',
        'Export-DistributionGroups',
        'Export-DynamicDistributionGroups',
        'Export-EmptyEntraGroups',
        'Export-EnterpriseApplication',
        'Export-IntuneAppInventory',
        'Export-M365Group',
        'Export-MboxDeletedItemSize',
        'Export-MboxPermission',
        'Export-MboxStatistics',
        'Export-MsolAccountSku',
        'Export-QuarantineEml',
        'Format-MessageIDsFromClipboard',
        'Format-QuotedListFromClipboard',
        'Format-SortedEmailsFromClipboard',
        'Get-ContentFilterPolicy',
        'Get-DynamicDistributionGroupFilter',
        'Get-EntraGroupDevice',
        'Get-EntraGroupMembers',
        'Get-EntraGroupUser',
        'Get-IntuneProfileAssignmentsByGroup',
        'Get-MboxAlias',
        'Get-MboxLastMessageTrace',
        'Get-MboxMrmCleanup',
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
        'Get-UserUsageLocation',
        'Import-EnterpriseApplication',
        'Move-UserMsolAccountSku',
        'New-EntraSecurityGroup',
        'New-IntuneAppBasedGroup',
        'New-SharedMailbox',
        'Remove-EntraGroupDevice',
        'Remove-EntraGroupOwner',
        'Remove-EntraGroupUser',
        'Remove-MboxAlias',
        'Remove-MboxMrmCleanup',
        'Remove-MboxPermission',
        'Remove-UserMsolAccountSku',
        'Revoke-UserSessions',
        'Remove-EntraUser',
        'Search-EntraUser',
        'Search-EntraGroup',
        'Search-IntuneProfileLocation',
        'Search-MboxCutoffWindow',
        'Set-EntraGroupDescription',
        'Set-EntraGroupDisplayName',
        'Set-MboxLanguage',
        'Set-MboxMrmCleanup',
        'Set-MboxRulesQuota',
        'Set-OoO',
        'Set-SharedMboxCopyForSent',
        'Set-UserUsageLocation',
        'Sync-NebulaConfig',
        'Test-SharedMailboxCompliance',
        'Get-IntuneAppPresence',
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
                'App-Registration',
                'Automation',
                'Calendar',
                'Configuration',
                'Enterprise-Applications',
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
                'Security',
                'Service-Principal'
            )
            ProjectUri   = 'https://github.com/gioxx/Nebula.Core'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/icon.png'
ReleaseNotes = @'
- Add: `Export-EnterpriseApplication`, `Import-EnterpriseApplication`, `Copy-EnterpriseApplication`, and `Compare-EnterpriseApplication` to snapshot, recreate, clone, and diff Enterprise Applications (App Registration + Service Principal) within the same Entra tenant, including optional App Role Assignment copying, owner sync, and CSV/JSON diff reports. Client secrets and certificates are never copied; only their metadata is captured for reporting.
- Fix: `Compare-EnterpriseApplication`'s CSV report now renders non-scalar diff values (redirect URIs, permissions, app roles, owners) as readable JSON instead of an identical, uninformative string on both sides, and now honors the module's configured CSV encoding and delimiter.
- Improve: the Pester test suite (`Tests/Public/*.Tests.ps1`) now correctly scopes its fixtures inside `BeforeAll` blocks so tests run for real under Pester 5, instead of only succeeding at test discovery.
'@
        }
    }
}
