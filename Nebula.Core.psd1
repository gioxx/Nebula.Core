@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.3.1'
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
        'Get-QuarantineForMailbox',
        'Get-QuarantineFrom',
        'Get-QuarantineFromDomain',
        'Get-QuarantineToRelease',
        'Get-RoleGroupsMembers',
        'Get-RoomDetails',
        'Get-TenantMsolAccountSku',
        'Get-UserDevices',
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
        'Set-UserMsolAccountSku',
        'Set-UserUsageLocation',
        'Switch-UserMsolAccountSku',
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
Full release notes: https://github.com/gioxx/Nebula.Core/releases/tag/v1.3.1

Improvements
- Update check: on PowerShell 7 it now also reports a newer stable PowerShell release (read from gioxx/SWUpdates-Alert), since PowerShell 7.6.8 no longer shows its own update notice. Turn it off with CheckPowerShellUpdates = $false.

Fixes
- License catalog cache: the 7-day check against GitHub now counts from the last real check, not the last use, so new catalog versions are downloaded again. When the commit date cannot be read, the download time is no longer recorded as the catalog's "last updated" date (it blocked later updates). Run any license cmdlet once with `-ForceLicenseCatalogRefresh` to repair an existing cache.
- Connect-Nebula: a canceled Microsoft Graph sign-in now shows a short message instead of the raw Connect-MgGraph error, and on PowerShell 7 connection failures print one line instead of the source position block.
'@
        }
    }
}
