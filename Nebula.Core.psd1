@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.2.1'
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
        'Export-MboxAlias',
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
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/Assets/icon.png'
            ReleaseNotes = @'
- Change: Added `Export-IntuneAppInventory` for Intune app inventory reporting, with optional deployed-app status enrichment and CSV/JSON export.
- Change: Added `New-IntuneAppBasedGroup` to create or update Entra security groups from Intune-managed devices and installed applications, with support for an explicit full group name that aggregates all matches into one group.
- Change: Added `Search-IntuneProfileLocation` to locate which Intune Graph surface hosts a profile and return its source, ID, and OData type.
- Change: Added an optional `-Domain` filter to `Export-MsolAccountSku` so exports can be limited to users in a specific domain, matching `Mail`, `UserPrincipalName`, and `ProxyAddresses`.
- Change: Added resilient Exchange Online connection handling in `Connect-EOL`, including optional `-DisableWAM`, `-Device`, `-NoWamFallback`, and automatic retry without WAM after broker-related sign-in failures.
- Change: Refactored Intune group usage logic into dedicated private helpers to keep `NC.Intune.ps1` focused on public cmdlets.
- Fix: `Get-UserMsolAccountSku -Clipboard` no longer claims success when user lookup or license retrieval fails; it now warns when there is no license data to copy.
- Fix: Quarantine workflows now benefit from the improved EXO reconnection path when WAM/MSAL broker state breaks after idle, lock, or sleep.
- Fix: Reworked `Get-IntuneProfileAssignmentsByGroup` to correctly report Entra group usage across Intune device configurations, settings catalog policies, and app assignments.
- Improve: Added support for nested group matching, diagnostic output, mixed include/exclude aggregation, and console highlighting for exclusion rows in Intune group usage results.
- Improve: License user resolution now prefers Microsoft Graph identity (via shared `Find-UserRecipient -PreferGraphIdentity`) in `Add/Copy/Move/Get/Remove-UserMsolAccountSku`, with better handling of object IDs and hybrid alias lookups.
'@
        }
    }
}
