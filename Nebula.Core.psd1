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
        'Connect-EOL',
        'Connect-Nebula',
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
- Fix: `Add/Remove-EntraGroupDevice`, `Add/Remove-EntraGroupOwner`, and `Add/Remove-EntraGroupUser` now support the positional form `<GroupName> <MemberIdentifier>` in addition to named parameters.
- Fix: `Add/Get/Remove-EntraGroupUser` now resolve invited Entra guests by external e-mail through a Graph-compatible fallback while preserving direct lookup for tenant members.
- Fix: `Export-IntuneAppInventory` now normalizes cached `LastInventory` values through Nebula's configured date/time formatter so export output matches the single-device helper.
- Fix: `Get-UserGroups` now falls back to Microsoft Graph resolution when Exchange mailbox lookup is not available, so Entra guest users can be queried without using the GUI.
- Improve: `Get-MboxPermission` now shows the source mailbox `RecipientTypeDetails` value in the output heading.
- Improve: `Get-UserGroups` keeps the existing Exchange-first behavior for regular users while handling guest identities more gracefully.
- Improve: add `Get-IntuneAppPresence` for quick single-device app presence checks with one-row output, always include `LastInventory`, and return the matched app name in `AppName`.
- Improve: add `Remove-EntraUser` for direct UPN-based Entra user removal with Graph.
- Improve: add `Search-EntraUser` to search users by display name, user principal name, or mail, including guest UPN fragments.
- Improve: add culture-safe date parsing plus optional timezone-aware formatting through `DateTimeTimeZone`.
- Improve: change the default CSV delimiter to comma for a more standard US-friendly baseline.
- Improve: set the default user-facing date/time zone to `Eastern Standard Time` to align with the module's US baseline.
- Improve: unify user-facing date formatting through Nebula's configured date/time patterns, including Intune inventory and license catalog outputs.
'@
        }
    }
}
