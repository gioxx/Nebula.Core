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
- Fix: `Add/Get/Remove-EntraGroupUser` now resolve invited Entra guests by external e-mail through a Graph-compatible fallback while preserving direct lookup for tenant members.
- Fix: `Add/Remove-EntraGroupDevice`, `Add/Remove-EntraGroupOwner`, and `Add/Remove-EntraGroupUser` now support the positional form `<GroupName> <MemberIdentifier>` in addition to named parameters.
- Fix: `Compare-EnterpriseApplication`'s CSV report now renders non-scalar diff values (redirect URIs, permissions, app roles, owners) as readable JSON instead of an identical, uninformative string on both sides, and now honors the module's configured CSV encoding and delimiter.
- Fix: `Connect-EOL` now detects whether the installed `ExchangeOnlineManagement` version actually supports `-DisableWAM`/`-Device` (introduced in 3.7.2) before passing them to `Connect-ExchangeOnline`, instead of always forwarding them; this prevents a parameter-binding error when Exchange Online is intentionally pinned below 3.7.2 as a workaround for the Graph/EOL assembly clash, where WAM isn't the default and those parameters don't exist.
- Improve: `Connect-EOL` suppresses `Connect-ExchangeOnline`'s cosmetic "Sign in by Web Account Manager (WAM) is enabled by default" notice on every WAM sign-in, unless called with `-Verbose`; it never affected the existing WAM-failure detection, which inspects the thrown exception, not warnings.
- Fix: `Connect-Nebula` now initializes Microsoft Graph before Exchange Online and uses a WAM-disabled EXO sign-in for the combined flow, avoiding the known cross-module authentication assembly and broker conflict.
- Fix: `Copy-UserMsolAccountSku` and `Move-UserMsolAccountSku` now check tenant seat availability per SKU before assigning; a license with no available units is skipped (or left on the source, for `Move`) with a warning instead of failing the entire `Set-MgUserLicense` batch and copying/moving nothing.
- Fix: `Export-IntuneAppInventory` now normalizes cached `LastInventory` values through Nebula's configured date/time formatter so export output matches the single-device helper.
- Fix: `Get-UserGroups` now falls back to Microsoft Graph resolution when Exchange mailbox lookup is not available, so Entra guest users can be queried without using the GUI.
- Fix: `Invoke-NCGraphAllPagesCore` now distinguishes a Graph collection with zero items from a non-paged single object instead of relying on truthiness, fixing phantom-item failures (e.g. an app with no owners) in downstream cmdlets like `Copy-EnterpriseApplication`.
- Fix: `Set-NCEnterpriseApplicationFromSnapshot` no longer copies `identifierUris` (unique per tenant, caused Graph BadRequest on apps exposing an API); a warning reports the source value instead.
- Fix: `Set-NCEnterpriseApplicationFromSnapshot` now applies the Service Principal's `Tags` and `Homepage` on both create and update-in-place, so cloned apps correctly show up under Entra's "Enterprise applications" blade.
- Fix: `Set-NCEnterpriseApplicationFromSnapshot` strips the read-only `redirectUriSettings` before writing `web`/`spa`/`publicClient`, avoiding a Graph BadRequest when both are sent together.
- Fix: License catalog download (`Get-LicenseSourceData`) now falls back to the existing stale cache when GitHub is unreachable after all retry attempts, instead of failing outright; a warning reports the fallback and its cache age. The cache is then honored for the configured `LicenseCacheDays` (minimum 1 day) before retrying, instead of re-attempting the full download on every subsequent call.
- Improve: `Get-MboxPermission` now shows the source mailbox `RecipientTypeDetails` value in the output heading.
- Improve: `Get-UserGroups` keeps the existing Exchange-first behavior for regular users while handling guest identities more gracefully.
- Improve: add `Get-IntuneAppPresence` for quick single-device app presence checks with one-row output, always include `LastInventory`, and return the matched app name in `AppName`.
- Improve: add `Remove-EntraUser` for direct UPN-based Entra user removal with Graph.
- Improve: add `Search-EntraUser` to search users by display name, user principal name, or mail, including guest UPN fragments.
- Improve: add culture-safe date parsing plus optional timezone-aware formatting through `DateTimeTimeZone`.
- Improve: change the default CSV delimiter to comma for a more standard US-friendly baseline.
- Improve: set the default user-facing date/time zone to `Eastern Standard Time` to align with the module's US baseline.
- Improve: the Pester test suite (`Tests/Public/*.Tests.ps1`) now correctly scopes its fixtures inside `BeforeAll` blocks so tests run for real under Pester 5, instead of only succeeding at test discovery.
- Improve: unify user-facing date formatting through Nebula's configured date/time patterns, including Intune inventory and license catalog outputs.
'@
        }
    }
}
