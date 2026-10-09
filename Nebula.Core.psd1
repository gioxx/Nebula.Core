@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.3.0'
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
Full release notes: https://github.com/gioxx/Nebula.Core/releases/tag/v1.3.0

Before you update
- Default changes: CSV files now use a comma delimiter (was `;`), and dates are shown in `Eastern Standard Time` unless `DateTimeTimeZone` is set. Restore the previous behavior with `CSV_DefaultLimiter = ';'` and/or your own `DateTimeTimeZone` in `settings.psd1` (see `Get-NebulaConfig`).
- On PowerShell 7.6.6 `Out-GridView` hangs (PowerShell/PowerShell#27994): every `-GridView` writes the results to the console there instead.

New commands
- `Export-`, `Import-`, `Copy-` and `Compare-EnterpriseApplication`: snapshot, recreate, clone and diff Enterprise Applications (App Registration + Service Principal) in the same tenant. The docs list exactly which settings are copied; owners and App Role Assignments are only added; SAML/password SSO, identifier URIs and secrets/certificates are not copied.
- `Get-UserDevices`: one table with a user's Entra (registered/owned) and Intune devices, matched on the Azure AD device ID, including hybrid-joined PCs.
- `Get-QuarantineForMailbox`: quarantine across all of a mailbox's SMTP aliases (15-day default window).
- `Get-IntuneAppPresence`, `Search-EntraUser`, `Remove-EntraUser`.

Improvements
- Bulk and pipeline cmdlets for groups, licenses, security, users, mailbox sign-in checks and Intune send Microsoft Graph requests through `$batch` (20 per call) and read tenant-wide data once per run; far fewer WAM prompts with WAM-per-request Graph builds.
- Graph requests use URIs relative to the active environment, so national clouds (e.g. USGov) work.
- `Connect-Nebula` connects Graph before Exchange Online (WAM-disabled EXO sign-in) to avoid the assembly/broker clash; `-GraphLoginHint` (Microsoft.Graph.Authentication 2.39.0+) targets an account; an active session keeps its account and tenant.
- `Update-NebulaConnections` probes and repairs unhealthy sessions (Graph first), keeping each session's account and tenant.
- `Connect-EOL` hides the cosmetic WAM notice and only passes `-DisableWAM`/`-Device` when ExchangeOnlineManagement supports them.
- User lookups query Graph first and fall back to Exchange only when needed.
- `New-IntuneAppBasedGroup` adds devices one reference at a time and adds duplicate Entra devices once.
- Culture-safe date parsing and time-zone-aware formatting (`DateTimeTimeZone`); date-only values are shown as given.
- `Get-MboxPermission` shows the source `RecipientTypeDetails`; the `LicenseMapping` action flags stale custom mappings.

Fixes
- `New-IntuneAppBasedGroup -UpdateExisting` never removes members when device resolution is incomplete, a device can't be resolved, or an addition fails; no duplicate group when the lookup fails; `$select` is applied.
- `Add-UserMsolAccountSku` skips licenses a user already has (no seat used, disabled plans kept) and processes repeated users once; `Copy-`/`Move-UserMsolAccountSku` check seat availability per SKU.
- Quarantine cmdlets read every page of results (they stopped at 100 messages).
- Graph collection reads stop with an error when a page fails, instead of returning partial data.
- `Export-IntuneAppInventory` no longer closes the PowerShell session on error.
- `Add/Get/Remove-EntraGroupUser` resolve invited guests by external e-mail; group member/owner cmdlets accept `<GroupName> <MemberIdentifier>`.
- `Copy-EntraGroup`, `Copy-EntraGroupOwner` and `Remove-EntraGroupOwner -ClearAll` read Graph results correctly.
- `Search-EntraUser -SearchIn` honors the selected field; name lookups URI-encode their filters (names with `&` or `#`).
- `Get-IntuneAppPresence` uses the most recently synced record when several share a device name.
- `Test-EOLConnection` keeps an active Exchange Online session instead of switching to the Windows identity.
- License catalog falls back to the cached copy when GitHub is unreachable; invisible characters in SKU names no longer break catalog matching.
- `Get-UserGroups` works for guests through Graph; `Export-IntuneAppInventory` dates match the single-device helper.
'@
        }
    }
}
