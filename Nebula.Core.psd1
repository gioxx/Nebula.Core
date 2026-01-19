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
        'Get-MboxPrimarySmtpAddress',
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
        'gpa',
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
- Fix: Export-QuarantineEml now creates the destination folder when missing and resolves relative paths correctly.
- Improve: Export-QuarantineEml now accepts Identity values (GUID\GUID) in addition to MessageId, supports multiple inputs, and returns all exported items.
- Improve: Format-MessageIDsFromClipboard now outputs the copied quarantine identity values (one per line).
- Improve: Get-UserMsolAccountSku now accepts also pipeline input. Also, it now supports -Clipboard to copy a quoted list of licenses.
- Improve: Remove-UserMsolAccountSku now uses two parameter sets. With -All, it removes all licenses assigned to the user, displaying the names (resolved via catalog if available). The -License parameter remains for selective removal.
- Improve: Get-TenantMsolAccountSku now supports -Filter to show only licenses matching the provided text (name or SkuPartNumber).
- Improve: Remove-MboxPermission now supports -RemoveAllAdditionalPermissions to remove non-inherited FullAccess, SendAs, and SendOnBehalfTo from a mailbox.
- New: Get-MboxPrimarySmtpAddress resolves the PrimarySmtpAddress for any mailbox/recipient, with -Raw for string-only output.
- New: Added alias gpa for Get-MboxPrimarySmtpAddress.
'@
        }
    }
}
