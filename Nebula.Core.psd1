@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.1.0'
    GUID                 = '07acc3c0-14dc-4c1d-a1d0-6140e83c2a41'
    Author               = 'Giovanni Solone'
    Description          = 'A PowerShell module that go beyond your workstations. It will make your Microsoft 365 life easier!'

    # Minimum required PowerShell (PS 5.1 works; better with PS 7+)
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredAssemblies   = @()
    FunctionsToExport    = @(
        'Add-MboxAlias',
        'Add-MboxPermission',
        'Add-MsolAccountSku',
        'Connect-EOL',
        'Connect-Nebula',
        'Disconnect-Nebula',
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
        'Get-TenantMsolAccountSku',
        'Get-UserLastSeen',
        'Get-UserGroups',
        'Get-UserMsolAccountSku',
        'Move-MsolAccountSku',
        'New-SharedMailbox',
        'Remove-MboxAlias',
        'Remove-MboxPermission',
        'Set-MboxLanguage',
        'Set-MboxRulesQuota',
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
            Tags         = @('Microsoft', 'PowerShell', 'Microsoft365', 'Office365', 'Exchange', 'Entra')
            ProjectUri   = 'https://github.com/gioxx/Nebula.Core'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/Assets/icon.png'
            ReleaseNotes = @'
- Added Get-TenantMsolAccountSku to list tenant license availability with resolved names and counts.
- Added Move-MsolAccountSku to transfer licenses from one user to another.
- Added Add-MsolAccountSku to assign licenses by friendly name, SKU part number, or SKU ID.
- Added Get-UserLastSeen to retrieve the last sign-in date for users.
'@
        }
    }
}
