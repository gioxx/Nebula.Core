@{
    RootModule           = 'Nebula.Core.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = '07acc3c0-14dc-4c1d-a1d0-6140e83c2a41'
    Author               = 'Giovanni Solone'
    Description          = 'A PowerShell module that go beyond your workstations. It will make your Microsoft 365 life easier!'

    # Minimum required PowerShell (PS 5.1 works; better with PS 7+)
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredAssemblies   = @()
    FunctionsToExport    = @(
        'Connect-EOL',
        'Connect-Nebula',
        'Disconnect-Nebula',
        'Export-MboxStatistics',
        'Export-MsolAccountSku',
        'Get-UserMsolAccountSku',
        'Update-LicenseCatalog'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @('Leave-Nebula')

    PrivateData          = @{
        PSData = @{
            Tags         = @('Microsoft', 'PowerShell', 'Microsoft 365', 'Office 365', 'Exchange', 'Microsoft Entra')
            ProjectUri   = 'https://github.com/gioxx/Nebula.Core'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/Assets/icon.png'
            ReleaseNotes = @'
- Hello World :-) This is the first public version of the module!
'@
        }
    }
}
