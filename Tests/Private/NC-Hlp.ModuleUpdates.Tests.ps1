BeforeAll {
    function Write-NCMessage { param([string]$Message, [string]$Level) }

    . "$PSScriptRoot/../../Private/NC-Hlp.ModuleUpdates.ps1"
}

Describe 'Get-NCPowerShellUpdate' {
    BeforeEach {
        Mock Write-NCMessage {}
    }

    It 'returns the update when a newer stable release exists' {
        Mock Invoke-RestMethod { "v7.6.9`n" }

        $result = Get-NCPowerShellUpdate -CurrentVersion '7.6.8' -Edition 'Core'

        $result.Name | Should -Be 'PowerShell'
        $result.InstalledVersion | Should -Be ([version]'7.6.8')
        $result.LatestVersion | Should -Be ([version]'7.6.9')
        Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter { $Uri -eq 'https://raw.githubusercontent.com/gioxx/SWUpdates-Alert/main/updates/LATEST_PWSH' }
    }

    It 'returns nothing when the running version is current' {
        Mock Invoke-RestMethod { 'v7.6.9' }

        Get-NCPowerShellUpdate -CurrentVersion '7.6.9' -Edition 'Core' | Should -BeNullOrEmpty
    }

    It 'returns nothing when the running version is ahead (preview)' {
        Mock Invoke-RestMethod { 'v7.6.9' }

        Get-NCPowerShellUpdate -CurrentVersion '7.7.0' -Edition 'Core' | Should -BeNullOrEmpty
    }

    It 'uses the running version when none is given' {
        Mock Invoke-RestMethod { 'v999.0.0' }

        $result = Get-NCPowerShellUpdate -Edition 'Core'

        $v = $PSVersionTable.PSVersion
        $result.InstalledVersion | Should -Be ([version]::new($v.Major, $v.Minor, $v.Patch))
    }

    It 'warns and returns nothing when the download fails' {
        Mock Invoke-RestMethod { throw 'network down' }

        Get-NCPowerShellUpdate -CurrentVersion '7.6.8' -Edition 'Core' | Should -BeNullOrEmpty
        Should -Invoke Write-NCMessage -Times 1 -ParameterFilter { $Level -eq 'WARNING' -and $Message -like '*network down*' }
    }

    It 'warns and returns nothing when the content is not a version' {
        Mock Invoke-RestMethod { '<html>Not Found</html>' }

        Get-NCPowerShellUpdate -CurrentVersion '7.6.8' -Edition 'Core' | Should -BeNullOrEmpty
        Should -Invoke Write-NCMessage -Times 1 -ParameterFilter { $Level -eq 'WARNING' }
    }

    It 'skips Windows PowerShell without any request' {
        Mock Invoke-RestMethod { 'v7.6.9' }

        Get-NCPowerShellUpdate -CurrentVersion '5.1.0' -Edition 'Desktop' | Should -BeNullOrEmpty
        Should -Invoke Invoke-RestMethod -Times 0
    }
}

Describe 'Test-NebulaModuleUpdates PowerShell check' {
    BeforeAll {
        function Get-InstalledModule { param($Name, $ErrorAction) }
        function Find-Module { param($Name, $Repository, $ErrorAction) }
    }

    BeforeEach {
        $global:NCVars = @{ CheckPowerShellUpdates = $true; UserConfigRoot = $TestDrive }
        Mock Write-NCMessage {}
        Mock Save-NCModuleUpdateLastCheck {}
        Mock Get-NCPowerShellUpdate { [pscustomobject]@{ Name = 'PowerShell'; InstalledVersion = [version]'7.6.8'; LatestVersion = [version]'7.6.9' } }
        Mock Get-InstalledModule { [pscustomobject]@{ Name = 'Nebula.Core'; Version = '1.3.1' } } -ParameterFilter { $Name -eq 'Nebula.*' }
        Mock Get-InstalledModule { } -ParameterFilter { $Name -ne 'Nebula.*' }
        Mock Find-Module { [pscustomobject]@{ Name = 'Nebula.Core'; Version = '1.3.1' } }
    }

    AfterAll {
        Remove-Variable -Name NCVars -Scope Global -ErrorAction SilentlyContinue
    }

    It 'reports a PowerShell update when modules are current' {
        Test-NebulaModuleUpdates -Force | Should -BeTrue
        Should -Invoke Write-NCMessage -Times 1 -ParameterFilter { $Level -eq 'WARNING' -and $Message -eq 'Update available for PowerShell: 7.6.8 -> 7.6.9' }
    }

    It 'still reports the PowerShell update when the Gallery is unreachable' {
        Mock Find-Module { throw 'Gallery down' }

        Test-NebulaModuleUpdates -Force | Should -BeTrue
        Should -Invoke Write-NCMessage -Times 1 -ParameterFilter { $Message -eq 'Update available for PowerShell: 7.6.8 -> 7.6.9' }
    }

    It 'returns false when nothing is outdated' {
        Mock Get-NCPowerShellUpdate { $null }

        Test-NebulaModuleUpdates -Force | Should -BeFalse
    }

    It 'skips the PowerShell check when <Value> is configured' -ForEach @(
        @{ Value = $false }
        @{ Value = 'false' }
    ) {
        $global:NCVars.CheckPowerShellUpdates = $Value

        Test-NebulaModuleUpdates -Force | Should -BeFalse
        Should -Invoke Get-NCPowerShellUpdate -Times 0
    }

    It 'accepts the string value yes' {
        $global:NCVars.CheckPowerShellUpdates = 'yes'

        Test-NebulaModuleUpdates -Force | Should -BeTrue
    }
}
