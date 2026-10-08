BeforeAll {
    function Write-NCMessage {
        param(
            [string]$Message,
            [string]$Level
        )
    }
    . "$PSScriptRoot/../../Private/NC-Hlp.ModuleUtils.ps1"
}

Describe 'Test-NCGridViewSupport' {
    AfterEach {
        $script:NCGridViewBrokenVersions = @('7.6.6')
    }

    It 'lists PowerShell 7.6.6 as broken' {
        $script:NCGridViewBrokenVersions | Should -Contain '7.6.6'
    }

    It 'returns true when the current release is not listed as broken' {
        $script:NCGridViewBrokenVersions = @('0.0.0')

        Test-NCGridViewSupport | Should -BeTrue
    }

    It 'returns false when the current release is listed as broken' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
        $version = $PSVersionTable.PSVersion
        $script:NCGridViewBrokenVersions = @('{0}.{1}.{2}' -f $version.Major, $version.Minor, $version.Patch)

        Test-NCGridViewSupport | Should -BeFalse
    }
}

Describe 'Out-NCGridView' {
    BeforeEach {
        Mock Write-NCMessage {}
        Mock Out-GridView { $InputObject }
    }

    Context 'when Out-GridView works' {
        BeforeEach {
            Mock Test-NCGridViewSupport { $true }
        }

        It 'pipes every row to Out-GridView' {
            $rows = @([pscustomobject]@{ Name = 'a' }, [pscustomobject]@{ Name = 'b' })

            $result = @($rows | Out-NCGridView -Title 'Rows')

            $result.Name | Should -Be @('a', 'b')
            Should -Invoke Out-GridView -Times 2 -Exactly -Scope It -ParameterFilter { $Title -eq 'Rows' }
            Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It
        }

        It 'forwards -PassThru to Out-GridView' {
            $null = [pscustomobject]@{ Name = 'a' } | Out-NCGridView -Title 'Rows' -PassThru

            Should -Invoke Out-GridView -Times 1 -Exactly -Scope It -ParameterFilter { $PassThru }
        }
    }

    Context 'when Out-GridView hangs' {
        BeforeEach {
            Mock Test-NCGridViewSupport { $false }
        }

        It 'returns the rows to the console with a warning instead of opening the grid' {
            $rows = @([pscustomobject]@{ Name = 'a' }, [pscustomobject]@{ Name = 'b' })

            $result = @($rows | Out-NCGridView -Title 'Rows')

            $result.Count | Should -Be 2
            $result[0].Name | Should -Be 'a'
            Should -Invoke Out-GridView -Times 0 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -like '*console*' }
        }

        It 'selects nothing with -PassThru so callers do not act on every row' {
            $rows = @([pscustomobject]@{ Name = 'a' }, [pscustomobject]@{ Name = 'b' })

            $result = @($rows | Out-NCGridView -Title 'Rows' -PassThru)

            $result.Count | Should -Be 0
            Should -Invoke Out-GridView -Times 0 -Exactly -Scope It
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -like '*nothing was selected*' }
        }
    }
}
