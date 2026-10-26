BeforeAll {
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Add-EmptyLine {}
    function Test-EOLConnection { param([switch]$AutoInstall) $true }
    function Get-MgContext {}

    . "$PSScriptRoot/../../Private/NC-Hlp.Connections.ps1"
}

Describe 'Test-MgGraphConnection -LoginHint' {
    BeforeEach {
        Mock Write-NCMessage {}
        Mock Get-Module { [pscustomobject]@{ Name = 'Microsoft.Graph' } } -ParameterFilter { $ListAvailable }
        Mock Import-Module {}
        $script:connected = $false
        Mock Get-MgContext { if ($script:connected) { [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @('User.Read.All') } } }
    }

    Context 'Microsoft.Graph.Authentication 2.39.0 or later' {
        BeforeAll {
            function Connect-MgGraph { param([string[]]$Scopes, [switch]$NoWelcome, [string]$TenantId, [switch]$UseDeviceCode, [string]$LoginHint) }
        }

        It 'passes the login hint to Connect-MgGraph' {
            Mock Connect-MgGraph { $script:connected = $true }

            Test-MgGraphConnection -EnsureExchangeOnline:$false -LoginHint 'admin@contoso.com' | Should -BeTrue

            Should -Invoke Connect-MgGraph -Times 1 -Exactly -Scope It -ParameterFilter { $LoginHint -eq 'admin@contoso.com' }
        }
    }

    Context 'Microsoft.Graph.Authentication before 2.39.0' {
        BeforeAll {
            # Not mocked: a Pester mock accepts undeclared parameters, a real function fails to bind them
            function Connect-MgGraph {
                [CmdletBinding()]
                param([string[]]$Scopes, [switch]$NoWelcome, [string]$TenantId, [switch]$UseDeviceCode)
                $script:connected = $true
            }
        }

        It 'connects without -LoginHint instead of failing parameter binding' {
            Test-MgGraphConnection -EnsureExchangeOnline:$false -LoginHint 'admin@contoso.com' | Should -BeTrue

            Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        }
    }
}

Describe 'Test-MgGraphConnection sign-in canceled' {
    BeforeAll {
        # Connect-MgGraph reports a canceled sign-in as a non-terminating error
        function Connect-MgGraph {
            [CmdletBinding()]
            param([string[]]$Scopes, [switch]$NoWelcome, [string]$TenantId, [switch]$UseDeviceCode, [string]$LoginHint)
            Write-Error 'InteractiveBrowserCredential authentication failed: User canceled authentication.'
        }
    }

    BeforeEach {
        Mock Write-NCMessage {}
        Mock Get-Module { [pscustomobject]@{ Name = 'Microsoft.Graph' } } -ParameterFilter { $ListAvailable }
        Mock Import-Module {}
        Mock Get-MgContext { $null }
    }

    It 'reports a short canceled message without the raw Connect-MgGraph error' {
        $output = @(Test-MgGraphConnection -EnsureExchangeOnline:$false 2>&1)

        $output | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | Should -BeNullOrEmpty
        $output | Should -Be @($false)
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -match 'sign-in was canceled' }
    }

    It 'keeps the real error when something else was canceled' {
        function Connect-MgGraph {
            [CmdletBinding()]
            param([string[]]$Scopes, [switch]$NoWelcome, [string]$TenantId, [switch]$UseDeviceCode, [string]$LoginHint)
            Write-Error 'The operation was canceled.'
        }

        Test-MgGraphConnection -EnsureExchangeOnline:$false | Should -BeFalse

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -match 'The operation was canceled' }
    }
}
