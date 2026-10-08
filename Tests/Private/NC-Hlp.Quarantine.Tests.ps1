BeforeAll {
    function Get-QuarantineMessage {
        [CmdletBinding()]
        param(
            [string]$RecipientAddress,
            [string]$SenderAddress,
            [datetime]$StartReceivedDate,
            [datetime]$EndReceivedDate,
            [int]$PageSize,
            [int]$Page
        )
    }

    . "$PSScriptRoot/../../Private/NC-Hlp.Quarantine.ps1"
}

Describe 'Get-NCQuarantineMessageAllPages' {
    It 'reads every page until a partial page is returned' {
        Mock Get-QuarantineMessage {
            $count = switch ($Page) { 1 { 1000 } 2 { 1000 } 3 { 5 } default { 0 } }
            @(1..$count | Where-Object { $_ -gt 0 } | ForEach-Object { [pscustomobject]@{ Identity = "p$Page-$_" } })
        }

        $result = @(Get-NCQuarantineMessageAllPages -Parameters @{ RecipientAddress = 'alice@contoso.com' })

        $result.Count | Should -Be 2005
        Should -Invoke Get-QuarantineMessage -Times 3 -Exactly -Scope It
        Should -Invoke Get-QuarantineMessage -Times 3 -Exactly -Scope It -ParameterFilter {
            $RecipientAddress -eq 'alice@contoso.com' -and $PageSize -eq 1000
        }
    }

    It 'makes a single call when the first page is partial' {
        Mock Get-QuarantineMessage { @([pscustomobject]@{ Identity = 'only' }) }

        $result = @(Get-NCQuarantineMessageAllPages -Parameters @{ SenderAddress = 'bob@fabrikam.com' })

        $result.Count | Should -Be 1
        Should -Invoke Get-QuarantineMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Page -eq 1 -and $SenderAddress -eq 'bob@fabrikam.com' }
    }

    It 'lets Get-QuarantineMessage errors reach the caller' {
        Mock Get-QuarantineMessage { throw 'EXO unavailable' }

        { Get-NCQuarantineMessageAllPages -Parameters @{} } | Should -Throw '*EXO unavailable*'
    }
}
