BeforeAll {
    function Set-ProgressAndInfoPreferences {}
    function Restore-ProgressAndInfoPreferences {}
    function Test-EOLConnection {}
    function Add-EmptyLine {}
    function Write-NCMessage {
        param(
            [string]$Message,
            [string]$Level
        )
    }
    function Get-Mailbox {}
    function Get-MailboxPermission {}
    function Get-RecipientPermission {}
    function Get-User {}

    . "$PSScriptRoot/../../Public/NC.Mailboxes.ps1"
}

Describe 'Get-MboxPermission' {
    It 'shows the source mailbox RecipientTypeDetails in the heading' {
        Mock Set-ProgressAndInfoPreferences {}
        Mock Restore-ProgressAndInfoPreferences {}
        Mock Test-EOLConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Get-Mailbox {
            [pscustomobject]@{
                DisplayName          = 'Human Resources'
                PrimarySmtpAddress   = 'hr@contoso.com'
                RecipientTypeDetails = 'SharedMailbox'
                GrantSendOnBehalfTo  = @()
            }
        }
        Mock Get-MailboxPermission { @() }
        Mock Get-RecipientPermission { @() }
        Mock Get-User { $null }

        $null = Get-MboxPermission -SourceMailbox 'hr@contoso.com'

        Assert-MockCalled Write-NCMessage -Times 1 -ParameterFilter {
            $Message -eq 'Access Rights on Human Resources (hr@contoso.com) - SharedMailbox' -and
            $Level -eq 'WARNING'
        }
    }
}
