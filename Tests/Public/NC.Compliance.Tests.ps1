BeforeAll {
    function Test-EOLConnection { $true }
    function Add-EmptyLine {}
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Set-ProgressAndInfoPreferences {}
    function Restore-ProgressAndInfoPreferences {}
    function Get-Mailbox { param($Identity) }
    function Get-RetentionPolicyTag { param($Identity) }
    function New-RetentionPolicyTag { param($Name, $Type, $RetentionEnabled, $AgeLimitForRetention, $RetentionAction) }
    function Get-RetentionPolicy { param($Identity) }
    function New-RetentionPolicy { param($Name, $RetentionPolicyTagLinks) }
    function Set-RetentionPolicy { param($Identity, $RetentionPolicyTagLinks) }
    function Set-Mailbox { param($Identity, $RetentionPolicy) }
    function Start-ManagedFolderAssistant { param($Identity) }
    # Simulates a display time zone behind the host: a time-zone-aware formatter moves the date back one day
    function Format-NCDateTime { param($Value, $Format, [switch]$AsLocalTime) ([datetime]$Value).AddDays(-1).ToString($Format) }

    . "$PSScriptRoot/../../Public/NC.Compliance.ps1"
}

Describe 'Set-MboxMrmCleanup' {
    BeforeEach {
        Mock Write-NCMessage {}
        Mock Get-Mailbox { throw 'not needed' }
    }

    It 'prints the date-only cutoff as given, without time-zone conversion' {
        Set-MboxMrmCleanup -Mailbox 'user1@contoso.com' -FixedCutoffDate ([datetime]'2025-01-01') -WhatIf

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Fixed cutoff date: 01/01/2025' }
    }
}
