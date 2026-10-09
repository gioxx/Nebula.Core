BeforeAll {
    function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
    function Add-EmptyLine {}
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Set-ProgressAndInfoPreferences {}
    function Restore-ProgressAndInfoPreferences {}
    function Find-UserRecipient { param([string]$UserPrincipalName, [switch]$PreferGraphIdentity, [switch]$SkipDirectGraphLookup) }
    function Invoke-MgGraphRequest {
        param(
            [string]$Uri,
            [string]$Method,
            [object]$Body,
            [string]$ContentType,
            [string]$OutputType
        )
    }
    function Get-MgContext {}
    function Get-NCProgressPercent { param($Current, $Total) if ($Total) { [int](100 * $Current / $Total) } else { 0 } }

    function New-TestBatchResponse {
        param(
            [object]$Body,
            [scriptblock]$Responder
        )
        $payload = $Body | ConvertFrom-Json
        @{
            responses = @(foreach ($request in $payload.requests) {
                    $answer = & $Responder $request
                    @{ id = $request.id; status = $answer.status; headers = $answer.headers; body = $answer.body }
                })
        }
    }

    . "$PSScriptRoot/../../Private/NC-Hlp.Intune.ps1"
    . "$PSScriptRoot/../../Private/NC-Hlp.GraphBatch.ps1"
    . "$PSScriptRoot/../../Public/NC.Users.ps1"
}

Describe 'Remove-EntraUser batching' {
    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Write-NCMessage {}
        Mock Add-EmptyLine {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Find-UserRecipient {}
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        $global:DeleteFailsFor = ''
    }

    BeforeAll {
        function New-Upns { param([int]$Count) 1..$Count | ForEach-Object { "user$_@contoso.com" } }

        function Set-UsersGraphMock {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($request.method -eq 'GET' -and $url -match '^/users/ghost%40contoso\.com') {
                        return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'not found' } } }
                    }
                    if ($request.method -eq 'GET' -and $url -match '^/users/user(\d+)%40contoso\.com\?\$select=id,displayName,userPrincipalName,mail,userType$') {
                        $n = $Matches[1]
                        return @{ status = 200; body = @{ id = "id$n"; displayName = "User $n"; userPrincipalName = "user$n@contoso.com"; mail = "user$n@contoso.com"; userType = 'Member' } }
                    }
                    if ($request.method -eq 'DELETE' -and $url -match '^/users/(id\d+)$') {
                        if ($Matches[1] -eq $global:DeleteFailsFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 204 }
                    }
                    @{ status = 500 }
                }
            }
        }
    }

    It 'uses 2 Graph calls for 14 users' {
        Set-UsersGraphMock
        New-Upns 14 | Remove-EntraUser -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
        @($global:SeenRequests | Where-Object { $_.method -eq 'DELETE' }).Count | Should -Be 14
        Should -Invoke Test-MgGraphConnection -Times 1 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 14 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' -and $Message -eq "Removed Entra user 'User 1' (user1@contoso.com)." }
        Should -Invoke Find-UserRecipient -Times 0 -Exactly -Scope It
    }

    It 'flushes a batch every 20 users while streaming from the pipeline' {
        Set-UsersGraphMock
        New-Upns 25 | Remove-EntraUser -Confirm:$false

        Should -Invoke Invoke-MgGraphRequest -Times 4 -Exactly -Scope It
        Should -Invoke Write-NCMessage -Times 25 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' }
    }

    It 'only reads with -WhatIf' {
        Set-UsersGraphMock
        New-Upns 3 | Remove-EntraUser -WhatIf

        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
        @($global:SeenRequests | Where-Object { $_.method -eq 'DELETE' }).Count | Should -Be 0
    }

    It 'emits the removed user details in input order with -PassThru' {
        Set-UsersGraphMock
        $result = @(New-Upns 3 | Remove-EntraUser -Confirm:$false -PassThru)

        $result.Count | Should -Be 3
        $result[0].'Display Name' | Should -Be 'User 1'
        $result[0].'User Principal Name' | Should -Be 'user1@contoso.com'
        $result[0].Mail | Should -Be 'user1@contoso.com'
        $result[0].'User Type' | Should -Be 'Member'
        $result[0].'User Id' | Should -Be 'id1'
        $result[2].'User Id' | Should -Be 'id3'
    }

    It 'reports an unresolved user with the existing message and does not try the Exchange fallback' {
        Set-UsersGraphMock
        @('ghost@contoso.com', 'user1@contoso.com') | Remove-EntraUser -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq "Unable to resolve user 'ghost@contoso.com': not found" }
        Should -Invoke Find-UserRecipient -Times 0 -Exactly -Scope It
        @($global:SeenRequests | Where-Object { $_.method -eq 'DELETE' }).Count | Should -Be 1
    }

    It 'reports a failed delete with the existing message' {
        $global:DeleteFailsFor = 'id2'
        Set-UsersGraphMock
        @('user1@contoso.com', 'user2@contoso.com') | Remove-EntraUser -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq "Unable to remove user 'User 2': denied" }
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' }
    }

    It 'warns on an empty identifier' {
        Set-UsersGraphMock
        '  ' | Remove-EntraUser -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -eq 'UserPrincipalName cannot be empty.' }
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -Scope It
    }

    It 'reports the connection failure once and does nothing' {
        Mock Test-MgGraphConnection { $false }
        Set-UsersGraphMock
        New-Upns 3 | Remove-EntraUser -Confirm:$false

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' }
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -Scope It
    }
}
