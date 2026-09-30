BeforeAll {
    function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
    function Add-EmptyLine {}
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Set-ProgressAndInfoPreferences {}
    function Restore-ProgressAndInfoPreferences {}
    function Find-UserRecipient { param([string]$UserPrincipalName, [switch]$PreferGraphIdentity, [switch]$SkipDirectGraphLookup) }
    function Invoke-NCRetry {
        param([scriptblock]$Action, [int]$MaxAttempts, [int]$DelaySeconds, [string]$OperationDescription, [scriptblock]$OnError)
        & $Action
    }
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
    function Get-MgUser { [CmdletBinding()] param($Filter, $ConsistencyLevel, [switch]$All, $Property) }

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
    . "$PSScriptRoot/../../Public/NC.Security.ps1"
}

Describe 'Security batching' {
    BeforeEach {
        Mock Test-MgGraphConnection { $true }
        Mock Write-NCMessage {}
        Mock Add-EmptyLine {}
        Mock Set-ProgressAndInfoPreferences {}
        Mock Restore-ProgressAndInfoPreferences {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        Mock Find-UserRecipient {}
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
        $global:PatchFailsFor = ''
        $global:NoDevicesFor = ''
        $global:DevicesFailFor = ''
    }

    BeforeAll {
        function New-Upns { param([int]$Count) 1..$Count | ForEach-Object { "user$_@contoso.com" } }

        function Set-SecurityGraphMock {
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($request.method -eq 'GET' -and $url -match '^/users/ghost%40contoso\.com') {
                        return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'not found' } } }
                    }
                    if ($request.method -eq 'GET' -and $url -match '^/users/user(\d+)%40contoso\.com') {
                        $n = $Matches[1]
                        return @{ status = 200; body = @{ id = "id$n"; userPrincipalName = "user$n@contoso.com"; displayName = "User $n" } }
                    }
                    if ($request.method -eq 'GET' -and $url -match '^/users/(id(\d+))/registeredDevices') {
                        if ($Matches[1] -eq $global:DevicesFailFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'nope' } } } }
                        if ($Matches[1] -eq $global:NoDevicesFor) { return @{ status = 200; body = @{ value = @() } } }
                        $n = $Matches[2]
                        return @{ status = 200; body = @{ value = @(
                                    @{ id = "dev${n}a"; displayName = "Laptop $n" },
                                    @{ id = "dev${n}b"; displayName = $null }
                                ) } }
                    }
                    if ($request.method -eq 'PATCH' -and $url -match '^/users/(id\d+)$') {
                        if ($Matches[1] -eq $global:PatchFailsFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 204 }
                    }
                    if ($request.method -eq 'PATCH' -and $url -match '^/devices/(dev\w+)$') {
                        if ($Matches[1] -eq $global:PatchFailsFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 204 }
                    }
                    if ($request.method -eq 'POST' -and $url -match '^/users/(id\d+)/revokeSignInSessions$') {
                        if ($Matches[1] -eq $global:PatchFailsFor) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        return @{ status = 200; body = @{ value = $true } }
                    }
                    @{ status = 500 }
                }
            }
        }
    }

    Context 'Disable-UserSignIn' {
        It 'splits 25 users into 2 write batches (20 + 5) with the right ids' {
            Set-SecurityGraphMock
            $result = New-Upns 25 | Disable-UserSignIn -Confirm:$false -PassThru

            # 2 chunks x (resolve batch + write batch)
            Should -Invoke Invoke-MgGraphRequest -Times 4 -Exactly -Scope It
            $patches = @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' })
            $patches.Count | Should -Be 25
            # Sub-request ids restart at 1 in every $batch call: one '1' per write batch.
            @($patches | Where-Object { $_.id -eq '1' }).Count | Should -Be 2
            @($patches | Where-Object { $_.id -eq '20' }).Count | Should -Be 1
            @($patches | ForEach-Object { $_.url }) | Should -Be @(1..25 | ForEach-Object { "/users/id$_" })
            @($result).Count | Should -Be 25
        }

        It 'PATCHes only the users approved at the -Confirm prompt when the middle one is declined' {
            if (-not ('NCScriptedHost' -as [type])) {
                Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Security;

public class NCScriptedHostUI : PSHostUserInterface {
    public Queue<int> Answers = new Queue<int>();
    public List<string> Prompts = new List<string>();
    public override PSHostRawUserInterface RawUI { get { return null; } }
    public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice) {
        Prompts.Add(message);
        return Answers.Count > 0 ? Answers.Dequeue() : defaultChoice;
    }
    public override Dictionary<string, PSObject> Prompt(string caption, string message, Collection<FieldDescription> descriptions) { throw new NotSupportedException(); }
    public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName) { throw new NotSupportedException(); }
    public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options) { throw new NotSupportedException(); }
    public override string ReadLine() { return string.Empty; }
    public override SecureString ReadLineAsSecureString() { return new SecureString(); }
    public override void Write(string value) { }
    public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) { }
    public override void WriteLine(string value) { }
    public override void WriteErrorLine(string value) { }
    public override void WriteDebugLine(string message) { }
    public override void WriteProgress(long sourceId, ProgressRecord record) { }
    public override void WriteVerboseLine(string message) { }
    public override void WriteWarningLine(string message) { }
}

public class NCScriptedHost : PSHost {
    private readonly NCScriptedHostUI ui = new NCScriptedHostUI();
    private readonly Guid instanceId = Guid.NewGuid();
    public NCScriptedHostUI ScriptedUI { get { return ui; } }
    public override string Name { get { return "NCScriptedHost"; } }
    public override Version Version { get { return new Version(1, 0); } }
    public override Guid InstanceId { get { return instanceId; } }
    public override PSHostUserInterface UI { get { return ui; } }
    public override CultureInfo CurrentCulture { get { return CultureInfo.InvariantCulture; } }
    public override CultureInfo CurrentUICulture { get { return CultureInfo.InvariantCulture; } }
    public override void SetShouldExit(int exitCode) { }
    public override void EnterNestedPrompt() { throw new NotSupportedException(); }
    public override void ExitNestedPrompt() { }
    public override void NotifyBeginApplication() { }
    public override void NotifyEndApplication() { }
}
'@
            }

            # $PSCmdlet.ShouldProcess cannot be mocked: run the real function in a runspace whose host answers the
            # -Confirm prompts (Yes, No, Yes), with plain stubs in place of the Pester mocks.
            $scriptedHost = [NCScriptedHost]::new()
            $scriptedHost.ScriptedUI.Answers.Enqueue(0)
            $scriptedHost.ScriptedUI.Answers.Enqueue(2)
            $scriptedHost.ScriptedUI.Answers.Enqueue(0)
            $runspace = [runspacefactory]::CreateRunspace($scriptedHost)
            $runspace.Open()
            try {
                $ps = [powershell]::Create()
                $ps.Runspace = $runspace
                $null = $ps.AddScript({
                        param([string]$Root)
                        function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
                        function Add-EmptyLine {}
                        function Write-NCMessage { param([string]$Message, [string]$Level) }
                        function Set-ProgressAndInfoPreferences {}
                        function Restore-ProgressAndInfoPreferences {}
                        function Find-UserRecipient { param([string]$UserPrincipalName, [switch]$PreferGraphIdentity, [switch]$SkipDirectGraphLookup) }
                        function Get-MgContext {}
                        function Get-NCProgressPercent { param($Current, $Total) 0 }
                        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
                        function Invoke-MgGraphRequest {
                            param([string]$Uri, [string]$Method, [object]$Body, [string]$ContentType, [string]$OutputType)
                            $payload = $Body | ConvertFrom-Json
                            @{
                                responses = @(foreach ($request in $payload.requests) {
                                        $global:SeenRequests.Add($request)
                                        if ($request.method -eq 'GET' -and ([string]$request.url) -match '^/users/user(\d+)%40contoso\.com') {
                                            $n = $Matches[1]
                                            @{ id = $request.id; status = 200; body = @{ id = "id$n"; userPrincipalName = "user$n@contoso.com"; displayName = "User $n" } }
                                        }
                                        else {
                                            @{ id = $request.id; status = 204 }
                                        }
                                    })
                            }
                        }
                        . "$Root/Private/NC-Hlp.Intune.ps1"
                        . "$Root/Private/NC-Hlp.GraphBatch.ps1"
                        . "$Root/Public/NC.Security.ps1"

                        $result = @('user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com') | Disable-UserSignIn -Confirm -PassThru
                        [pscustomobject]@{
                            Result  = $result
                            Patches = @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' } | ForEach-Object { [string]$_.url })
                        }
                    }).AddArgument((Resolve-Path "$PSScriptRoot/../..").Path)
                $output = $ps.Invoke()
                $ps.Streams.Error | ForEach-Object { throw $_ }
            }
            finally {
                $runspace.Dispose()
            }

            $scriptedHost.ScriptedUI.Prompts.Count | Should -Be 3
            @($output[0].Patches) | Should -Be @('/users/id1', '/users/id3')
            @($output[0].Result.UserPrincipalName) | Should -Be @('user1@contoso.com', 'user3@contoso.com')
        }

        It 'uses 2 Graph calls for 14 users and PATCHes accountEnabled=false' {
            Set-SecurityGraphMock
            New-Upns 14 | Disable-UserSignIn -Confirm:$false

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            $patches = @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' })
            $patches.Count | Should -Be 14
            $patches[0].url | Should -Be '/users/id1'
            $patches[0].body.accountEnabled | Should -BeFalse
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' -and $Message -eq 'Sign-in disabled for 14 users.' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'INFO' -and $Message -like 'Processing 14 user(s) in Graph batches*' }
        }

        It 'only reads with -WhatIf' {
            Set-SecurityGraphMock
            New-Upns 3 | Disable-UserSignIn -WhatIf

            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -Scope It
            @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' }).Count | Should -Be 0
        }

        It 'reports a missing user once and keeps going' {
            Set-SecurityGraphMock
            $result = @('user1@contoso.com', 'ghost@contoso.com', 'user2@contoso.com') | Disable-UserSignIn -Confirm:$false -PassThru

            @($result).Count | Should -Be 2
            $result[0].UserPrincipalName | Should -Be 'user1@contoso.com'
            $result[0].DisplayName | Should -Be 'User 1'
            $result[0].Action | Should -Be 'SignInDisabled'
            $result[1].UserPrincipalName | Should -Be 'user2@contoso.com'
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq "Can't find Azure AD account for user ghost@contoso.com." }
        }

        It 'reports a failed PATCH with the existing message and omits the user from output' {
            $global:PatchFailsFor = 'id2'
            Set-SecurityGraphMock
            $result = @('user1@contoso.com', 'user2@contoso.com') | Disable-UserSignIn -Confirm:$false -PassThru

            @($result).Count | Should -Be 1
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq 'Failed to disable sign-in for user2@contoso.com. denied' }
        }
    }

    Context 'Revoke-UserSessions' {
        It 'uses 2 Graph calls for 14 users and POSTs revokeSignInSessions with an empty JSON body' {
            Set-SecurityGraphMock
            New-Upns 14 | Revoke-UserSessions -Confirm:$false

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            $posts = @($global:SeenRequests | Where-Object { $_.method -eq 'POST' })
            $posts.Count | Should -Be 14
            $posts[0].url | Should -Be '/users/id1/revokeSignInSessions'
            foreach ($post in $posts) {
                $post.PSObject.Properties['body'] | Should -Not -BeNullOrEmpty
                (ConvertTo-Json -InputObject $post.body -Compress) | Should -Be '{}'
                $post.headers.'Content-Type' | Should -Be 'application/json'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' -and $Message -eq 'Revoked sessions for 14 users.' }
        }

        It 'only reads with -WhatIf' {
            Set-SecurityGraphMock
            New-Upns 3 | Revoke-UserSessions -WhatIf

            @($global:SeenRequests | Where-Object { $_.method -eq 'POST' }).Count | Should -Be 0
        }

        It 'skips excluded users and reports failures' {
            $global:PatchFailsFor = 'id3'
            Set-SecurityGraphMock
            $result = New-Upns 3 | Revoke-UserSessions -Exclude 'user1@contoso.com' -Confirm:$false -PassThru

            @($result).Count | Should -Be 1
            $result.UserPrincipalName | Should -Be 'user2@contoso.com'
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'INFO' -and $Message -eq 'Skipping user user1@contoso.com' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq 'Failed to revoke sessions for user3@contoso.com. denied' }
        }

        It 'reports a missing user once' {
            Set-SecurityGraphMock
            'ghost@contoso.com', 'user1@contoso.com' | Revoke-UserSessions -Confirm:$false

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq "Can't find Azure AD account for user ghost@contoso.com." }
        }

        It 'batches the revoke loop for -All after a single user read' {
            Mock Get-MgUser { 1..25 | ForEach-Object { [pscustomobject]@{ Id = "id$_"; UserPrincipalName = "user$_@contoso.com"; DisplayName = "User $_" } } }
            Set-SecurityGraphMock
            Revoke-UserSessions -All -Confirm:$false

            Should -Invoke Get-MgUser -Times 1 -Exactly -Scope It
            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            @($global:SeenRequests | Where-Object { $_.method -eq 'POST' }).Count | Should -Be 25
        }
    }

    Context 'Disable-UserDevices' {
        It 'uses 3 Graph calls for 3 users with 2 devices each' {
            Set-SecurityGraphMock
            $result = New-Upns 3 | Disable-UserDevices -Confirm:$false -PassThru

            Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -Scope It
            $patches = @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' })
            $patches.Count | Should -Be 6
            $patches[0].body.accountEnabled | Should -BeFalse
            @($result).Count | Should -Be 6
            $result[0].UserPrincipalName | Should -Be 'user1@contoso.com'
            $result[0].UserDisplayName | Should -Be 'User 1'
            $result[0].DeviceId | Should -Be 'dev1a'
            $result[0].DeviceDisplayName | Should -Be 'Laptop 1'
            $result[0].Action | Should -Be 'Disabled'
            $result[5].DeviceId | Should -Be 'dev3b'
        }

        It 'prints the summary without -PassThru' {
            Set-SecurityGraphMock
            New-Upns 3 | Disable-UserDevices -Confirm:$false

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'SUCCESS' -and $Message -eq 'Disabled 6 devices.' }
        }

        It 'only reads with -WhatIf' {
            Set-SecurityGraphMock
            New-Upns 3 | Disable-UserDevices -WhatIf

            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly -Scope It
            @($global:SeenRequests | Where-Object { $_.method -eq 'PATCH' }).Count | Should -Be 0
        }

        It 'warns when a user has no devices and errors when device read fails' {
            $global:NoDevicesFor = 'id1'
            $global:DevicesFailFor = 'id2'
            Set-SecurityGraphMock
            $result = New-Upns 3 | Disable-UserDevices -Confirm:$false -PassThru

            @($result).Count | Should -Be 2
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARNING' -and $Message -eq 'No registered devices found for user1@contoso.com.' }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq 'Unable to retrieve registered devices for user2@contoso.com. nope' }
        }

        It 'reports a failed device update using the device label' {
            $global:PatchFailsFor = 'dev1b'
            Set-SecurityGraphMock
            $result = New-Upns 1 | Disable-UserDevices -Confirm:$false -PassThru

            @($result).Count | Should -Be 1
            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq 'Failed to disable device dev1b for user1@contoso.com. denied' }
        }

        It 'reports a missing user once' {
            Set-SecurityGraphMock
            'ghost@contoso.com' | Disable-UserDevices -Confirm:$false

            Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $Message -eq "Can't find Azure AD account for user ghost@contoso.com." }
        }
    }
}
