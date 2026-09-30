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
    function Get-Mailbox { param($Identity) }
    function Get-MailboxPermission {}
    function Get-RecipientPermission {}
    function Get-User {}
    function Get-ExoMailbox {}
    function Get-MailboxStatisticsSafe { param($Identity) }
    function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
    function Get-NCProgressPercent { param($Current, $Total) if ($Total) { [int](100 * $Current / $Total) } else { 0 } }
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


Describe 'Sign-in log batching' {
    BeforeEach {
        Mock Set-ProgressAndInfoPreferences {}
        Mock Restore-ProgressAndInfoPreferences {}
        Mock Test-EOLConnection { $true }
        Mock Test-MgGraphConnection { $true }
        Mock Add-EmptyLine {}
        Mock Write-NCMessage {}
        Mock Write-Progress {}
        Mock Start-Sleep {}
        $global:SeenRequests = [System.Collections.Generic.List[object]]::new()
    }

    BeforeAll {
        function New-TestMailboxes {
            param([int]$Count)
            1..$Count | ForEach-Object {
                [pscustomobject]@{
                    DisplayName               = "Shared $_"
                    PrimarySmtpAddress        = "shared$_@contoso.com"
                    ExternalDirectoryObjectId = "id$_"
                }
            }
        }
    }

    Context 'Get-UserLastSeen' {
        BeforeEach {
            Mock Get-Mailbox {
                $n = ([string]$Identity -replace '\D', '')
                [pscustomobject]@{
                    DisplayName               = "User $n"
                    PrimarySmtpAddress        = "user$n@contoso.com"
                    ExternalDirectoryObjectId = "id$n"
                }
            }
            Mock Get-MailboxStatisticsSafe { [pscustomobject]@{ LastUserActionTime = [datetime]'2026-01-10T08:00:00' } }
            $global:SignInDates = @('2026-02-01T10:00:00Z', '2026-01-20T10:00:00Z')
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    if ($request.url -match 'userId%20eq%20%27id(\d+)%27') {
                        $n = [int]$Matches[1]
                        if ($n -eq 3) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        if ($n -eq 4) { return @{ status = 200; body = @{ value = @() } } }
                        return @{ status = 200; body = @{ value = @(foreach ($d in $global:SignInDates) { @{ createdDateTime = $d } }) } }
                    }
                    @{ status = 500 }
                }
            }
        }

        It 'uses 1 Graph call for 14 mailboxes' {
            $result = @(1..14 | ForEach-Object { "user$_@contoso.com" } | Get-UserLastSeen)
            $result.Count | Should -Be 14
            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
            Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter { $Message -like 'Processing mailboxes in Graph batches*' }
        }

        It 'uses 2 Graph calls for 21 mailboxes and keeps input order' {
            $result = @(1..21 | ForEach-Object { "user$_@contoso.com" } | Get-UserLastSeen)
            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly
            ($result.PrimarySmtpAddress) | Should -Be @(1..21 | ForEach-Object { "user$_@contoso.com" })
        }

        It 'emits the same object shape with the latest sign-in as datetime' {
            $result = @('user1@contoso.com' | Get-UserLastSeen)
            $result[0].PSObject.Properties.Name | Should -Be @('DisplayName', 'PrimarySmtpAddress', 'LastUserActionTime', 'LastInteractiveSignIn', 'LastSeen', 'Source')
            $result[0].LastInteractiveSignIn | Should -BeOfType ([datetime])
            $result[0].LastInteractiveSignIn.Kind | Should -Be 'Utc'
            $result[0].LastInteractiveSignIn | Should -Be ([datetime]::SpecifyKind([datetime]'2026-02-01T10:00:00', 'Utc'))
            $result[0].LastSeen | Should -Be $result[0].LastInteractiveSignIn
            $result[0].Source | Should -Be 'MailboxAction,SignInLog'
        }

        It 'picks the latest sign-in even when the page is not newest-first' {
            $global:SignInDates = @('2026-02-01T10:00:00Z', '2026-03-05T12:30:00Z', '2026-01-20T10:00:00Z')
            $result = @('user1@contoso.com' | Get-UserLastSeen)
            $result[0].LastInteractiveSignIn.Kind | Should -Be 'Utc'
            $result[0].LastInteractiveSignIn | Should -Be ([datetime]::SpecifyKind([datetime]'2026-03-05T12:30:00', 'Utc'))
        }

        It 'normalizes datetime (Local) and datetimeoffset values to UTC' {
            $local = [datetime]::SpecifyKind([datetime]'2026-04-01T09:00:00', 'Local')
            $global:SignInDates = @($local, [datetimeoffset]::new(2026, 2, 1, 10, 0, 0, [timespan]::FromHours(2)))
            $result = @('user1@contoso.com' | Get-UserLastSeen)
            $result[0].LastInteractiveSignIn.Kind | Should -Be 'Utc'
            $result[0].LastInteractiveSignIn | Should -Be $local.ToUniversalTime()

            $global:SignInDates = @([datetimeoffset]::new(2026, 2, 1, 10, 0, 0, [timespan]::FromHours(2)))
            $result = @('user1@contoso.com' | Get-UserLastSeen)
            $result[0].LastInteractiveSignIn.Kind | Should -Be 'Utc'
            $result[0].LastInteractiveSignIn | Should -Be ([datetime]::SpecifyKind([datetime]'2026-02-01T08:00:00', 'Utc'))
        }

        It 'treats an unspecified-kind datetime as UTC' {
            $global:SignInDates = @([datetime]::SpecifyKind([datetime]'2026-02-01T10:00:00', 'Unspecified'))
            $result = @('user1@contoso.com' | Get-UserLastSeen)
            $result[0].LastInteractiveSignIn.Kind | Should -Be 'Utc'
            $result[0].LastInteractiveSignIn | Should -Be ([datetime]::SpecifyKind([datetime]'2026-02-01T10:00:00', 'Utc'))
        }
        It 'requests the encoded userId filter with top 20' {
            $null = 'user7@contoso.com' | Get-UserLastSeen
            $global:SeenRequests[0].url | Should -Be '/auditLogs/signIns?$filter=userId%20eq%20%27id7%27&$top=20'
        }

        It 'warns and keeps mailbox activity when the sign-in read fails, and handles empty logs' {
            $result = @(3, 4 | ForEach-Object { "user$_@contoso.com" } | Get-UserLastSeen)
            $result[0].Source | Should -Be 'MailboxAction'
            $result[1].Source | Should -Be 'MailboxAction'
            Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
                $Message -eq 'Unable to retrieve sign-in logs for user3@contoso.com. denied' -and $Level -eq 'WARNING'
            }
        }

        It 'skips Graph entirely when the connection is not available' {
            Mock Test-MgGraphConnection { $false }
            $result = @('user1@contoso.com', 'user2@contoso.com' | Get-UserLastSeen)
            $result.Count | Should -Be 2
            Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
            Should -Invoke Write-NCMessage -Times 2 -Exactly -ParameterFilter { $Message -like 'Microsoft Graph AuditLog.Read.All is not available*' }
        }
    }

    Context 'Test-SharedMailboxCompliance' {
        BeforeEach {
            Mock Get-ExoMailbox { New-TestMailboxes -Count 14 }
            Mock Invoke-MgGraphRequest {
                New-TestBatchResponse -Body $Body -Responder {
                    param($request)
                    $global:SeenRequests.Add($request)
                    $url = [string]$request.url
                    if ($url -match '^/auditLogs/signIns\?\$filter=userid%20eq%20%27id(\d+)%27&\$top=20$') {
                        $n = [int]$Matches[1]
                        if ($n -eq 5) { return @{ status = 403; body = @{ error = @{ code = 'Forbidden'; message = 'denied' } } } }
                        if ($n -in 1, 2) { return @{ status = 200; body = @{ value = @(@{ status = @{ errorCode = 50126 } }, @{ status = @{ errorCode = 0 } }) } } }
                        return @{ status = 200; body = @{ value = @(@{ status = @{ errorCode = 50126 } }) } }
                    }
                    if ($url -match '^/users/id(\d+)\?\$select=userPrincipalName,assignedPlans$') {
                        $n = [int]$Matches[1]
                        if ($n -eq 2) { return @{ status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'gone' } } } }
                        return @{ status = 200; body = @{ userPrincipalName = "shared$n@contoso.com"; assignedPlans = @(
                                    @{ service = 'exchange'; capabilityStatus = 'Enabled'; servicePlanId = '9aaf7827-d63c-4b61-89c3-182f06f82e5c' },
                                    @{ service = 'exchange'; capabilityStatus = 'Deleted'; servicePlanId = 'efb87545-963c-4e0d-99df-69c6916d9eb0' }
                                ) } }
                    }
                    @{ status = 500 }
                }
            }
        }

        It 'uses 2 Graph calls for 14 mailboxes and keeps the report unchanged' {
            $report = @(Test-SharedMailboxCompliance -GridView:$false)
            Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly
            $report.Count | Should -Be 14
            $report[0].PSObject.Properties.Name | Should -Be @('DisplayName', 'ExternalDirectoryObjectId', 'Sign in Record Found', 'Exchange Online Plan 1', 'Exchange Online Plan 2')
            $report[0].DisplayName | Should -Be 'Shared 1'
            $report[0].'Sign in Record Found' | Should -Be 'Yes'
            $report[0].'Exchange Online Plan 1' | Should -BeTrue
            $report[0].'Exchange Online Plan 2' | Should -BeFalse
            $report[2].'Sign in Record Found' | Should -Be 'No'
        }

        It 'reports read failures with the original texts' {
            $null = Test-SharedMailboxCompliance -GridView:$false
            Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
                $Message -eq 'Unable to retrieve sign-in records for Shared 5. denied' -and $Level -eq 'ERROR'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
                $Message -eq 'Unable to read license info for Shared 2. gone' -and $Level -eq 'ERROR'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
                $Message -eq 'Sign-in records found for shared mailbox Shared 1' -and $Level -eq 'WARNING'
            }
            Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter { $Message -like 'Processing 14 mailbox(es) in Graph batches*' }
        }
    }
}
