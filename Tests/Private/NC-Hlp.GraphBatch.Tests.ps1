BeforeAll {
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
    function Get-MgEnvironment {
        param([string]$Name)
    }
    function Write-NCMessage {
        param(
            [string]$Message,
            [string]$Level
        )
    }
    function Find-UserRecipient {
        param(
            [string]$UserPrincipalName,
            [switch]$PreferGraphIdentity,
            [switch]$SkipDirectGraphLookup
        )
    }

    . "$PSScriptRoot/../../Private/NC-Hlp.Intune.ps1"
    . "$PSScriptRoot/../../Private/NC-Hlp.GraphBatch.ps1"
}

Describe 'Invoke-NCGraphBatch' {
    BeforeEach {
        Mock Start-Sleep {}
        Mock Write-Progress {}
    }

    It 'sends nothing for an empty request list' {
        Mock Invoke-MgGraphRequest {}

        $result = @(Invoke-NCGraphBatch -Requests @())

        $result.Count | Should -Be 0
        Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
    }

    It 'splits requests into chunks of 20 and returns results in input order' {
        $script:chunkSizes = [System.Collections.Generic.List[int]]::new()
        Mock Invoke-MgGraphRequest {
            $payload = $Body | ConvertFrom-Json
            $script:chunkSizes.Add(@($payload.requests).Count)
            $responses = @(foreach ($r in $payload.requests) { @{ id = $r.id; status = 200; body = @{ url = $r.url } } })
            [array]::Reverse($responses)
            @{ responses = $responses }
        }
        $requests = @(1..45 | ForEach-Object { @{ Id = "r$_"; Method = 'GET'; Url = "/users/u$_" } })

        $result = @(Invoke-NCGraphBatch -Requests $requests)

        @($script:chunkSizes) | Should -Be @(20, 20, 5)
        $result.Count | Should -Be 45
        $result[0].Id | Should -Be 'r1'
        $result[0].Body.url | Should -Be '/users/u1'
        $result[44].Body.url | Should -Be '/users/u45'
        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -eq 'v1.0/$batch' }
    }

    It 'writes ASCII progress on its own progress id and completes only that bar' {
        Mock Invoke-MgGraphRequest {
            $payload = $Body | ConvertFrom-Json
            @{ responses = @(foreach ($r in $payload.requests) { @{ id = $r.id; status = 200; body = @{} } }) }
        }
        $requests = @(1..25 | ForEach-Object { @{ Id = "r$_"; Method = 'GET'; Url = "/users/u$_" } })

        $null = Invoke-NCGraphBatch -Requests $requests -Activity 'Test activity'

        Should -Invoke Write-Progress -Times 1 -Exactly -ParameterFilter { $Id -eq 7781 -and $Status -eq 'Batch 1 - 0 items processed' }
        Should -Invoke Write-Progress -Times 1 -Exactly -ParameterFilter { $Id -eq 7781 -and $Status -eq 'Batch 2 - 20 items processed' }
        Should -Invoke Write-Progress -Times 1 -Exactly -ParameterFilter { $Id -eq 7781 -and $Completed }
        Should -Invoke Write-Progress -Times 0 -Exactly -ParameterFilter { $Id -ne 7781 }
    }

    It 'reports mixed per-request outcomes independently' {
        Mock Invoke-MgGraphRequest {
            @{
                responses = @(
                    @{ id = '1'; status = 204; body = $null }
                    @{ id = '2'; status = 400; body = @{ error = @{ code = 'Request_BadRequest'; message = 'One or more added object references already exist for the following modified properties: ''members''.' } } }
                    @{ id = '3'; status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'Resource ''x'' does not exist.' } } }
                )
            }
        }
        $requests = @(
            @{ Id = 'a'; Method = 'POST'; Url = '/groups/g/members/$ref'; Body = @{ '@odata.id' = 'x' } }
            @{ Id = 'b'; Method = 'POST'; Url = '/groups/g/members/$ref'; Body = @{ '@odata.id' = 'y' } }
            @{ Id = 'c'; Method = 'POST'; Url = '/groups/g/members/$ref'; Body = @{ '@odata.id' = 'z' } }
        )

        $result = @(Invoke-NCGraphBatch -Requests $requests)

        $result[0].Success | Should -BeTrue
        $result[0].Status | Should -Be 204
        $result[1].Success | Should -BeFalse
        $result[1].ErrorMessage | Should -Match 'added object references already exist'
        $result[2].Status | Should -Be 404
        $result[2].ErrorCode | Should -Be 'Request_ResourceNotFound'
    }

    It 'adds a JSON content type to sub-requests that carry a body' {
        $script:sent = $null
        Mock Invoke-MgGraphRequest {
            $script:sent = $Body | ConvertFrom-Json
            @{ responses = @(@{ id = '1'; status = 204 }) }
        }

        $null = Invoke-NCGraphBatch -Requests @(@{ Id = 'a'; Method = 'post'; Url = '/groups/g/members/$ref'; Body = @{ '@odata.id' = 'https://graph.microsoft.com/v1.0/directoryObjects/u1' } })

        $script:sent.requests[0].method | Should -Be 'POST'
        $script:sent.requests[0].headers.'Content-Type' | Should -Be 'application/json'
        $script:sent.requests[0].body.'@odata.id' | Should -Be 'https://graph.microsoft.com/v1.0/directoryObjects/u1'
    }

    It 'retries only throttled sub-requests after the Retry-After delay' {
        $script:calls = 0
        $script:secondPayload = $null
        Mock Invoke-MgGraphRequest {
            $script:calls++
            $payload = $Body | ConvertFrom-Json
            if ($script:calls -eq 1) {
                return @{
                    responses = @(
                        @{ id = '1'; status = 200; body = @{ id = 'u1' } }
                        @{ id = '2'; status = 429; headers = @{ 'Retry-After' = '7' } }
                    )
                }
            }
            $script:secondPayload = $payload
            @{ responses = @(@{ id = $payload.requests[0].id; status = 200; body = @{ id = 'u2' } }) }
        }
        $requests = @(
            @{ Id = 'a'; Method = 'GET'; Url = '/users/u1' }
            @{ Id = 'b'; Method = 'GET'; Url = '/users/u2' }
        )

        $result = @(Invoke-NCGraphBatch -Requests $requests)

        $script:calls | Should -Be 2
        @($script:secondPayload.requests).Count | Should -Be 1
        $script:secondPayload.requests[0].url | Should -Be '/users/u2'
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 7 }
        $result[1].Body.id | Should -Be 'u2'
    }

    It 'returns the throttled status once retries are exhausted' {
        Mock Invoke-MgGraphRequest {
            $payload = $Body | ConvertFrom-Json
            @{ responses = @(@{ id = $payload.requests[0].id; status = 429; headers = @{ 'Retry-After' = '1' }; body = @{ error = @{ code = 'TooManyRequests'; message = 'Too many requests.' } } }) }
        }

        $result = @(Invoke-NCGraphBatch -Requests @(@{ Id = 'a'; Method = 'GET'; Url = '/users/u1' }) -MaxRetries 2)

        Should -Invoke Invoke-MgGraphRequest -Times 3 -Exactly
        $result[0].Status | Should -Be 429
        $result[0].Success | Should -BeFalse
    }

    It 'retries the whole batch on a transient failure' {
        $script:calls = 0
        Mock Invoke-MgGraphRequest {
            $script:calls++
            if ($script:calls -eq 1) { throw 'Response status code does not indicate success: 503 (Service Unavailable).' }
            @{ responses = @(@{ id = '1'; status = 200; body = @{ id = 'u1' } }) }
        }

        $result = @(Invoke-NCGraphBatch -Requests @(@{ Id = 'a'; Method = 'GET'; Url = '/users/u1' }))

        $script:calls | Should -Be 2
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 5 }
        $result[0].Success | Should -BeTrue
    }

    It 'marks a chunk as not attempted on a permanent failure and still sends later chunks' {
        $script:calls = 0
        Mock Invoke-MgGraphRequest {
            $script:calls++
            if ($script:calls -eq 1) { throw 'Response status code does not indicate success: 403 (Forbidden).' }
            $payload = $Body | ConvertFrom-Json
            @{ responses = @(foreach ($r in $payload.requests) { @{ id = $r.id; status = 200; body = @{} } }) }
        }
        $requests = @(1..25 | ForEach-Object { @{ Id = "r$_"; Method = 'GET'; Url = "/users/u$_" } })

        $result = @(Invoke-NCGraphBatch -Requests $requests)

        $result.Count | Should -Be 25
        $result[0].Status | Should -Be 0
        $result[0].ErrorCode | Should -Be 'BatchRequestFailed'
        $result[0].ErrorMessage | Should -Match '^Graph batch request failed, operation not attempted: .*403'
        $result[19].ErrorCode | Should -Be 'BatchRequestFailed'
        $result[20].Success | Should -BeTrue
        Should -Invoke Start-Sleep -Times 0 -Exactly
    }

    It 'rejects duplicate request ids' {
        Mock Invoke-MgGraphRequest {}
        $requests = @(
            @{ Id = 'a'; Method = 'GET'; Url = '/users/u1' }
            @{ Id = 'a'; Method = 'GET'; Url = '/users/u2' }
        )

        { Invoke-NCGraphBatch -Requests $requests } | Should -Throw '*Duplicate request Id*'
    }
}

Describe 'Invoke-NCGraphBatchCollection' {
    BeforeEach {
        Mock Start-Sleep {}
        Mock Write-Progress {}
    }

    It 'returns the first page and follows nextLink for results that have more pages' {
        Mock Invoke-MgGraphRequest {
            @{
                responses = @(
                    @{ id = '1'; status = 200; body = @{ value = @(@{ id = 'g1' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/u1/memberOf?$skiptoken=abc' } }
                    @{ id = '2'; status = 200; body = @{ value = @() } }
                    @{ id = '3'; status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } }
                )
            }
        }
        Mock Invoke-NCGraphAllPagesCore { @([pscustomobject]@{ id = 'g2' }) }
        $requests = @(
            @{ Id = 'a'; Method = 'GET'; Url = '/users/u1/memberOf' }
            @{ Id = 'b'; Method = 'GET'; Url = '/users/u2/memberOf' }
            @{ Id = 'c'; Method = 'GET'; Url = '/users/u3/memberOf' }
        )

        $result = @(Invoke-NCGraphBatchCollection -Requests $requests)

        @($result[0].Items).Count | Should -Be 2
        $result[0].Items[1].id | Should -Be 'g2'
        @($result[1].Items).Count | Should -Be 0
        $result[1].Success | Should -BeTrue
        $result[2].Success | Should -BeFalse
        @($result[2].Items).Count | Should -Be 0
        Should -Invoke Invoke-NCGraphAllPagesCore -Times 1 -Exactly -ParameterFilter { $Uri -like '*skiptoken=abc' }
    }
}

Describe 'Resolve-NCGraphUserBatch' {
    BeforeEach {
        Mock Start-Sleep {}
        Mock Write-Progress {}
        Mock Write-NCMessage {}
    }

    It 'resolves direct identifiers in one batch without the fallback' {
        Mock Invoke-MgGraphRequest {
            $payload = $Body | ConvertFrom-Json
            @{ responses = @(foreach ($r in $payload.requests) { @{ id = $r.id; status = 200; body = @{ id = "id-$($r.id)"; userPrincipalName = ($r.url -replace '^/users/([^?]+)\?.*$', '$1') } } }) }
        }
        Mock Find-UserRecipient {}

        $map = Resolve-NCGraphUserBatch -Identifier @('alice@contoso.com', ' bob@contoso.com ', 'ALICE@contoso.com')

        $map.Count | Should -Be 2
        $map['alice@contoso.com'].userPrincipalName | Should -Be 'alice%40contoso.com'
        $map['bob@contoso.com'] | Should -Not -BeNullOrEmpty
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Should -Invoke Find-UserRecipient -Times 0 -Exactly
    }

    It 'requests the selected properties with an encoded identifier' {
        $script:urls = @()
        Mock Invoke-MgGraphRequest {
            $payload = $Body | ConvertFrom-Json
            $script:urls = @($payload.requests.url)
            @{ responses = @(foreach ($r in $payload.requests) { @{ id = $r.id; status = 200; body = @{ id = 'x' } } }) }
        }

        $null = Resolve-NCGraphUserBatch -Identifier @('guest_contoso.com#EXT#@tenant.onmicrosoft.com') -Property @('id', 'usageLocation')

        $script:urls[0] | Should -Be '/users/guest_contoso.com%23EXT%23%40tenant.onmicrosoft.com?$select=id,usageLocation'
    }

    It 'falls back to Find-UserRecipient only for identifiers Graph could not find' {
        $script:calls = 0
        Mock Invoke-MgGraphRequest {
            $script:calls++
            $payload = $Body | ConvertFrom-Json
            if ($script:calls -eq 1) {
                return @{
                    responses = @(
                        @{ id = '1'; status = 200; body = @{ id = 'id-alice'; userPrincipalName = 'alice@contoso.com' } }
                        @{ id = '2'; status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } }
                        @{ id = '3'; status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } }
                    )
                }
            }
            @{ responses = @(@{ id = $payload.requests[0].id; status = 200; body = @{ id = 'id-guest'; userPrincipalName = 'guest_ext.com#EXT#@contoso.onmicrosoft.com' } }) }
        }
        Mock Find-UserRecipient {
            if ($UserPrincipalName -eq 'guest@ext.com') { return 'id-guest' }
        }

        $map = Resolve-NCGraphUserBatch -Identifier @('alice@contoso.com', 'guest@ext.com', 'ghost@contoso.com')

        $map['alice@contoso.com'].id | Should -Be 'id-alice'
        $map['guest@ext.com'].id | Should -Be 'id-guest'
        $map['ghost@contoso.com'] | Should -BeNullOrEmpty
        $map.Contains('ghost@contoso.com') | Should -BeTrue
        Should -Invoke Find-UserRecipient -Times 2 -Exactly
        Should -Invoke Find-UserRecipient -Times 0 -Exactly -ParameterFilter { $UserPrincipalName -eq 'alice@contoso.com' }
        Should -Invoke Find-UserRecipient -Times 2 -Exactly -ParameterFilter { $PreferGraphIdentity -and $SkipDirectGraphLookup }
        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly
    }

    It 'reports non-404 lookup failures with the existing resolve message' {
        Mock Invoke-MgGraphRequest {
            @{ responses = @(@{ id = '1'; status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'Insufficient privileges.' } } }) }
        }
        Mock Find-UserRecipient {}

        $map = Resolve-NCGraphUserBatch -Identifier @('alice@contoso.com')

        $map['alice@contoso.com'] | Should -BeNullOrEmpty
        Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter { $Message -eq "Unable to resolve user 'alice@contoso.com': Insufficient privileges." -and $Level -eq 'ERROR' }
        Should -Invoke Find-UserRecipient -Times 0 -Exactly
    }

    It 'collects non-not-found failures in FailedIdentifier and leaves not-found identifiers out' {
        $script:calls = 0
        Mock Invoke-MgGraphRequest {
            $script:calls++
            if ($script:calls -eq 1) {
                return @{
                    responses = @(
                        @{ id = '1'; status = 403; body = @{ error = @{ code = 'Authorization_RequestDenied'; message = 'Insufficient privileges.' } } }
                        @{ id = '2'; status = 404; body = @{ error = @{ code = 'Request_ResourceNotFound'; message = 'missing' } } }
                    )
                }
            }
            throw 'unexpected second batch'
        }
        Mock Find-UserRecipient {}

        $failed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $null = Resolve-NCGraphUserBatch -Identifier @('denied@contoso.com', 'ghost@contoso.com') -FailedIdentifier $failed

        $failed.Contains('DENIED@contoso.com') | Should -BeTrue
        $failed.Contains('ghost@contoso.com') | Should -BeFalse
        $failed.Count | Should -Be 1
    }
}

Describe 'Get-NCGraphDirectoryObjectUri' {
    It 'uses the global endpoint by default' {
        Mock Get-MgContext { [pscustomobject]@{ Environment = 'Global' } }

        Get-NCGraphDirectoryObjectUri -Id 'u1' | Should -Be 'https://graph.microsoft.com/v1.0/directoryObjects/u1'
    }

    It 'uses the endpoint of a national cloud environment' {
        Mock Get-MgContext { [pscustomobject]@{ Environment = 'USGov' } }
        Mock Get-MgEnvironment { [pscustomobject]@{ GraphEndpoint = 'https://graph.microsoft.us/' } }

        Get-NCGraphDirectoryObjectUri -Id 'u1' | Should -Be 'https://graph.microsoft.us/v1.0/directoryObjects/u1'
    }
}

Describe 'Graph endpoint usage' {
    It 'never hard-codes the global Graph endpoint outside Get-NCGraphDirectoryObjectUri' {
        # Request URIs must be relative so Invoke-MgGraphRequest resolves them against the active cloud
        $sources = Get-ChildItem "$PSScriptRoot/../../Public", "$PSScriptRoot/../../Private" -Filter '*.ps1'
        $hits = @($sources | Select-String -SimpleMatch 'https://graph.microsoft.com' | Where-Object {
                -not ($_.Filename -eq 'NC-Hlp.GraphBatch.ps1' -and $_.Line -match "^\s*(\`$endpoint = 'https://graph\.microsoft\.com'|falling back to https://graph\.microsoft\.com\.)")
            } | ForEach-Object { "$($_.Filename):$($_.LineNumber)" })

        $hits | Should -BeNullOrEmpty
    }
}

Describe 'Write-NCGraphBatchNotice' {
    BeforeEach {
        Mock Write-NCMessage {}
    }

    It 'stays silent when the total fits one batch request' {
        Write-NCGraphBatchNotice -Count 1 -Noun 'user(s)'
        Write-NCGraphBatchNotice -Count 20 -Noun 'user(s)'

        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It
    }

    It 'writes the notice with the count when the total spans more than one request' {
        Write-NCGraphBatchNotice -Count 21 -Noun 'user(s)'

        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing 21 user(s) in Graph batches (20 per request) ...' -and $Level -eq 'INFO' }
    }

    It 'writes the notice without a count when the first streamed chunk is full' {
        Write-NCGraphBatchNotice -Count 19 -Noun 'users' -Streaming
        Should -Invoke Write-NCMessage -Times 0 -Exactly -Scope It

        Write-NCGraphBatchNotice -Count 20 -Noun 'users' -Streaming
        Should -Invoke Write-NCMessage -Times 1 -Exactly -Scope It -ParameterFilter { $Message -eq 'Processing users in Graph batches (20 per request) ...' -and $Level -eq 'INFO' }
    }

    It 'returns whether the notice was written only with PassThru' {
        Write-NCGraphBatchNotice -Count 5 -Noun 'app(s)' | Should -BeNullOrEmpty
        Write-NCGraphBatchNotice -Count 5 -Noun 'app(s)' -PassThru | Should -BeFalse
        Write-NCGraphBatchNotice -Count 30 -Noun 'app(s)' -PassThru | Should -BeTrue
    }
}
