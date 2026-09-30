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
            [switch]$PreferGraphIdentity
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
