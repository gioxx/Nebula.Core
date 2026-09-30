# Nebula.Core: Microsoft Graph $batch helpers =======================================================================
# Groups per-item Graph calls into $batch requests (20 per HTTP call) so bulk operations issue far fewer
# requests; each HTTP request can trigger an interactive WAM broker round-trip in the Graph SDK.

function Get-NCGraphDirectoryObjectUri {
    <#
    .SYNOPSIS
        Builds the absolute directoryObjects URI used in @odata.id reference bodies.
    .DESCRIPTION
        Uses the Graph endpoint of the current Microsoft Graph environment (national clouds included),
        falling back to https://graph.microsoft.com.
    .PARAMETER Id
        Directory object ID.
    .PARAMETER ApiVersion
        Graph API version segment.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0'
    )

    $endpoint = 'https://graph.microsoft.com'
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        if ($ctx -and $ctx.Environment -and $ctx.Environment -ne 'Global') {
            $environment = Get-MgEnvironment -Name $ctx.Environment -ErrorAction Stop
            if ($environment -and $environment.GraphEndpoint) {
                $endpoint = ([string]$environment.GraphEndpoint).TrimEnd('/')
            }
        }
    }
    catch {
        Write-Verbose "Unable to read the Microsoft Graph environment, using the global endpoint: $($_.Exception.Message)"
    }

    return "$endpoint/$ApiVersion/directoryObjects/$Id"
}

function Test-NCGraphTransientError {
    <#
    .SYNOPSIS
        Tells whether a Graph request failure message describes a transient condition worth retrying.
    .PARAMETER Message
        Exception message to inspect.
    #>
    [CmdletBinding()]
    param(
        [string]$Message
    )

    return [bool]($Message -match '(?i)\b(429|500|502|503|504)\b|TooManyRequests|ServiceUnavailable|GatewayTimeout|BadGateway|timed out|timeout')
}

function Get-NCGraphRetryDelay {
    <#
    .SYNOPSIS
        Reads the Retry-After header of a batch sub-response, in seconds (default 5, capped at 60).
    .PARAMETER Headers
        Sub-response headers (dictionary or object).
    #>
    [CmdletBinding()]
    param(
        [object]$Headers
    )

    $value = $null
    if ($Headers -is [System.Collections.IDictionary]) {
        foreach ($key in $Headers.Keys) {
            if ([string]$key -ieq 'Retry-After') {
                $value = $Headers[$key]
                break
            }
        }
    }
    elseif ($Headers) {
        $value = $Headers.'Retry-After'
    }

    $seconds = 0
    if (-not [int]::TryParse([string]$value, [ref]$seconds) -or $seconds -le 0) {
        $seconds = 5
    }

    return [Math]::Min($seconds, 60)
}

function ConvertTo-NCGraphBatchResult {
    <#
    .SYNOPSIS
        Converts one $batch sub-response into the normalized Nebula result object.
    .PARAMETER Request
        Original request item (provides Id).
    .PARAMETER Response
        Sub-response from the $batch payload.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Request,
        [Parameter(Mandatory = $true)]
        [object]$Response
    )

    $status = [int]$Response.status
    $body = $Response.body
    $success = ($status -ge 200 -and $status -le 299)
    $errorCode = $null
    $errorMessage = $null

    if (-not $success) {
        $errorInfo = $null
        if ($body -is [System.Collections.IDictionary]) {
            if ($body.Contains('error')) { $errorInfo = $body['error'] }
        }
        elseif ($body) {
            $errorInfo = $body.error
        }

        if ($errorInfo) {
            $errorCode = [string]$errorInfo.code
            $errorMessage = [string]$errorInfo.message
        }
        if ([string]::IsNullOrWhiteSpace($errorMessage)) {
            $errorMessage = "Graph returned HTTP $status."
        }
    }

    return [pscustomobject]@{
        Id           = $Request.Id
        Status       = $status
        Success      = $success
        Body         = $body
        ErrorCode    = $errorCode
        ErrorMessage = $errorMessage
    }
}

function New-NCGraphBatchFailure {
    <#
    .SYNOPSIS
        Builds a failed result object for a request that Graph never answered.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Request,
        [Parameter(Mandatory = $true)]
        [string]$ErrorCode,
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    return [pscustomobject]@{
        Id           = $Request.Id
        Status       = 0
        Success      = $false
        Body         = $null
        ErrorCode    = $ErrorCode
        ErrorMessage = $ErrorMessage
    }
}

function Invoke-NCGraphBatch {
    <#
    .SYNOPSIS
        Sends Microsoft Graph requests through $batch, 20 per HTTP call.
    .DESCRIPTION
        Sub-requests are independent: a failing one never affects the others. Throttled sub-requests
        (429/503/504) are retried after Retry-After; transient failures of the whole batch call are retried
        with backoff. A permanent failure of the whole call marks that chunk as not attempted and the next
        chunks are still sent. Results are returned in input order and the function does not throw for
        request failures.
    .PARAMETER Requests
        Items shaped as @{ Id; Method; Url; Body; Headers }. Url is relative to the API version
        (for example /users/alice@contoso.com). Id must be unique.
    .PARAMETER ApiVersion
        Graph API version for the $batch endpoint (all requests in the call share it).
    .PARAMETER Activity
        Write-Progress activity label.
    .PARAMETER MaxRetries
        Maximum retry rounds for throttled sub-requests and transient batch failures.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Requests,
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0',
        [string]$Activity = 'Microsoft Graph batch',
        [ValidateRange(0, 10)]
        [int]$MaxRetries = 5
    )

    if ($Requests.Count -eq 0) {
        return
    }

    $seenIds = @{}
    foreach ($request in $Requests) {
        $key = [string]$request.Id
        if ($seenIds.ContainsKey($key)) {
            throw [System.ArgumentException]::new("Duplicate request Id '$key' in Invoke-NCGraphBatch.")
        }
        $seenIds[$key] = $true
    }

    $chunkSize = 20
    $results = @{}
    $processed = 0
    $batchNumber = 0

    for ($offset = 0; $offset -lt $Requests.Count; $offset += $chunkSize) {
        $batchNumber++
        $last = [Math]::Min($offset + $chunkSize, $Requests.Count) - 1
        $pending = @($Requests[$offset..$last])
        $attempt = 0

        while ($pending.Count -gt 0) {
            Write-Progress -Activity $Activity -Status "Batch $batchNumber · $processed items processed"

            $byBatchId = @{}
            $subRequests = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $pending.Count; $i++) {
                $request = $pending[$i]
                $batchId = [string]($i + 1)
                $byBatchId[$batchId] = $request

                $sub = [ordered]@{
                    id     = $batchId
                    method = ([string]$request.Method).ToUpperInvariant()
                    url    = [string]$request.Url
                }
                $headers = @{}
                if ($request.Headers) {
                    foreach ($headerName in $request.Headers.Keys) {
                        $headers[$headerName] = $request.Headers[$headerName]
                    }
                }
                if ($null -ne $request.Body) {
                    $sub.body = $request.Body
                    if (-not $headers.ContainsKey('Content-Type')) {
                        $headers['Content-Type'] = 'application/json'
                    }
                }
                if ($headers.Count -gt 0) {
                    $sub.headers = $headers
                }
                $subRequests.Add($sub)
            }

            $payload = @{ requests = @($subRequests) } | ConvertTo-Json -Depth 20 -Compress

            try {
                $response = Invoke-MgGraphRequest -Method POST -Uri "$ApiVersion/`$batch" -Body $payload -ContentType 'application/json' -OutputType HashTable -ErrorAction Stop
            }
            catch {
                $reason = $_.Exception.Message
                if ((Test-NCGraphTransientError -Message $reason) -and $attempt -lt $MaxRetries) {
                    $attempt++
                    $wait = [int][Math]::Min(60, 5 * [Math]::Pow(2, $attempt - 1))
                    Write-Progress -Activity $Activity -Status "Throttled by Graph, retrying in ${wait}s ..."
                    Start-Sleep -Seconds $wait
                    continue
                }

                foreach ($request in $pending) {
                    $results[[string]$request.Id] = New-NCGraphBatchFailure -Request $request -ErrorCode 'BatchRequestFailed' -ErrorMessage "Graph batch request failed, operation not attempted: $reason"
                }
                break
            }

            $retry = [System.Collections.Generic.List[object]]::new()
            $wait = 0
            $answered = @{}
            foreach ($sub in @($response.responses)) {
                if ($null -eq $sub) { continue }
                $batchId = [string]$sub.id
                if (-not $byBatchId.ContainsKey($batchId)) { continue }

                $request = $byBatchId[$batchId]
                $answered[$batchId] = $true
                $status = [int]$sub.status

                if (($status -in 429, 503, 504) -and $attempt -lt $MaxRetries) {
                    $retry.Add($request)
                    $wait = [Math]::Max($wait, (Get-NCGraphRetryDelay -Headers $sub.headers))
                    continue
                }

                $results[[string]$request.Id] = ConvertTo-NCGraphBatchResult -Request $request -Response $sub
            }

            foreach ($batchId in $byBatchId.Keys) {
                if (-not $answered.ContainsKey($batchId)) {
                    $request = $byBatchId[$batchId]
                    $results[[string]$request.Id] = New-NCGraphBatchFailure -Request $request -ErrorCode 'MissingBatchResponse' -ErrorMessage 'Graph batch returned no response for this request.'
                }
            }

            $pending = @($retry)
            if ($pending.Count -gt 0) {
                $attempt++
                Write-Progress -Activity $Activity -Status "Throttled by Graph, retrying in ${wait}s ..."
                Start-Sleep -Seconds $wait
            }
        }

        $processed += ($last - $offset + 1)
    }

    Write-Progress -Activity $Activity -Completed

    foreach ($request in $Requests) {
        $results[[string]$request.Id]
    }
}
