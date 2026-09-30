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
    function Get-Recipient {
        param($Identity, $ErrorAction)
    }
    function Get-MgUser {
        param($UserId, $Property, $Filter, [switch]$All, $ErrorAction)
    }
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
    . "$PSScriptRoot/../../Private/NC-Hlp.UserUtils.ps1"
}

Describe 'Find-UserRecipient batched filter fallback' {
    BeforeEach {
        Mock Start-Sleep {}
        Mock Write-Progress {}
        Mock Write-NCMessage {}
        Mock Get-Recipient { throw 'recipient not found' }
        Mock Get-MgUser { throw 'direct lookup failed' }
        $script:sentUrls = [System.Collections.Generic.List[string]]::new()
    }

    It 'sends one batch with 4 sub-requests for an alias and picks the first query with results' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                $script:sentUrls.Add([string]$r.url)
                if ($r.id -eq '2') {
                    @{ status = 200; body = @{ value = @(@{ id = 'id-sam'; userPrincipalName = 'sam@contoso.com'; mail = 'sam.mail@contoso.com'; displayName = 'Sam' }) } }
                }
                elseif ($r.id -eq '3') {
                    @{ status = 200; body = @{ value = @(@{ id = 'id-dn'; userPrincipalName = 'dn@contoso.com'; mail = $null; displayName = 'Dn' }) } }
                }
                else {
                    @{ status = 200; body = @{ value = @() } }
                }
            }
        }

        $result = Find-UserRecipient -UserPrincipalName "o'brien"

        $result | Should -Be 'sam.mail@contoso.com'
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        $script:sentUrls.Count | Should -Be 4
        $script:sentUrls[0] | Should -BeLike '*mailNickname%20eq%20%27o%27%27brien%27*'
        $script:sentUrls[1] | Should -BeLike '*onPremisesSamAccountName*'
        $script:sentUrls[2] | Should -BeLike '*displayName*'
        $script:sentUrls[3] | Should -BeLike '*startswith*'
    }

    It 'sends 2 sub-requests for an input containing @' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                $script:sentUrls.Add([string]$r.url)
                if ($r.id -eq '2') {
                    @{ status = 200; body = @{ value = @(@{ id = 'id-1'; userPrincipalName = 'a@contoso.com'; mail = $null; displayName = 'A' }) } }
                }
                else {
                    @{ status = 200; body = @{ value = @() } }
                }
            }
        }

        $result = Find-UserRecipient -UserPrincipalName 'alias@contoso.com'

        $result | Should -Be 'a@contoso.com'
        $script:sentUrls.Count | Should -Be 2
        $script:sentUrls[0] | Should -BeLike '*userPrincipalName%20eq*'
        $script:sentUrls[1] | Should -BeLike '*mail%20eq*'
    }

    It 'returns the object id with -PreferGraphIdentity' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                if ($r.id -eq '1') {
                    @{ status = 200; body = @{ value = @(@{ id = 'id-0'; userPrincipalName = 'x@contoso.com'; mail = 'x@contoso.com'; displayName = 'X' }) } }
                }
                else { @{ status = 200; body = @{ value = @() } } }
            }
        }

        Find-UserRecipient -UserPrincipalName 'x' -PreferGraphIdentity | Should -Be 'id-0'
    }

    It 'treats a failed sub-request as no results and keeps the multiple-match warning' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                if ($r.id -eq '1') {
                    @{ status = 400; body = @{ error = @{ code = 'Request_UnsupportedQuery'; message = 'unsupported' } } }
                }
                elseif ($r.id -eq '3') {
                    @{ status = 200; body = @{ value = @(
                                @{ id = 'b'; userPrincipalName = 'b@contoso.com'; mail = 'b@contoso.com'; displayName = 'B' },
                                @{ id = 'a'; userPrincipalName = 'a@contoso.com'; mail = 'a@contoso.com'; displayName = 'A' }
                            ) }
                    }
                }
                else { @{ status = 200; body = @{ value = @() } } }
            }
        }

        $result = Find-UserRecipient -UserPrincipalName 'dup'

        $result | Should -Be 'a@contoso.com'
        Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
            $Message -eq "Multiple users matched 'dup'. Using the first result (a@contoso.com)." -and $Level -eq 'WARNING'
        }
    }

    It 'writes the not-found error when every query fails or is empty' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                if ($r.id -eq '1') { @{ status = 200; body = @{ value = @() } } }
                else { @{ status = 400; body = @{ error = @{ code = 'Bad'; message = 'bad' } } } }
            }
        }

        $result = Find-UserRecipient -UserPrincipalName 'ghost'

        $result | Should -BeNullOrEmpty
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
            $Message -like 'Recipient not available or not found (ghost).*' -and $Level -eq 'ERROR'
        }
    }
}
