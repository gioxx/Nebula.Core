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

    It 'skips only the direct Get-MgUser lookup with -SkipDirectGraphLookup' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                @{ status = 200; body = @{ value = @() } }
            }
        }

        $result = Find-UserRecipient -UserPrincipalName 'ghost@contoso.com' -PreferGraphIdentity -SkipDirectGraphLookup

        $result | Should -BeNullOrEmpty
        Should -Invoke Get-Recipient -Times 1 -Exactly
        Should -Invoke Get-MgUser -Times 0 -Exactly
        Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly
        Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
            $Message -eq 'Recipient not available or not found (ghost@contoso.com).' -and $Level -eq 'ERROR'
        }
    }

    It 'still resolves through the filter queries with -SkipDirectGraphLookup' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                if ($r.id -eq '2') { @{ status = 200; body = @{ value = @(@{ id = 'id-mail'; userPrincipalName = 'u@contoso.com'; mail = 'alias@contoso.com'; displayName = 'U' }) } } }
                else { @{ status = 200; body = @{ value = @() } } }
            }
        }

        Find-UserRecipient -UserPrincipalName 'alias@contoso.com' -PreferGraphIdentity -SkipDirectGraphLookup | Should -Be 'id-mail'
        Should -Invoke Get-MgUser -Times 0 -Exactly
    }

    It 'keeps the direct lookup error in the not-found message without the switch' {
        Mock Invoke-MgGraphRequest {
            New-TestBatchResponse -Body $Body -Responder {
                param($r)
                @{ status = 200; body = @{ value = @() } }
            }
        }

        $null = Find-UserRecipient -UserPrincipalName 'ghost@contoso.com'

        Should -Invoke Get-MgUser -Times 1 -Exactly
        Should -Invoke Write-NCMessage -Times 1 -Exactly -ParameterFilter {
            $Message -eq 'Recipient not available or not found (ghost@contoso.com). direct lookup failed' -and $Level -eq 'ERROR'
        }
    }
}

Describe 'Resolve-EntraUserSearchResults field selection' {
    BeforeAll {
        function Test-MgGraphConnection { param([string[]]$Scopes, [bool]$EnsureExchangeOnline) $true }
        function Add-EmptyLine {}
        function Get-MgUser {
            [CmdletBinding()]
            param($UserId, $Property, $Filter, $Search, $ConsistencyLevel, $CountVariable, [switch]$All)
        }
        $script:alice = [pscustomobject]@{ Id = 'id-alice'; DisplayName = 'Alice Rossi'; UserPrincipalName = 'alice@contoso.com'; Mail = 'alice@contoso.com' }
    }

    BeforeEach {
        Mock Write-NCMessage {}
        Mock Get-MgUser {
            if ($UserId -in 'alice@contoso.com', 'id-alice') { return $script:alice }
            if ($UserId) { throw 'not found' }
            @()
        }
    }

    It 'does not return a direct UPN match when searching only in DisplayName' {
        $result = @(Resolve-EntraUserSearchResults -SearchText 'alice@contoso.com' -SearchIn DisplayName -IndexOnly)

        $result.Count | Should -Be 0
        Should -Invoke Get-MgUser -Times 1 -Exactly -Scope It -ParameterFilter { $Search -eq '"displayName:alice@contoso.com"' }
    }

    It 'does not return a direct object ID match when searching only in Mail' {
        $result = @(Resolve-EntraUserSearchResults -SearchText 'id-alice' -SearchIn Mail -IndexOnly)

        $result.Count | Should -Be 0
        Should -Invoke Get-MgUser -Times 1 -Exactly -Scope It -ParameterFilter { $Search -eq '"mail:id-alice"' }
    }

    It 'skips the tenant-wide scan when an exact UPN or object ID resolves directly' {
        $any = @(Resolve-EntraUserSearchResults -SearchText 'alice@contoso.com' -SearchIn Any)
        $upn = @(Resolve-EntraUserSearchResults -SearchText 'alice@contoso.com' -SearchIn UserPrincipalName)

        $any.Count | Should -Be 1
        $upn.Count | Should -Be 1
        Should -Invoke Get-MgUser -Times 0 -Exactly -Scope It -ParameterFilter { $All -and $null -eq $Search }
    }

    It 'scans the tenant once when partial matching is needed' {
        Mock Get-MgUser {
            if ($UserId) { throw 'not found' }
            if ($All -and $null -eq $Search) { return @($script:alice) }
            @()
        }

        $result = @(Resolve-EntraUserSearchResults -SearchText 'rossi' -SearchIn DisplayName)

        $result.Count | Should -Be 1
        Should -Invoke Get-MgUser -Times 1 -Exactly -Scope It -ParameterFilter { $All -and $null -eq $Search }
    }

    It 'keeps a direct match whose selected field contains the search text' {
        $result = @(Resolve-EntraUserSearchResults -SearchText 'alice@contoso.com' -SearchIn UserPrincipalName -IndexOnly)

        $result.Count | Should -Be 1
        Should -Invoke Get-MgUser -Times 0 -Exactly -Scope It -ParameterFilter { $null -ne $Search }
    }
}
