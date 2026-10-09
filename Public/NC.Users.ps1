#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: User helpers =========================================================================================================================

function Search-EntraUser {
    <#
    .SYNOPSIS
        Finds Entra users by display name, user principal name, or mail.
    .DESCRIPTION
        Uses Microsoft Graph search to find users by display name, user principal name, or mail.
        This is especially useful for guest users whose Entra UPN contains fragments of the
        invited domain or tenant-specific suffixes.
    .PARAMETER SearchText
        Text to search for in display name, user principal name, and/or mail. Accepts pipeline input.
    .PARAMETER SearchIn
        Where to search: DisplayName, UserPrincipalName, Mail, or Any.
    .PARAMETER IndexOnly
        Use Microsoft Graph indexed search only. This is faster, but it can miss guest identities
        and other partial matches that only show up in a broader scan.
    .PARAMETER Detailed
        Show extra diagnostic columns such as match source, guest fragment, and user metadata.
    .PARAMETER GridView
        Show the result set in Out-GridView instead of returning objects.
    .EXAMPLE
        Search-EntraUser -SearchText "step"
    .EXAMPLE
        Search-EntraUser -SearchText "federica" -SearchIn DisplayName
    .EXAMPLE
        "guest" | Search-EntraUser -SearchIn Any -GridView
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Search', 'Query', 'Text', 'Name', 'DisplayName', 'UserPrincipalName', 'Mail')]
        [string]$SearchText,

        [ValidateSet('DisplayName', 'UserPrincipalName', 'Mail', 'Any')]
        [string]$SearchIn = 'Any',

        [switch]$IndexOnly,

        [switch]$Detailed,

        [switch]$GridView
    )

    process {
        $resolvedUsers = @(Resolve-EntraUserSearchResults -SearchText $SearchText -SearchIn $SearchIn -IndexOnly:$IndexOnly)

        Add-EmptyLine
        Write-Verbose "Users found: $($resolvedUsers.Count) for '$SearchText'."

        if (-not $resolvedUsers -or $resolvedUsers.Count -eq 0) {
            Write-NCMessage "No users found for '$SearchText'." -Level WARNING
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($candidate in $resolvedUsers) {
            $user = $candidate.User
            $row = [ordered]@{
                'Display Name'        = $user.DisplayName
                'User Principal Name' = $user.UserPrincipalName
                'Mail'                = $user.Mail
                'User Type'           = $user.UserType
            }

            if ($Detailed.IsPresent) {
                $row['Matched By'] = if ($candidate.MatchedBy.Count -gt 0) { $candidate.MatchedBy -join ', ' } else { $null }
                $row['Guest Fragment'] = $candidate.GuestFragment
                $row['User Id'] = $user.Id
                $row['Account Enabled'] = $user.AccountEnabled
                $row['Department'] = $user.Department
                $row['Job Title'] = $user.JobTitle
                $row['Search Mode'] = $candidate.SearchMode
            }

            $results.Add([pscustomobject]$row) | Out-Null
        }

        if ($GridView.IsPresent) {
            $results | Out-NCGridView -Title "Entra Users - Search: $SearchText"
        }
        else {
            $results | Sort-Object 'Display Name', 'User Principal Name'
        }
    }
}

function Remove-EntraUser {
    <#
    .SYNOPSIS
        Removes an Entra user from the tenant.
    .DESCRIPTION
        Removes the user identified by user principal name (UPN) with Microsoft Graph.
    .PARAMETER UserPrincipalName
        The user principal name of the Entra user to remove.
    .PARAMETER PassThru
        Return the removed user details after deletion.
    .EXAMPLE
        Remove-EntraUser -UserPrincipalName "antonio.sala_stepsrl.org#EXT#@messita.onmicrosoft.com"
    .EXAMPLE
        "federica.arpaia_stepsrl.it#EXT#@messita.onmicrosoft.com" | Remove-EntraUser -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('UPN', 'Identity')]
        [string]$UserPrincipalName,

        [switch]$PassThru
    )

    begin {
        $graphConnected = $null
        $queue = [System.Collections.Generic.List[string]]::new()
        $state = @{ StartLineWritten = $false }

        $flush = {
            if ($queue.Count -eq 0) { return }
            $items = @($queue)
            $queue.Clear()

            if (-not $state.StartLineWritten) {
                Write-NCGraphBatchNotice -Count ($items.Count) -Noun 'users' -Streaming
                $state.StartLineWritten = $true
            }

            for ($offset = 0; $offset -lt $items.Count; $offset += 20) {
                $chunk = @($items[$offset..([Math]::Min($offset + 20, $items.Count) - 1)])

                # (a) Read the whole chunk.
                $lookups = @(for ($i = 0; $i -lt $chunk.Count; $i++) {
                        @{
                            Id     = "m$i"
                            Method = 'GET'
                            Url    = "/users/$([uri]::EscapeDataString($chunk[$i]))?`$select=id,displayName,userPrincipalName,mail,userType"
                        }
                    })
                $lookupResponses = @(Invoke-NCGraphBatch -Requests $lookups -Activity 'Resolving users')

                # (b) Confirm per user.
                $approved = [System.Collections.Generic.List[object]]::new()
                for ($i = 0; $i -lt $chunk.Count; $i++) {
                    $response = $lookupResponses[$i]
                    if (-not $response.Success) {
                        Write-NCMessage "Unable to resolve user '$($chunk[$i])': $($response.ErrorMessage)" -Level ERROR
                        continue
                    }

                    $user = [pscustomobject]$response.Body
                    if ($PSCmdlet.ShouldProcess($user.userPrincipalName, "Remove Entra user $($user.displayName)")) {
                        $approved.Add($user) | Out-Null
                    }
                }

                # (c) Delete the approved users in one batch.
                if ($approved.Count -eq 0) { continue }
                $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                        @{
                            Id     = "m$i"
                            Method = 'DELETE'
                            Url    = "/users/$([uri]::EscapeDataString([string]$approved[$i].id))"
                        }
                    })
                $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity 'Removing users')

                # (d) Report in input order.
                for ($i = 0; $i -lt $approved.Count; $i++) {
                    $user = $approved[$i]
                    $response = $responses[$i]
                    if ($response.Success) {
                        if ($PassThru.IsPresent) {
                            [pscustomobject]@{
                                'Display Name'        = $user.displayName
                                'User Principal Name' = $user.userPrincipalName
                                'Mail'                = $user.mail
                                'User Type'           = $user.userType
                                'User Id'             = $user.id
                            }
                        }

                        Write-NCMessage "Removed Entra user '$($user.displayName)' ($($user.userPrincipalName))." -Level SUCCESS
                    }
                    else {
                        Write-NCMessage "Unable to remove user '$($user.displayName)': $($response.ErrorMessage)" -Level ERROR
                    }
                }
            }
        }
    }

    process {
        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            Write-NCMessage "UserPrincipalName cannot be empty." -Level WARNING
            return
        }

        if ($null -eq $graphConnected) {
            $graphConnected = Test-MgGraphConnection -Scopes @('User.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
            if (-not $graphConnected) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            }
        }
        if (-not $graphConnected) { return }

        $queue.Add($UserPrincipalName.Trim()) | Out-Null
        if ($queue.Count -ge 20) { & $flush }
    }

    end {
        if ($graphConnected) { & $flush }
    }
}
