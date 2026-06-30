#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) User helpers ===============================================================================================================

function Find-UserConnected {
    <#
    .SYNOPSIS
        Retrieves the e-mail address of the currently logged-on domain user.
    .DESCRIPTION
        Uses ADSI to query Active Directory for the current user's 'mail' attribute.
        If not found, optionally prompts for manual entry.
    .NOTES
        Based on the original tip from PowerShell Magazine (2012).
    #>
    [CmdletBinding()]
    param()

    try {
        $sam = $env:USERNAME
        if (-not $sam) {
            throw "Environment variable USERNAME not set."
        }

        # Use ADSI to locate the current user
        $searcher = [adsisearcher]"(samaccountname=$sam)"
        $result = $searcher.FindOne()

        if ($result -and $result.Properties.mail) {
            return [string]($result.Properties.mail | Select-Object -First 1) # Return the first e-mail found
        }
        else {
            # Try UPN fallback (often same as e-mail)
            $upn = (& whoami /upn 2>$null).Trim()
            if ($upn -and $upn -match '@') {
                return $upn
            }
            else {
                Write-NCMessage "E-mail address not found for current user." -Level WARNING
                $manual = Read-Host "Please, specify your e-mail address"
                return $manual
            }
        }
    }
    catch {
        Write-NCMessage "Unable to automatically find e-mail address for current user." -Level WARNING
        $manual = Read-Host "Please, specify your e-mail address"
        return $manual
    }
}

function Find-UserRecipient {
    <#
    .SYNOPSIS
        Resolves and returns a user identity for Exchange or Microsoft Graph usage.
    .DESCRIPTION
        By default, this function preserves the historical behavior and returns the user's
        primary SMTP address when available.
        When -PreferGraphIdentity is specified, the function prefers a Graph-friendly identity
        such as the Entra object ID or the user principal name.
    .PARAMETER UserPrincipalName
        The UPN or identifier of the user recipient to resolve.
    .PARAMETER PreferGraphIdentity
        Returns a Graph-friendly identity instead of the primary SMTP address.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [switch]$PreferGraphIdentity
    )

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        return
    }

    try {
        $recipient = Get-Recipient -Identity $UserPrincipalName -ErrorAction Stop
    }
    catch {
        try {
            $user = Get-MgUser -UserId $UserPrincipalName -Property Id, UserPrincipalName, Mail -ErrorAction Stop

            if ($PreferGraphIdentity.IsPresent) {
                return $user.Id
            }

            if ($user.Mail) {
                return $user.Mail
            }

            return $user.UserPrincipalName
        }
        catch {
            $escaped = $UserPrincipalName.Replace("'", "''")
            $queries = @()

            if ($UserPrincipalName -match '@') {
                $queries += "userPrincipalName eq '$escaped'"
                $queries += "mail eq '$escaped'"
            }
            else {
                $queries += "mailNickname eq '$escaped'"
                $queries += "onPremisesSamAccountName eq '$escaped'"
                $queries += "displayName eq '$escaped'"
                $queries += "startswith(userPrincipalName,'$escaped@')"
            }

            $matchedUsers = @()
            foreach ($query in $queries) {
                try {
                    $matchedUsers = @(Get-MgUser -Filter $query -All -Property Id, UserPrincipalName, Mail, DisplayName -ErrorAction Stop)
                    if ($matchedUsers.Count -gt 0) {
                        break
                    }
                }
                catch {
                    continue
                }
            }

            if ($matchedUsers.Count -gt 0) {
                $selectedUser = $matchedUsers | Sort-Object UserPrincipalName | Select-Object -First 1

                if ($matchedUsers.Count -gt 1) {
                    $selectedLabel = if ($selectedUser.UserPrincipalName) { $selectedUser.UserPrincipalName } else { $selectedUser.DisplayName }
                    Write-NCMessage "Multiple users matched '$UserPrincipalName'. Using the first result ($selectedLabel)." -Level WARNING
                }

                if ($PreferGraphIdentity.IsPresent) {
                    return $selectedUser.Id
                }

                if ($selectedUser.Mail) {
                    return $selectedUser.Mail
                }

                return $selectedUser.UserPrincipalName
            }

            Write-NCMessage "Recipient not available or not found ($UserPrincipalName). $($_.Exception.Message)" -Level ERROR
        }

        return
    }

    if (-not $recipient) {
        Write-NCMessage "Recipient not available or not found ($UserPrincipalName)." -Level ERROR
        return
    }

    if ($PreferGraphIdentity.IsPresent) {
        if ($recipient.ExternalDirectoryObjectId) {
            return [string]$recipient.ExternalDirectoryObjectId
        }

        if ($recipient.WindowsLiveID -and $recipient.WindowsLiveID -match '@') {
            return [string]$recipient.WindowsLiveID
        }

        if ($recipient.PrimarySmtpAddress) {
            return [string]$recipient.PrimarySmtpAddress
        }

        $primaryGraphCandidate = @($recipient.EmailAddresses | Where-Object { $_ -clike 'SMTP:*' } | Select-Object -First 1)
        if ($primaryGraphCandidate.Count -gt 0) {
            return $primaryGraphCandidate[0].Substring(5)
        }

        return
    }

    $resolvedAddress = $recipient.PrimarySmtpAddress
    if (-not $resolvedAddress) {
        $resolvedAddress = $recipient.WindowsLiveID
    }

    if ($resolvedAddress -notmatch '@') {
        $primaryMatches = @($recipient.EmailAddresses | Where-Object { $_ -clike 'SMTP:*' })
        if ($primaryMatches.Count -eq 0) {
            Write-NCMessage "Complete e-mail address not specified and no primary SMTP address found for $resolvedAddress." -Level WARNING
            return
        }

        if ($primaryMatches.Count -gt 1) {
            Write-NCMessage "Multiple SMTP addresses detected for $resolvedAddress. Using the first match ($($primaryMatches[0].Substring(5)))." -Level WARNING
        }

        $resolvedAddress = $primaryMatches[0].Substring(5)
    }

    return $resolvedAddress
}

function Resolve-EntraUserSearchResults {
    <#
    .SYNOPSIS
        Resolves Entra users by partial or exact identity and returns match metadata.
    .DESCRIPTION
        Uses Microsoft Graph search first, then optionally falls back to a broader scan for
        guest identities and other partial matches.
    .PARAMETER SearchText
        Text used to find matching users.
    .PARAMETER SearchIn
        Where to search: DisplayName, UserPrincipalName, Mail, or Any.
    .PARAMETER IndexOnly
        Use Microsoft Graph indexed search only.
    .PARAMETER Scopes
        Microsoft Graph scopes to validate before searching.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchText,

        [ValidateSet('DisplayName', 'UserPrincipalName', 'Mail', 'Any')]
        [string]$SearchIn = 'Any',

        [switch]$IndexOnly,

        [string[]]$Scopes = @('User.Read.All', 'Directory.Read.All')
    )

    if ([string]::IsNullOrWhiteSpace($SearchText)) {
        return @()
    }

    $graphConnected = Test-MgGraphConnection -Scopes $Scopes -EnsureExchangeOnline:$false
    if (-not $graphConnected) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return @()
    }

    $escapedText = $SearchText.Replace('"', '""').Trim()
    if ([string]::IsNullOrWhiteSpace($escapedText)) {
        return @()
    }

    $selectProperties = @(
        'Id',
        'DisplayName',
        'UserPrincipalName',
        'Mail',
        'OtherMails',
        'ProxyAddresses',
        'UserType',
        'AccountEnabled',
        'Department',
        'JobTitle'
    )

    $searchNeedle = $escapedText.ToLowerInvariant()
    $allUsers = @()
    if (-not $IndexOnly.IsPresent) {
        try {
            $allUsers = @(Get-MgUser -All -Property $selectProperties -ErrorAction Stop)
        }
        catch {
            Write-NCMessage "Unable to load users for fallback matching: $($_.Exception.Message)" -Level ERROR
            return @()
        }
    }

    $users = @()

    try {
        switch ($SearchIn) {
            'DisplayName' {
                try {
                    $users = @(Get-MgUser -UserId $SearchText -Property $selectProperties -ErrorAction Stop)
                }
                catch {
                    $searchClause = "`"displayName:$escapedText`""
                    $users = @(Get-MgUser -Search $searchClause -ConsistencyLevel eventual -CountVariable count -All -Property $selectProperties -ErrorAction Stop)
                }

                if (-not $IndexOnly.IsPresent) {
                    $fallbackUsers = @($allUsers | Where-Object {
                        $_.DisplayName -and $_.DisplayName.ToLowerInvariant().Contains($searchNeedle)
                    })
                    $users = @($users + $fallbackUsers | Sort-Object Id -Unique)
                }
            }
            'UserPrincipalName' {
                try {
                    $users = @(Get-MgUser -UserId $SearchText -Property $selectProperties -ErrorAction Stop)
                }
                catch {
                    $searchClause = "`"userPrincipalName:$escapedText`""
                    $users = @(Get-MgUser -Search $searchClause -ConsistencyLevel eventual -CountVariable count -All -Property $selectProperties -ErrorAction Stop)
                }

                if (-not $IndexOnly.IsPresent) {
                    $fallbackUsers = @($allUsers | Where-Object {
                        $_.UserPrincipalName -and $_.UserPrincipalName.ToLowerInvariant().Contains($searchNeedle)
                    })
                    $users = @($users + $fallbackUsers | Sort-Object Id -Unique)
                }
            }
            'Mail' {
                try {
                    $users = @(Get-MgUser -UserId $SearchText -Property $selectProperties -ErrorAction Stop)
                }
                catch {
                    $searchClause = "`"mail:$escapedText`""
                    $users = @(Get-MgUser -Search $searchClause -ConsistencyLevel eventual -CountVariable count -All -Property $selectProperties -ErrorAction Stop)
                }

                if (-not $IndexOnly.IsPresent) {
                    $fallbackUsers = @($allUsers | Where-Object {
                        ($_.Mail -and $_.Mail.ToLowerInvariant().Contains($searchNeedle)) -or
                        ($_.OtherMails -and @($_.OtherMails | Where-Object { $_ -and $_.ToLowerInvariant().Contains($searchNeedle) }).Count -gt 0)
                    })
                    $users = @($users + $fallbackUsers | Sort-Object Id -Unique)
                }
            }
            'Any' {
                try {
                    $users = @(Get-MgUser -UserId $SearchText -Property $selectProperties -ErrorAction Stop)
                }
                catch {
                    $searchDisplay = "`"displayName:$escapedText`""
                    $searchUpn = "`"userPrincipalName:$escapedText`""
                    $searchMail = "`"mail:$escapedText`""

                    $byDisplay = @(Get-MgUser -Search $searchDisplay -ConsistencyLevel eventual -CountVariable countDisplay -All -Property $selectProperties -ErrorAction Stop)
                    $byUpn = @(Get-MgUser -Search $searchUpn -ConsistencyLevel eventual -CountVariable countUpn -All -Property $selectProperties -ErrorAction Stop)
                    $byMail = @(Get-MgUser -Search $searchMail -ConsistencyLevel eventual -CountVariable countMail -All -Property $selectProperties -ErrorAction Stop)

                    if ($IndexOnly.IsPresent) {
                        $users = @($byDisplay + $byUpn + $byMail | Sort-Object Id -Unique)
                    }
                    else {
                        $fallbackUsers = @($allUsers | Where-Object {
                            $candidates = @(
                                $_.DisplayName,
                                $_.UserPrincipalName,
                                $_.Mail
                            )

                            if ($_.OtherMails) {
                                $candidates += @($_.OtherMails)
                            }

                            if ($_.ProxyAddresses) {
                                $candidates += @($_.ProxyAddresses)
                            }

                            foreach ($candidate in $candidates) {
                                if ($candidate -and $candidate.ToLowerInvariant().Contains($searchNeedle)) {
                                    return $true
                                }
                            }

                            return $false
                        })

                        $users = @($byDisplay + $byUpn + $byMail + $fallbackUsers | Sort-Object Id -Unique)
                    }
                }
            }
        }
    }
    catch {
        Write-NCMessage "Unable to search users with '$SearchText': $($_.Exception.Message)" -Level ERROR
        return @()
    }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($user in $users) {
        $matchedBy = [System.Collections.Generic.List[string]]::new()

        if ($user.DisplayName -and $user.DisplayName.ToLowerInvariant().Contains($searchNeedle)) {
            $matchedBy.Add('DisplayName') | Out-Null
        }

        if ($user.UserPrincipalName -and $user.UserPrincipalName.ToLowerInvariant().Contains($searchNeedle)) {
            $matchedBy.Add('UserPrincipalName') | Out-Null
        }

        if ($user.Mail -and $user.Mail.ToLowerInvariant().Contains($searchNeedle)) {
            $matchedBy.Add('Mail') | Out-Null
        }

        if ($user.OtherMails) {
            foreach ($otherMail in $user.OtherMails) {
                if ($otherMail -and $otherMail.ToLowerInvariant().Contains($searchNeedle)) {
                    $matchedBy.Add('OtherMails') | Out-Null
                    break
                }
            }
        }

        if ($user.ProxyAddresses) {
            foreach ($proxyAddress in $user.ProxyAddresses) {
                if ($proxyAddress -and $proxyAddress.ToLowerInvariant().Contains($searchNeedle)) {
                    $matchedBy.Add('ProxyAddresses') | Out-Null
                    break
                }
            }
        }

        $guestFragment = $null
        if ($user.UserType -eq 'Guest' -and $user.UserPrincipalName -match '^(?<fragment>.+?)#EXT#@') {
            $guestFragment = $matches.fragment
        }

        $results.Add([pscustomobject]@{
            User          = $user
            MatchedBy     = @($matchedBy)
            GuestFragment = $guestFragment
            SearchMode    = if ($IndexOnly.IsPresent) { 'IndexOnly' } else { 'Index + Fallback' }
        }) | Out-Null
    }

    return $results
}
