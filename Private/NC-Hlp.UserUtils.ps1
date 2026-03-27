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
