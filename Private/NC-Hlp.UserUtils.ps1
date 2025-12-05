#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) User's Utilities ===========================================================================================================

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
        Resolves and returns the primary SMTP address of a user recipient.
    .DESCRIPTION
        Given a User Principal Name (UPN) or identifier, this function queries Exchange to find the recipient.
        It returns the primary SMTP address, ensuring a complete e-mail format.
    .PARAMETER UserPrincipalName
        The UPN or identifier of the user recipient to resolve.
    .EXAMPLE
        Find-UserRecipient -UserPrincipalName "john.doe"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
    
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        return
    }

    try {
        $recipient = Get-Recipient -Identity $UserPrincipalName -ErrorAction Stop
    }
    catch {
        # If no Exchange recipient exists yet (e.g., user without mailbox), fall back to Graph user info
        try {
            $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
            $fallbackUpn = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
            if ($fallbackUpn) {
                return $fallbackUpn
            }
        }
        catch {
            Write-NCMessage "Recipient not available or not found ($UserPrincipalName). $($_.Exception.Message)" -Level ERROR
        }

        return
    }

    if (-not $recipient) {
        Write-NCMessage "Recipient not available or not found ($UserPrincipalName)." -Level ERROR
        return
    }

    $UserPrincipalName = $recipient.PrimarySmtpAddress
    if (-not $UserPrincipalName) {
        $UserPrincipalName = $recipient.WindowsLiveID
    }

    if ($UserPrincipalName -notmatch '@') {
        $primaryMatches = @($recipient.EmailAddresses | Where-Object { $_ -clike 'SMTP:*' })
        if ($primaryMatches.Count -eq 0) {
            Write-NCMessage "Complete e-mail address not specified and no primary SMTP address found for $UserPrincipalName." -Level WARNING
            return
        }

        if ($primaryMatches.Count -gt 1) {
            Write-NCMessage "Multiple SMTP addresses detected for $UserPrincipalName. Using the first match ($($primaryMatches[0].Substring(5)))." -Level WARNING
        }

        $UserPrincipalName = $primaryMatches[0].Substring(5)
    }

    return $UserPrincipalName
}
