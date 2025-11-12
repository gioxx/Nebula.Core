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
