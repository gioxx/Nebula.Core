#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Security helpers =====================================================================================================================

function Disable-UserDevices {
    <#
    .SYNOPSIS
        Disables all registered devices for specified users.
    .DESCRIPTION
        Ensures Microsoft Graph connectivity, resolves the target users, and sets AccountEnabled = $false
        on each registered device. Skips missing users and reports progress.
    .PARAMETER UserPrincipalName
        One or more user principal names. Accepts pipeline input.
    .PARAMETER PassThru
        Emit the impacted devices as objects.
    .EXAMPLE
        Disable-UserDevices -UserPrincipalName user1@contoso.com,user2@contoso.com -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $targets.Add($upn.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if ($targets.Count -eq 0) {
                Write-NCMessage "No user principal names provided." -Level WARNING
                return
            }

            $scopes = @('Directory.ReadWrite.All', 'Device.ReadWrite.All')
            if (-not (Test-MgGraphConnection -Scopes $scopes -EnsureExchangeOnline:$false)) {
                Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $queue = foreach ($entry in $targets) { if ($dedup.Add($entry)) { $entry } }
            $counter = 0

            foreach ($upn in $queue) {
                $counter++
                $percent = (($counter / $queue.Count) * 100)
                Write-Progress -Activity "Resolving user $upn" -Status "$counter of $($queue.Count) ($($percent.ToString('0.00'))%)" -PercentComplete $percent

                $escaped = $upn.Replace("'", "''")
                try {
                    $user = Get-MgUser -Filter "userPrincipalName eq '$escaped'" -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -First 1
                }
                catch {
                    Write-NCMessage "Can't find Azure AD account for user $upn. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                if (-not $user) {
                    Write-NCMessage "Can't find Azure AD account for user $upn." -Level ERROR
                    continue
                }

                try {
                    $devices = Get-MgUserRegisteredDevice -UserId $user.Id -All
                }
                catch {
                    Write-NCMessage "Unable to retrieve registered devices for $($user.UserPrincipalName). $($_.Exception.Message)" -Level ERROR
                    continue
                }

                if (-not $devices -or $devices.Count -eq 0) {
                    Write-NCMessage ("No registered devices found for {0}." -f $user.UserPrincipalName) -Level WARNING
                    continue
                }

                foreach ($device in $devices) {
                    $deviceLabel = if ($device.DisplayName) { $device.DisplayName } else { $device.Id }
                    if (-not $PSCmdlet.ShouldProcess($deviceLabel, "Disable device for user $($user.UserPrincipalName)")) {
                        continue
                    }

                    try {
                        Update-MgDevice -DeviceId $device.Id -AccountEnabled:$false -ErrorAction Stop | Out-Null
                        $results.Add([pscustomobject]@{
                                UserPrincipalName = $user.UserPrincipalName
                                UserDisplayName   = $user.DisplayName
                                DeviceId          = $device.Id
                                DeviceDisplayName = $device.DisplayName
                                Action            = 'Disabled'
                            }) | Out-Null
                    }
                    catch {
                        Write-NCMessage "Failed to disable device $deviceLabel for $($user.UserPrincipalName). $($_.Exception.Message)" -Level ERROR
                    }
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($results.Count -gt 0) {
                Write-NCMessage ("Disabled {0} device(s)." -f $results.Count) -Level SUCCESS
            }
        }
        finally {
            Write-Progress -Activity "Resolving user" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Disable-UserSignIn {
    <#
    .SYNOPSIS
        Blocks sign-in for specified users.
    .DESCRIPTION
        Ensures Microsoft Graph connectivity, resolves the target users, and sets AccountEnabled = $false.
        Skips missing users and supports WhatIf/Confirm.
    .PARAMETER UserPrincipalName
        One or more user principal names. Accepts pipeline input.
    .PARAMETER PassThru
        Emit the impacted users as objects.
    .EXAMPLE
        Disable-UserSignIn -UserPrincipalName user1@contoso.com,user2@contoso.com -Confirm:$false
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $targets.Add($upn.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if ($targets.Count -eq 0) {
                Write-NCMessage "No user principal names provided." -Level WARNING
                return
            }

            $scopes = @('Directory.ReadWrite.All')
            if (-not (Test-MgGraphConnection -Scopes $scopes -EnsureExchangeOnline:$false)) {
                Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $queue = foreach ($entry in $targets) { if ($dedup.Add($entry)) { $entry } }
            $counter = 0

            foreach ($upn in $queue) {
                $counter++
                $percent = (($counter / $queue.Count) * 100)
                Write-Progress -Activity "Processing $upn" -Status "$counter of $($queue.Count) ($($percent.ToString('0.00'))%)" -PercentComplete $percent

                $escaped = $upn.Replace("'", "''")
                try {
                    $user = Get-MgUser -Filter "userPrincipalName eq '$escaped'" -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -First 1
                }
                catch {
                    Write-NCMessage "Can't find Azure AD account for user $upn. $($_.Exception.Message)" -Level ERROR
                    continue
                }

                if (-not $user) {
                    Write-NCMessage "Can't find Azure AD account for user $upn." -Level ERROR
                    continue
                }

                if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, "Disable sign-in")) {
                    continue
                }

                try {
                    Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop | Out-Null
                    $results.Add([pscustomobject]@{
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName       = $user.DisplayName
                            Action            = 'SignInDisabled'
                        }) | Out-Null
                }
                catch {
                    Write-NCMessage "Failed to disable sign-in for $($user.UserPrincipalName). $($_.Exception.Message)" -Level ERROR
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($results.Count -gt 0) {
                Write-NCMessage ("Sign-in disabled for {0} user(s)." -f $results.Count) -Level SUCCESS
            }
        }
        finally {
            Write-Progress -Activity "Processing users" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Revoke-UserSessions {
    <#
    .SYNOPSIS
        Forces sign-out by revoking refresh tokens for users.
    .DESCRIPTION
        Ensures Microsoft Graph connectivity, targets all users or a selection (with optional exclusions),
        and calls Revoke-MgUserSignInSession. Supports WhatIf/Confirm.
    .PARAMETER All
        Target every user in the tenant.
    .PARAMETER UserPrincipalName
        Specific users to target. Accepts pipeline input.
    .PARAMETER Exclude
        Users to skip when using -All or a list.
    .PARAMETER PassThru
        Emit the impacted users as objects.
    .EXAMPLE
        Revoke-UserSessions -UserPrincipalName user1@contoso.com,user2@contoso.com
    .EXAMPLE
        Revoke-UserSessions -All -Exclude breakglass@contoso.com -Confirm:$false
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [switch]$All,
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Identity')]
        [string[]]$UserPrincipalName,
        [string[]]$Exclude,
        [switch]$PassThru
    )

    begin {
        Set-ProgressAndInfoPreferences
        $targets = [System.Collections.Generic.List[string]]::new()
        $exclusions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $targets.Add($upn.Trim()) | Out-Null
            }
        }
    }

    end {
        try {
            if (-not $All.IsPresent -and $targets.Count -eq 0) {
                Write-NCMessage "No target specified. Use -All or provide user principal names." -Level WARNING
                return
            }

            foreach ($ex in $Exclude) {
                if (-not [string]::IsNullOrWhiteSpace($ex)) { $exclusions.Add($ex.Trim()) | Out-Null }
            }

            $scopes = @('Directory.ReadWrite.All')
            if (-not (Test-MgGraphConnection -Scopes $scopes -EnsureExchangeOnline:$false)) {
                Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
                return
            }

            $queue = [System.Collections.Generic.List[Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]]::new()

            if ($All.IsPresent) {
                try {
                    $allUsers = Get-MgUser -All -ConsistencyLevel eventual -ErrorAction Stop
                    foreach ($u in $allUsers) { $queue.Add($u) | Out-Null }
                }
                catch {
                    Write-NCMessage "Unable to retrieve all users. $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            else {
                $dedup = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $uniqueTargets = foreach ($entry in $targets) { if ($dedup.Add($entry)) { $entry } }

                foreach ($upn in $uniqueTargets) {
                    $escaped = $upn.Replace("'", "''")
                    try {
                        $user = Get-MgUser -Filter "userPrincipalName eq '$escaped'" -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -First 1
                        if ($user) {
                            $queue.Add($user) | Out-Null
                        }
                        else {
                            Write-NCMessage "Can't find Azure AD account for user $upn." -Level ERROR
                        }
                    }
                    catch {
                        Write-NCMessage "Can't find Azure AD account for user $upn. $($_.Exception.Message)" -Level ERROR
                    }
                }
            }

            if ($queue.Count -eq 0) {
                Write-NCMessage "No users to process." -Level WARNING
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $counter = 0

            foreach ($user in $queue) {
                $counter++
                $percent = (($counter / $queue.Count) * 100)
                Write-Progress -Activity "Revoking sessions" -Status "$counter of $($queue.Count) ($($percent.ToString('0.00'))%)" -PercentComplete $percent

                if ($exclusions.Contains($user.UserPrincipalName)) {
                    Write-NCMessage ("Skipping user {0}" -f $user.UserPrincipalName) -Level INFO
                    continue
                }

                if (-not $PSCmdlet.ShouldProcess($user.UserPrincipalName, "Revoke sign-in sessions")) {
                    continue
                }

                try {
                    Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop | Out-Null
                    $results.Add([pscustomobject]@{
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName       = $user.DisplayName
                            Action            = 'SessionsRevoked'
                        }) | Out-Null
                }
                catch {
                    Write-NCMessage "Failed to revoke sessions for $($user.UserPrincipalName). $($_.Exception.Message)" -Level ERROR
                }
            }

            if ($PassThru.IsPresent) {
                $results
            }
            elseif ($results.Count -gt 0) {
                Write-NCMessage ("Revoked sessions for {0} user(s)." -f $results.Count) -Level SUCCESS
            }
        }
        finally {
            Write-Progress -Activity "Revoking sessions" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}
