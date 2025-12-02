#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Mailboxes ============================================================================================================================

function Add-MboxAlias {
    <#
    .SYNOPSIS
        Adds a new alias to a mailbox or mail-enabled recipient.
    .DESCRIPTION
        Validates Exchange Online connectivity, checks the recipient, and adds the alias when it does not already exist.
    .PARAMETER SourceMailbox
        Mailbox or recipient identity. Accepts pipeline input.
    .PARAMETER MailboxAlias
        Alias to add (SMTP address).
    .EXAMPLE
        Add-MboxAlias -SourceMailbox info@contoso.com -MailboxAlias alias@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox,
        [Parameter(Mandatory)]
        [string]$MailboxAlias
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $recipient = Get-Recipient -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Mailbox or recipient '$SourceMailbox' not found. $($_.Exception.Message)" -Level ERROR
            return
        }

        $normalizedAlias = $MailboxAlias.ToLowerInvariant()
        $existingAliases = $recipient.EmailAddresses | ForEach-Object { ($_ -replace '^smtp:', '').ToLowerInvariant() }
        if ($existingAliases -contains $normalizedAlias) {
            Write-NCMessage ("Alias '{0}' already exists for '{1}'. No action taken." -f $MailboxAlias, $recipient.PrimarySmtpAddress) -Level WARNING
            return
        }

        try {
            switch ($recipient.RecipientTypeDetails) {
                'MailContact' { Set-MailContact -Identity $recipient.Identity -EmailAddresses @{ add = $MailboxAlias } -ErrorAction Stop }
                'MailUser' { Set-MailUser -Identity $recipient.Identity -EmailAddresses @{ add = $MailboxAlias } -ErrorAction Stop }
                {($_ -eq 'MailUniversalDistributionGroup') -or ($_ -eq 'DynamicDistributionGroup') -or ($_ -eq 'MailUniversalSecurityGroup')} {
                    Set-DistributionGroup -Identity $recipient.Identity -EmailAddresses @{ add = $MailboxAlias } -ErrorAction Stop
                }
                default { Set-Mailbox -Identity $recipient.Identity -EmailAddresses @{ add = $MailboxAlias } -ErrorAction Stop }
            }

            Write-NCMessage ("Alias '{0}' added to {1}." -f $MailboxAlias, $recipient.PrimarySmtpAddress) -Level SUCCESS
        }
        catch {
            Write-NCMessage "Unable to add alias '$MailboxAlias' to '$($recipient.PrimarySmtpAddress)'. $($_.Exception.Message)" -Level ERROR
            return
        }

        Get-MboxAlias -SourceMailbox $recipient.PrimarySmtpAddress
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Add-MboxPermission {
    <#
    .SYNOPSIS
        Grants mailbox permissions to users.
    .DESCRIPTION
        Ensures Exchange Online connectivity, validates target and user mailboxes, and assigns FullAccess,
        SendAs, SendOnBehalfTo, or both (All). Returns the applied permissions.
    .PARAMETER SourceMailbox
        Mailbox identity to update.
    .PARAMETER UserMailbox
        One or more users to grant permissions to. Accepts pipeline input.
    .PARAMETER AccessRights
        Permission type: All, FullAccess, SendAs, SendOnBehalfTo. Defaults to All (FullAccess + SendAs).
    .PARAMETER AutoMapping
        Enable Outlook automapping when granting FullAccess.
    .EXAMPLE
        Add-MboxPermission -SourceMailbox info@contoso.com -UserMailbox mario.rossi@contoso.com -AccessRights FullAccess -AutoMapping
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox,
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$UserMailbox,
        [ValidateSet('All', 'FullAccess', 'SendAs', 'SendOnBehalfTo')]
        [string]$AccessRights = 'All',
        [switch]$AutoMapping
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $targetMailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Mailbox '$SourceMailbox' not found. $($_.Exception.Message)" -Level ERROR
            return
        }

        foreach ($user in $UserMailbox) {
            try {
                $userObject = Get-User -Identity $user -ErrorAction Stop
            }
            catch {
                Write-NCMessage "`nThe mailbox '$user' does not exist. Please check the provided e-mail address." -Level ERROR
                continue
            }

            $userIdentity = $userObject.UserPrincipalName
            switch ($AccessRights) {
                'FullAccess' {
                    $existingPermission = Get-MailboxPermission -Identity $targetMailbox.PrimarySmtpAddress -User $userIdentity -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
                    if ($existingPermission) {
                        Write-NCMessage ("{0} already has FullAccess permission to {1}, skipping." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level WARNING
                        continue
                    }

                    $added = Add-MailboxPermission -Identity $targetMailbox.PrimarySmtpAddress -User $userIdentity -AccessRights FullAccess -AutoMapping:$AutoMapping.IsPresent -Confirm:$false
                    Write-NCMessage ("Added FullAccess for {0} on {1}." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level SUCCESS
                    [pscustomobject]@{
                        Identity     = $added.Identity
                        User         = $added.User
                        DisplayName  = $userObject.DisplayName
                        AccessRights = $added.AccessRights
                        IsInherited  = $added.IsInherited
                        Deny         = $added.Deny
                    }
                }
                'SendAs' {
                    $existingPermission = Get-RecipientPermission -Identity $targetMailbox.PrimarySmtpAddress -Trustee $userIdentity -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'SendAs' }
                    if ($existingPermission) {
                        Write-NCMessage ("{0} already has SendAs permission to {1}, skipping." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level WARNING
                        continue
                    }

                    $added = Add-RecipientPermission -Identity $targetMailbox.PrimarySmtpAddress -Trustee $userIdentity -AccessRights SendAs -Confirm:$false
                    Write-NCMessage ("Added SendAs for {0} on {1}." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level SUCCESS
                    [pscustomobject]@{
                        Identity          = $added.Identity
                        Trustee           = $added.Trustee
                        DisplayName       = $userObject.DisplayName
                        AccessControlType = $added.AccessControlType
                        AccessRights      = $added.AccessRights
                    }
                }
                'SendOnBehalfTo' {
                    $existingPermission = $targetMailbox.GrantSendOnBehalfTo | Where-Object { $_ -eq $userIdentity }
                    if ($existingPermission) {
                        Write-NCMessage ("{0} already has SendOnBehalfTo permission to {1}, skipping." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level WARNING
                        continue
                    }

                    Set-Mailbox -Identity $targetMailbox.PrimarySmtpAddress -GrantSendOnBehalfTo @{ add = $userIdentity } -Confirm:$false | Out-Null
                    Write-NCMessage ("Added SendOnBehalfTo for {0} on {1}." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level SUCCESS
                    [pscustomobject]@{
                        Identity     = $targetMailbox.PrimarySmtpAddress
                        Trustee      = $userIdentity
                        DisplayName  = $userObject.DisplayName
                        AccessRights = 'SendOnBehalfTo'
                    }
                }
                'All' {
                    $created = @()

                    $existingFullAccess = Get-MailboxPermission -Identity $targetMailbox.PrimarySmtpAddress -User $userIdentity -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
                    if (-not $existingFullAccess) {
                        $added = Add-MailboxPermission -Identity $targetMailbox.PrimarySmtpAddress -User $userIdentity -AccessRights FullAccess -AutoMapping:$AutoMapping.IsPresent -Confirm:$false
                        Write-NCMessage ("Added FullAccess for {0} on {1}." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level SUCCESS
                        $created += [pscustomobject]@{
                            Identity     = $added.Identity
                            User         = $added.User
                            DisplayName  = $userObject.DisplayName
                            AccessRights = $added.AccessRights
                            IsInherited  = $added.IsInherited
                            Deny         = $added.Deny
                        }
                    }
                    else {
                        Write-NCMessage ("{0} already has FullAccess permission to {1}, skipping." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level WARNING
                    }

                    $existingSendAs = Get-RecipientPermission -Identity $targetMailbox.PrimarySmtpAddress -Trustee $userIdentity -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'SendAs' }
                    if (-not $existingSendAs) {
                        $addedSendAs = Add-RecipientPermission -Identity $targetMailbox.PrimarySmtpAddress -Trustee $userIdentity -AccessRights SendAs -Confirm:$false
                        Write-NCMessage ("Added SendAs for {0} on {1}." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level SUCCESS
                        $created += [pscustomobject]@{
                            Identity          = $addedSendAs.Identity
                            Trustee           = $addedSendAs.Trustee
                            DisplayName       = $userObject.DisplayName
                            AccessControlType = $addedSendAs.AccessControlType
                            AccessRights      = $addedSendAs.AccessRights
                        }
                    }
                    else {
                        Write-NCMessage ("{0} already has SendAs permission to {1}, skipping." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level WARNING
                    }

                    $created
                }
            }
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Set-MboxLanguage {
    <#
    .SYNOPSIS
        Sets mailbox regional language.
    .DESCRIPTION
        Changes language for a single mailbox or a list provided via CSV (EmailAddress column).
    .PARAMETER SourceMailbox
        Mailbox to update. Accepts pipeline input. Ignored when -Csv is provided.
    .PARAMETER Language
        Language tag (default it-IT).
    .PARAMETER Csv
        CSV file with EmailAddress column containing mailboxes to update.
    .EXAMPLE
        Set-MboxLanguage -SourceMailbox info@contoso.com -Language en-US
    .EXAMPLE
        Set-MboxLanguage -Csv C:\temp\mailboxes.csv -Language it-IT
    #>
    [CmdletBinding(DefaultParameterSetName = 'Mailbox')]
    param(
        [Parameter(ParameterSetName = 'Mailbox', Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string[]]$SourceMailbox,
        [Parameter()]
        [string]$Language = 'it-IT',
        [Parameter(ParameterSetName = 'Csv', Mandatory = $true)]
        [string]$Csv
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        $mailboxes = @()
        if ($PSCmdlet.ParameterSetName -eq 'Csv') {
            if (-not (Test-Path -LiteralPath $Csv)) {
                Write-NCMessage "CSV file '$Csv' not found." -Level ERROR
                return
            }

            try {
                $mailboxes = (Import-Csv -LiteralPath $Csv) | Where-Object { $_.EmailAddress }
            }
            catch {
                Write-NCMessage "Unable to read CSV file '$Csv'. $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            $mailboxes = $SourceMailbox
        }

        $counter = 0
        $total = $mailboxes.Count
        foreach ($entry in $mailboxes) {
            $address = if ($entry.PSObject.Properties.Match('EmailAddress')) { $entry.EmailAddress } else { $entry }
            if (-not $address) { continue }

            $counter++
            $percentComplete = (($counter / $total) * 100)
            Write-Progress -Activity "Changing language to $Language" -Status "$counter of $total ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            try {
                Set-MailboxRegionalConfiguration -Identity $address -LocalizeDefaultFolderName:$true -Language $Language -ErrorAction Stop
                $result = Get-MailboxRegionalConfiguration -Identity $address -ErrorAction Stop
                [pscustomobject]@{
                    PrimarySmtpAddress = $result.Identity
                    Language           = $result.Language
                    TimeZone           = $result.TimeZone
                }
            }
            catch {
                Write-NCMessage "Failed to update mailbox '$address'. $($_.Exception.Message)" -Level ERROR
            }
        }
    }

    end {
        Write-Progress -Activity "Changing mailbox language" -Completed
        Restore-ProgressAndInfoPreferences
    }
}

function Test-SharedMailboxCompliance {
    <#
    .SYNOPSIS
        Reports shared mailbox sign-in activity and licensing.
    .DESCRIPTION
        Uses Microsoft Graph sign-in logs and assigned plans to flag shared mailboxes with successful sign-ins and missing licenses.
    .PARAMETER GridView
        Show the result in Out-GridView (default behaviour). Specify -GridView:$false to return objects.
    .EXAMPLE
        Test-SharedMailboxCompliance
    #>
    [CmdletBinding()]
    param(
        [switch]$GridView
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        $graphConnected = Test-MgGraphConnection -Scopes @('AuditLog.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
            return
        }

        Write-NCMessage "`nFinding shared mailboxes..." -NoNewline
        $mailboxes = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Sort-Object DisplayName
        if (-not $mailboxes) {
            Write-NCMessage "No shared mailboxes found." -Level WARNING
            return
        }

        Write-NCMessage (" {0} shared mailboxes found." -f $mailboxes.Count) -Level SUCCESS
        $exoPlan1 = '9aaf7827-d63c-4b61-89c3-182f06f82e5c'
        $exoPlan2 = 'efb87545-963c-4e0d-99df-69c6916d9eb0'
        $report = [System.Collections.Generic.List[object]]::new()
        $counter = 0

        foreach ($mbx in $mailboxes) {
            $counter++
            $percentComplete = (($counter / $mailboxes.Count) * 100)
            Write-Progress -Activity "Checking $($mbx.DisplayName)" -Status "$counter of $($mailboxes.Count) ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            $logsFound = $false
            $exoPlan1Found = $false
            $exoPlan2Found = $false

            try {
                $signIns = Get-MgAuditLogSignIn -Filter "userid eq '$($mbx.ExternalDirectoryObjectId)'" -All -Top 20 -ErrorAction Stop
                if ($signIns) {
                    foreach ($log in $signIns) {
                        if ($log.Status.ErrorCode -eq 0) {
                            $logsFound = $true
                            break
                        }
                    }
                }
            }
            catch {
                Write-NCMessage ("Unable to retrieve sign-in records for {0}. {1}" -f $mbx.DisplayName, $_.Exception.Message) -Level ERROR
            }

            if ($logsFound) {
                Write-NCMessage ("Sign-in records found for shared mailbox {0}" -f $mbx.DisplayName) -Level WARNING
                try {
                    $user = Get-MgUser -UserId $mbx.ExternalDirectoryObjectId -Property UserPrincipalName, assignedPlans
                    $exoPlans = @($user.AssignedPlans | Where-Object { $_.Service -eq 'exchange' -and $_.capabilityStatus -eq 'Enabled' })
                    $exoPlan1Found = $exoPlan1 -in $exoPlans.ServicePlanId
                    $exoPlan2Found = $exoPlan2 -in $exoPlans.ServicePlanId
                }
                catch {
                    Write-NCMessage ("Unable to read license info for {0}. {1}" -f $mbx.DisplayName, $_.Exception.Message) -Level ERROR
                }
            }
            else {
                Write-NCMessage ("No successful sign-in records found for shared mailbox {0}" -f $mbx.DisplayName) -Level SUCCESS
            }

            $report.Add([pscustomobject]@{
                    DisplayName               = $mbx.DisplayName
                    ExternalDirectoryObjectId = $mbx.ExternalDirectoryObjectId
                    'Sign in Record Found'    = if ($logsFound) { 'Yes' } else { 'No' }
                    'Exchange Online Plan 1'  = $exoPlan1Found
                    'Exchange Online Plan 2'  = $exoPlan2Found
                }) | Out-Null
        }

        Write-Progress -Activity "Checking shared mailboxes" -Completed

        $showGrid = if ($PSBoundParameters.ContainsKey('GridView')) { $GridView.IsPresent } else { $true }
        if ($showGrid) {
            $report | Out-GridView -Title "Shared Mailbox Sign-In Records and Licensing Status"
        }
        else {
            $report
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Export-MboxAlias {
    <#
    .SYNOPSIS
        Exports aliases for one or more recipients.
    .DESCRIPTION
        Enumerates aliases for specified mailboxes, all mailboxes, or recipients filtered by domain,
        and optionally writes them to a CSV file.
    .PARAMETER SourceMailbox
        Mailbox identities to analyze.
    .PARAMETER Csv
        Export the results to CSV.
    .PARAMETER CsvFolder
        Destination folder for the CSV file. Defaults to current directory.
    .PARAMETER All
        Export aliases for every non-guest recipient.
    .PARAMETER Domain
        Export aliases for recipients whose addresses match the provided domain.
    .EXAMPLE
        Export-MboxAlias -SourceMailbox info@contoso.com
    .EXAMPLE
        Export-MboxAlias -All -CsvFolder C:\Temp
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string[]]$SourceMailbox,
        [switch]$Csv,
        [string]$CsvFolder,
        [switch]$All,
        [string]$Domain
    )

    begin {
        Set-ProgressAndInfoPreferences
        $aliases = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        if ((-not $SourceMailbox -or $SourceMailbox.Count -eq 0) -and -not $All -and [string]::IsNullOrWhiteSpace($Domain)) {
            $All = $true
        }

        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $SourceMailbox = Get-Recipient -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -ne 'GuestMailUser' -and $_.EmailAddresses -like "*@$Domain" }
            $Csv = $true
        }

        if ($All) {
            Write-NCMessage "No mailbox specified; scanning all recipients. This may take a while." -Level WARNING
            $SourceMailbox = Get-Recipient -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -ne 'GuestMailUser' }
            $Csv = $true
        }

        if ($Csv -and -not $script:ExportMboxAliasFolder) {
            try {
                $script:ExportMboxAliasFolder = Test-Folder $CsvFolder
            }
            catch {
                Write-NCMessage "Invalid CSV folder. $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        $counter = 0
        $total = $SourceMailbox.Count
        foreach ($entry in $SourceMailbox) {
            try {
                $recipient = Get-Recipient -Identity $entry -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Recipient '$entry' not found. $($_.Exception.Message)" -Level WARNING
                continue
            }

            $counter++
            $percentComplete = (($counter / $total) * 100)
            Write-Progress -Activity "Processing $($recipient.PrimarySmtpAddress)" -Status "$counter of $total ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            $primary = $recipient.PrimarySmtpAddress
            foreach ($address in $recipient.EmailAddresses | Where-Object { $_ -clike 'smtp*' }) {
                $aliases.Add([pscustomobject]@{
                        PrimarySmtpAddress = $primary
                        Alias              = $address.ToString().Substring(5)
                    }) | Out-Null
            }
        }
    }

    end {
        try {
            if ($Csv) {
                $folder = if ($script:ExportMboxAliasFolder) { $script:ExportMboxAliasFolder } else { Test-Folder $CsvFolder }
                $csvPath = New-File "$folder\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-Alias-Report.csv"
                $aliases | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                Write-NCMessage "Alias report exported to $csvPath" -Level SUCCESS
                $csvPath
            }
            else {
                $aliases
            }
        }
        finally {
            Write-Progress -Activity "Processing aliases" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Export-MboxPermission {
    <#
    .SYNOPSIS
        Exports mailbox permissions for selected recipient types.
    .DESCRIPTION
        Gathers FullAccess, SendAs, and SendOnBehalfTo permissions for user, shared, room, or all mailboxes
        and writes them to a CSV report.
    .PARAMETER RecipientType
        Recipient type to analyze: User, Shared, Room, All.
    .PARAMETER CsvFolder
        Destination folder for the CSV file. Defaults to current directory.
    .EXAMPLE
        Export-MboxPermission -RecipientType All -CsvFolder C:\Temp
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('User', 'Shared', 'Room', 'All')]
        [string]$RecipientType,
        [string]$CsvFolder
    )

    begin {
        Set-ProgressAndInfoPreferences
        $permissions = [System.Collections.Generic.List[object]]::new()
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        switch ($RecipientType) {
            'User' { $mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' } }
            'Shared' { $mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' } }
            'Room' { $mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where-Object { $_.RecipientTypeDetails -eq 'RoomMailbox' } }
            'All' {
                Write-NCMessage "No recipient type specified, scanning User, Shared, and Room mailboxes." -Level WARNING
                $mailboxes = Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue | Where-Object { $_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox', 'RoomMailbox') }
            }
        }

        $counter = 0
        foreach ($mailbox in $mailboxes) {
            $counter++
            $percentComplete = (($counter / $mailboxes.Count) * 100)
            Write-Progress -Activity "Processing $($mailbox.PrimarySmtpAddress)" -Status "$counter of $($mailboxes.Count) ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            $exoMailbox = Get-EXOMailbox -Identity $mailbox.Identity
            $sendAs = Get-RecipientPermission -Identity $exoMailbox.PrimarySmtpAddress -AccessRights SendAs | Where-Object { $_.Trustee.ToString() -ne 'NT AUTHORITY\SELF' -and $_.Trustee.ToString() -notlike 'S-1-5*' } | ForEach-Object { $_.Trustee.ToString() }
            $fullAccess = Get-MailboxPermission -Identity $exoMailbox.PrimarySmtpAddress | Where-Object { $_.AccessRights -eq 'FullAccess' -and -not $_.IsInherited } | ForEach-Object { $_.User.ToString() }

            $permissions.Add([pscustomobject]@{
                    Mailbox           = $exoMailbox.DisplayName
                    'Mailbox Address' = $exoMailbox.PrimarySmtpAddress
                    'Recipient Type'  = $exoMailbox.RecipientTypeDetails
                    FullAccess        = ($fullAccess -join ', ')
                    SendAs            = ($sendAs -join ', ')
                    SendOnBehalfTo    = $exoMailbox.GrantSendOnBehalfTo -join ', '
                }) | Out-Null
        }
    }

    end {
        try {
            $folder = Test-Folder $CsvFolder
            $csvPath = New-File "$folder\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-MboxPermissions-Report.csv"
            $permissions | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
            Write-NCMessage "Mailbox permissions exported to $csvPath" -Level SUCCESS
            $csvPath
        }
        finally {
            Write-Progress -Activity "Processing mailbox permissions" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Get-MboxAlias {
    <#
    .SYNOPSIS
        Lists primary and secondary SMTP addresses for a recipient.
    .DESCRIPTION
        Ensures Exchange Online connectivity and returns aliases with a flag for the primary address.
    .PARAMETER SourceMailbox
        Mailbox or recipient identity. Accepts pipeline input.
    .EXAMPLE
        Get-MboxAlias -SourceMailbox info@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $recipient = Get-Recipient -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Recipient '$SourceMailbox' not available or not found." -Level ERROR
            return
        }

        $aliases = [System.Collections.Generic.List[object]]::new()
        foreach ($address in $recipient.EmailAddresses) {
            if ($address -clike 'SMTP:*') {
                $aliases.Add([pscustomobject]@{
                        PrimarySmtpAddress = $recipient.PrimarySmtpAddress
                        Alias              = $address.Replace('SMTP:', '')
                        IsPrimary          = $true
                    }) | Out-Null
            }
            elseif ($address -clike 'smtp:*') {
                $aliases.Add([pscustomobject]@{
                        PrimarySmtpAddress = $recipient.PrimarySmtpAddress
                        Alias              = $address.Replace('smtp:', '')
                        IsPrimary          = $false
                    }) | Out-Null
            }
        }

        if ($aliases.Count -eq 0) {
            Write-NCMessage "No aliases found for '$($recipient.PrimarySmtpAddress)'." -Level WARNING
        }
        else {
            $aliases | Sort-Object -Property @{ Expression = 'IsPrimary'; Descending = $true }, Alias
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Get-MboxPermission {
    <#
    .SYNOPSIS
        Retrieves mailbox permissions for a single mailbox.
    .DESCRIPTION
        Shows FullAccess, SendAs, and SendOnBehalfTo permissions with optional summary counts.
    .PARAMETER SourceMailbox
        Mailbox identity to inspect.
    .PARAMETER IncludeSummary
        Display a short summary of counts.
    .EXAMPLE
        Get-MboxPermission -SourceMailbox info@contoso.com -IncludeSummary
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox,
        [switch]$IncludeSummary
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Mailbox '$SourceMailbox' not found." -Level ERROR
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $fullAccessCount = 0
        $sendAsCount = 0
        $sendOnBehalfToCount = 0

        Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -eq 'FullAccess' -and -not $_.IsInherited } | ForEach-Object {
            $userMailbox = $_.User.ToString()
            $primary = (Get-Mailbox -Identity $userMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
            $display = (Get-User -Identity $userMailbox -ErrorAction SilentlyContinue).DisplayName

            if ($primary) {
                $existing = $results | Where-Object { $_.UserMailbox -eq $primary }
                if ($existing) {
                    $existing.AccessRights += ', FullAccess'
                }
                else {
                    $results.Add([pscustomobject]@{
                            User         = $display
                            UserMailbox  = $primary
                            AccessRights = 'FullAccess'
                        }) | Out-Null
                }
                $fullAccessCount++
            }
        }
        Write-Progress -Activity "Gathered FullAccess permissions for $($mailbox.PrimarySmtpAddress) ..." -Status "35% Complete" -PercentComplete 35

        Get-RecipientPermission -Identity $mailbox.PrimarySmtpAddress -AccessRights SendAs -ErrorAction SilentlyContinue | Where-Object { $_.Trustee.ToString() -ne 'NT AUTHORITY\SELF' -and $_.Trustee.ToString() -notlike 'S-1-5*' } | ForEach-Object {
            $userMailbox = $_.Trustee.ToString()
            $primary = (Get-Mailbox -Identity $userMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
            $display = (Get-User -Identity $userMailbox -ErrorAction SilentlyContinue).DisplayName

            if ($primary) {
                $existing = $results | Where-Object { $_.UserMailbox -eq $primary }
                if ($existing) {
                    $existing.AccessRights += ', SendAs'
                }
                else {
                    $results.Add([pscustomobject]@{
                            User         = $display
                            UserMailbox  = $primary
                            AccessRights = 'SendAs'
                        }) | Out-Null
                }
                $sendAsCount++
            }
        }
        Write-Progress -Activity "Gathered SendAs permissions for $($mailbox.PrimarySmtpAddress) ..." -Status "50% Complete" -PercentComplete 50

        foreach ($userMailbox in $mailbox.GrantSendOnBehalfTo) {
            $primary = (Get-Mailbox -Identity $userMailbox -ErrorAction SilentlyContinue).PrimarySmtpAddress
            $display = (Get-User -Identity $userMailbox -ErrorAction SilentlyContinue).DisplayName

            if ($primary) {
                $existing = $results | Where-Object { $_.UserMailbox -eq $primary }
                if ($existing) {
                    $existing.AccessRights += ', SendOnBehalfTo'
                }
                else {
                    $results.Add([pscustomobject]@{
                            User         = $display
                            UserMailbox  = $primary
                            AccessRights = 'SendOnBehalfTo'
                        }) | Out-Null
                }
                $sendOnBehalfToCount++
            }
        }
        Write-Progress -Activity "Gathered SendOnBehalfTo permissions for $($mailbox.PrimarySmtpAddress) ..." -Status "90% Complete" -PercentComplete 90

        Write-NCMessage ("`nAccess Rights on {0} ({1})" -f $mailbox.DisplayName, $mailbox.PrimarySmtpAddress) -Level WARNING
        $results

        if ($IncludeSummary) {
            Write-NCMessage "`nSummary of Permissions Found:" -Level INFO
            Write-NCMessage ("FullAccess: {0}" -f $fullAccessCount) -Level SUCCESS
            Write-NCMessage ("SendAs: {0}" -f $sendAsCount) -Level SUCCESS
            Write-NCMessage ("SendOnBehalfTo: {0}" -f $sendOnBehalfToCount) -Level SUCCESS
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function New-SharedMailbox {
    <#
    .SYNOPSIS
        Creates a shared mailbox with opinionated defaults.
    .DESCRIPTION
        Ensures Exchange Online connectivity, creates the shared mailbox, enables copies of sent messages, and sets deleted items retention.
    .PARAMETER SharedMailboxSMTPAddress
        Primary SMTP address for the new shared mailbox.
    .PARAMETER SharedMailboxDisplayName
        Display name for the mailbox.
    .PARAMETER SharedMailboxAlias
        Alias for the mailbox.
    .EXAMPLE
        New-SharedMailbox -SharedMailboxSMTPAddress info@contoso.com -SharedMailboxDisplayName "Contoso - Info" -SharedMailboxAlias contoso_info
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SharedMailboxSMTPAddress,
        [Parameter(Mandatory)]
        [string]$SharedMailboxDisplayName,
        [Parameter(Mandatory)]
        [string]$SharedMailboxAlias
    )

    if (-not (Test-EOLConnection)) {
        Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
        return
    }

    try {
        New-Mailbox -Name $SharedMailboxDisplayName -Alias $SharedMailboxAlias -Shared -PrimarySmtpAddress $SharedMailboxSMTPAddress -ErrorAction Stop
        Write-NCMessage ("Set outgoing e-mail copy save for {0}" -f $SharedMailboxSMTPAddress) -Level INFO
        Set-Mailbox -Identity $SharedMailboxSMTPAddress -MessageCopyForSentAsEnabled $true
        Set-Mailbox -Identity $SharedMailboxSMTPAddress -MessageCopyForSendOnBehalfEnabled $true
        Set-Mailbox -Identity $SharedMailboxSMTPAddress -RetainDeletedItemsFor 30
        Write-NCMessage "All done, remember to set access and editing rights to the new mailbox." -Level SUCCESS
    }
    catch {
        Write-NCMessage "Unable to create shared mailbox. $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-MboxAlias {
    <#
    .SYNOPSIS
        Removes an alias from a mailbox or mail-enabled recipient.
    .DESCRIPTION
        Validates Exchange Online connectivity, resolves the recipient, and removes the specified alias.
    .PARAMETER SourceMailbox
        Mailbox or recipient identity. Accepts pipeline input.
    .PARAMETER MailboxAlias
        Alias to remove (SMTP address).
    .EXAMPLE
        Remove-MboxAlias -SourceMailbox info@contoso.com -MailboxAlias alias@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox,
        [Parameter(Mandatory)]
        [string]$MailboxAlias
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $recipient = Get-Recipient -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Mailbox or recipient '$SourceMailbox' not found. $($_.Exception.Message)" -Level ERROR
            return
        }

        try {
            switch ($recipient.RecipientTypeDetails) {
                'MailContact' { Set-MailContact -Identity $recipient.Identity -EmailAddresses @{ remove = $MailboxAlias } -ErrorAction Stop }
                'MailUser' { Set-MailUser -Identity $recipient.Identity -EmailAddresses @{ remove = $MailboxAlias } -ErrorAction Stop }
                {($_ -eq 'MailUniversalDistributionGroup') -or ($_ -eq 'DynamicDistributionGroup') -or ($_ -eq 'MailUniversalSecurityGroup')} {
                    Set-DistributionGroup -Identity $recipient.Identity -EmailAddresses @{ remove = $MailboxAlias } -ErrorAction Stop
                }
                default { Set-Mailbox -Identity $recipient.Identity -EmailAddresses @{ remove = $MailboxAlias } -ErrorAction Stop }
            }

            Write-NCMessage ("Alias '{0}' removed from {1}." -f $MailboxAlias, $recipient.PrimarySmtpAddress) -Level SUCCESS
        }
        catch {
            Write-NCMessage "Unable to remove alias '$MailboxAlias' from '$($recipient.PrimarySmtpAddress)'. $($_.Exception.Message)" -Level ERROR
            return
        }

        Get-MboxAlias -SourceMailbox $recipient.PrimarySmtpAddress
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Remove-MboxPermission {
    <#
    .SYNOPSIS
        Removes mailbox permissions from users.
    .DESCRIPTION
        Ensures Exchange Online connectivity, validates target and user mailboxes, and removes FullAccess,
        SendAs, SendOnBehalfTo, or all of them.
    .PARAMETER SourceMailbox
        Mailbox identity to update.
    .PARAMETER UserMailbox
        One or more users to remove permissions from. Accepts pipeline input.
    .PARAMETER AccessRights
        Permission type to remove: All, FullAccess, SendAs, SendOnBehalfTo. Defaults to All.
    .EXAMPLE
        Remove-MboxPermission -SourceMailbox info@contoso.com -UserMailbox mario.rossi@contoso.com -AccessRights SendAs
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string]$SourceMailbox,
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$UserMailbox,
        [ValidateSet('All', 'FullAccess', 'SendAs', 'SendOnBehalfTo')]
        [string]$AccessRights = 'All'
    )

    begin { Set-ProgressAndInfoPreferences }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $targetMailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Mailbox '$SourceMailbox' not found. $($_.Exception.Message)" -Level ERROR
            return
        }

        foreach ($user in $UserMailbox) {
            try {
                $userObject = Get-User -Identity $user -ErrorAction Stop
            }
            catch {
                Write-NCMessage "`nThe mailbox '$user' does not exist. Please check the provided e-mail address." -Level ERROR
                continue
            }

            $userIdentity = $userObject.UserPrincipalName
            switch ($AccessRights) {
                'FullAccess' {
                    Write-NCMessage ("Removing FullAccess for {0} from {1} ..." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level INFO
                    Remove-MailboxPermission -Identity $targetMailbox.PrimarySmtpAddress -User $userIdentity -AccessRights FullAccess -Confirm:$false
                }
                'SendAs' {
                    Write-NCMessage ("Removing SendAs for {0} from {1} ..." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level INFO
                    Remove-RecipientPermission -Identity $targetMailbox.PrimarySmtpAddress -Trustee $userIdentity -AccessRights SendAs -Confirm:$false
                }
                'SendOnBehalfTo' {
                    Write-NCMessage ("Removing SendOnBehalfTo for {0} from {1} ..." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level INFO
                    Set-Mailbox -Identity $targetMailbox.PrimarySmtpAddress -GrantSendOnBehalfTo @{ remove = $userIdentity } -Confirm:$false | Out-Null
                }
                'All' {
                    Write-NCMessage ("Removing FullAccess for {0} from {1} ..." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level INFO
                    Remove-MailboxPermission -Identity $targetMailbox.PrimarySmtpAddress -User $userIdentity -AccessRights FullAccess -Confirm:$false
                    Write-NCMessage ("Removing SendAs for {0} from {1} ..." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level INFO
                    Remove-RecipientPermission -Identity $targetMailbox.PrimarySmtpAddress -Trustee $userIdentity -AccessRights SendAs -Confirm:$false
                    Write-NCMessage ("Removing SendOnBehalfTo for {0} from {1} ..." -f $userIdentity, $targetMailbox.PrimarySmtpAddress) -Level INFO
                    Set-Mailbox -Identity $targetMailbox.PrimarySmtpAddress -GrantSendOnBehalfTo @{ remove = $userIdentity } -Confirm:$false | Out-Null
                }
            }
        }
    }

    end { Restore-ProgressAndInfoPreferences }
}

function Set-MboxRulesQuota {
    <#
    .SYNOPSIS
        Sets mailbox rules quota to 256KB.
    .DESCRIPTION
        Iterates the provided mailboxes, sets the RulesQuota to 256KB, and returns the updated values.
    .PARAMETER SourceMailbox
        Mailboxes to update. Accepts pipeline input.
    .EXAMPLE
        Set-MboxRulesQuota -SourceMailbox info@contoso.com, support@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string[]]$SourceMailbox
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
        $counter = 0
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        foreach ($mailbox in $SourceMailbox) {
            try {
                $recipient = Get-Recipient -Identity $mailbox -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Mailbox '$mailbox' not found. $($_.Exception.Message)" -Level ERROR
                continue
            }

            $counter++
            $percentComplete = (($counter / $SourceMailbox.Count) * 100)
            Write-Progress -Activity "Processing $($recipient.PrimarySmtpAddress)" -Status "$counter of $($SourceMailbox.Count) ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            try {
                Set-Mailbox -Identity $recipient.PrimarySmtpAddress -RulesQuota 256KB
                $results.Add([pscustomobject]@{
                        PrimarySmtpAddress = $recipient.PrimarySmtpAddress
                        'Rules Quota'      = (Get-Mailbox -Identity $recipient.PrimarySmtpAddress).RulesQuota
                    }) | Out-Null
            }
            catch {
                Write-NCMessage $_.Exception.Message -Level ERROR
            }
        }
    }

    end {
        Write-Progress -Activity "Processing mailbox rules quota" -Completed
        Restore-ProgressAndInfoPreferences
        $results
    }
}

function Set-SharedMboxCopyForSent {
    <#
    .SYNOPSIS
        Enables sent-item copy options on shared mailboxes.
    .DESCRIPTION
        For each shared mailbox, enables MessageCopyForSentAsEnabled and MessageCopyForSendOnBehalfEnabled,
        returning the updated status and listing any errors.
    .PARAMETER SourceMailbox
        Shared mailboxes to update. Accepts pipeline input.
    .EXAMPLE
        Set-SharedMboxCopyForSent -SourceMailbox info@contoso.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Identity')]
        [string[]]$SourceMailbox
    )

    begin {
        Set-ProgressAndInfoPreferences
        $results = [System.Collections.Generic.List[object]]::new()
        $errors = [System.Collections.Generic.List[string]]::new()
        $counter = 0
    }

    process {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        foreach ($mailbox in $SourceMailbox) {
            try {
                $recipient = Get-Recipient -Identity $mailbox -ErrorAction Stop
                if ($recipient.RecipientTypeDetails -ne 'SharedMailbox') {
                    $errors.Add("{0} is not a Shared Mailbox." -f $mailbox) | Out-Null
                    continue
                }

                $counter++
                $percentComplete = (($counter / $SourceMailbox.Count) * 100)
                Write-Progress -Activity "Processing $($recipient.PrimarySmtpAddress)" -Status "$counter of $($SourceMailbox.Count) ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

                Set-Mailbox -Identity $recipient.PrimarySmtpAddress -MessageCopyForSentAsEnabled $true
                Set-Mailbox -Identity $recipient.PrimarySmtpAddress -MessageCopyForSendOnBehalfEnabled $true

                $updated = Get-Mailbox -Identity $recipient.PrimarySmtpAddress
                $results.Add([pscustomobject]@{
                        PrimarySmtpAddress      = $recipient.PrimarySmtpAddress
                        'Copy for SentAs'       = $updated.MessageCopyForSentAsEnabled
                        'Copy for SendOnBehalf' = $updated.MessageCopyForSendOnBehalfEnabled
                    }) | Out-Null
            }
            catch {
                Write-NCMessage $_.Exception.Message -Level ERROR
            }
        }
    }

    end {
        Write-Progress -Activity "Processing shared mailbox sent items copy" -Completed
        Restore-ProgressAndInfoPreferences

        $results
        if ($errors.Count -gt 0) {
            Write-NCMessage ($errors -join [Environment]::NewLine) -Level WARNING
        }
    }
}
