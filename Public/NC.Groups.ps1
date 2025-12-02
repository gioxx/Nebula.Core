#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Groups ===============================================================================================================================

function Export-DistributionGroups {
    <#
    .SYNOPSIS
        Exports Exchange distribution group membership information.
    .DESCRIPTION
        Ensures an Exchange Online session, enumerates either all distribution groups or the
        provided identities, and gathers their members for CSV, GridView, or pipeline output.
    .PARAMETER DistributionGroup
        Distribution group identity (name, alias, or SMTP). Accepts pipeline input.
    .PARAMETER Csv
        Export the results to CSV (implied when -All or -CsvFolder is specified).
    .PARAMETER CsvFolder
        Destination folder for the CSV report. Defaults to the current directory.
    .PARAMETER All
        Export every distribution group in the tenant.
    .PARAMETER GridView
        Show the result in Out-GridView instead of returning objects.
    .EXAMPLE
        Export-DistributionGroups -DistributionGroup "Support Team"
    .EXAMPLE
        Export-DistributionGroups -All -CsvFolder C:\Temp
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('DG', 'Identity')]
        [string[]]$DistributionGroup,
        [switch]$Csv,
        [string]$CsvFolder,
        [switch]$All,
        [switch]$GridView
    )

    begin {
        Set-ProgressAndInfoPreferences
        $requestedGroups = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if ($DistributionGroup) {
            foreach ($entry in $DistributionGroup) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    [void]$requestedGroups.Add($entry.Trim())
                }
            }
        }
    }

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
                return
            }

            $exportAll = $All.IsPresent -or $requestedGroups.Count -eq 0
            $emitCsv = $Csv.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvFolder) -or $exportAll
            $folder = $null

            if ($emitCsv) {
                try {
                    $folder = Test-Folder $CsvFolder
                }
                catch {
                    Write-NCMessage "Invalid CSV folder. $($_.Exception.Message)" -Level ERROR
                    return
                }
            }

            $groups = @()
            if ($exportAll) {
                try {
                    $groups = Get-DistributionGroup -ResultSize Unlimited -WarningAction SilentlyContinue
                }
                catch {
                    Write-NCMessage "Failed to retrieve distribution groups: $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            else {
                foreach ($identity in ($requestedGroups.ToArray() | Select-Object -Unique)) {
                    try {
                        $groups += Get-DistributionGroup -Identity $identity -ErrorAction Stop
                    }
                    catch {
                        Write-NCMessage "Distribution group '$identity' not found: $($_.Exception.Message)" -Level WARNING
                    }
                }
            }

            if (-not $groups -or $groups.Count -eq 0) {
                Write-NCMessage "No distribution groups found matching the provided criteria." -Level WARNING
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $counter = 0
            $total = $groups.Count

            foreach ($group in $groups) {
                $counter++
                $percentComplete = (($counter / $total) * 100)
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $total ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

                try {
                    $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Failed to retrieve members for '$($group.DisplayName)': $($_.Exception.Message)" -Level WARNING
                    continue
                }

                if (-not $members -or $members.Count -eq 0) {
                    if ($exportAll) {
                        $results.Add([pscustomobject][ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                                'Member Display Name'         = $null
                                'Member FirstName'            = $null
                                'Member LastName'             = $null
                                'Member Primary Smtp Address' = $null
                                'Member Company'              = $null
                                'Member City'                 = $null
                            }) | Out-Null
                    }
                    continue
                }

                foreach ($member in $members) {
                    if ($exportAll) {
                        $row = [ordered]@{
                            'Group Name'                  = $group.DisplayName
                            'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                            'Member Display Name'         = $member.DisplayName
                            'Member FirstName'            = $member.FirstName
                            'Member LastName'             = $member.LastName
                            'Member Primary Smtp Address' = $member.PrimarySmtpAddress
                            'Member Company'              = $member.Company
                            'Member City'                 = $member.City
                        }
                    }
                    else {
                        $row = [ordered]@{
                            'Member Display Name'         = $member.DisplayName
                            'Member FirstName'            = $member.FirstName
                            'Member LastName'             = $member.LastName
                            'Member Primary Smtp Address' = $member.PrimarySmtpAddress
                            'Member Company'              = $member.Company
                            'Member City'                 = $member.City
                        }
                    }

                    $results.Add([pscustomobject]$row) | Out-Null
                }
            }

            if (-not $results -or $results.Count -eq 0) {
                Write-NCMessage "No members found for the specified distribution groups." -Level WARNING
                return
            }

            if ($GridView.IsPresent) {
                $results | Out-GridView -Title "M365 Distribution Groups"
            }
            elseif ($emitCsv) {
                $csvPath = New-File("$($folder)\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-DistributionGroups-Report.csv")
                $results | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                Write-NCMessage "Distribution group membership exported to $csvPath." -Level SUCCESS
            }
            else {
                $results
            }
        }
        finally {
            Write-Progress -Activity "Processing distribution groups" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Export-DynamicDistributionGroups {
    <#
    .SYNOPSIS
        Exports dynamic distribution group membership information.
    .DESCRIPTION
        Ensures an Exchange Online session, enumerates either all dynamic distribution groups or
        the provided identities, and gathers their evaluated members for CSV, GridView, or output.
    .PARAMETER DynamicDistributionGroup
        Dynamic distribution group identity (name, alias, or SMTP). Accepts pipeline input.
    .PARAMETER Csv
        Export the results to CSV (implied when -All or -CsvFolder is specified).
    .PARAMETER CsvFolder
        Destination folder for the CSV report. Defaults to the current directory.
    .PARAMETER All
        Export every dynamic distribution group in the tenant.
    .PARAMETER GridView
        Show the result in Out-GridView instead of returning objects.
    .EXAMPLE
        Export-DynamicDistributionGroups -DynamicDistributionGroup "HR Auto"
    .EXAMPLE
        Export-DynamicDistributionGroups -All -CsvFolder C:\Temp
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('DDG', 'Identity')]
        [string[]]$DynamicDistributionGroup,
        [switch]$Csv,
        [string]$CsvFolder,
        [switch]$All,
        [switch]$GridView
    )

    begin {
        Set-ProgressAndInfoPreferences
        $requestedGroups = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if ($DynamicDistributionGroup) {
            foreach ($entry in $DynamicDistributionGroup) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    [void]$requestedGroups.Add($entry.Trim())
                }
            }
        }
    }

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
                return
            }

            $exportAll = $All.IsPresent -or $requestedGroups.Count -eq 0
            $emitCsv = $Csv.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvFolder) -or $exportAll
            $folder = $null

            if ($emitCsv) {
                try {
                    $folder = Test-Folder $CsvFolder
                }
                catch {
                    Write-NCMessage "Invalid CSV folder. $($_.Exception.Message)" -Level ERROR
                    return
                }
            }

            $groups = @()
            if ($exportAll) {
                try {
                    $groups = Get-DynamicDistributionGroup -ResultSize Unlimited -WarningAction SilentlyContinue
                }
                catch {
                    Write-NCMessage "Failed to retrieve dynamic distribution groups: $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            else {
                foreach ($identity in ($requestedGroups.ToArray() | Select-Object -Unique)) {
                    try {
                        $groups += Get-DynamicDistributionGroup -Identity $identity -ErrorAction Stop
                    }
                    catch {
                        Write-NCMessage "Dynamic distribution group '$identity' not found: $($_.Exception.Message)" -Level WARNING
                    }
                }
            }

            if (-not $groups -or $groups.Count -eq 0) {
                Write-NCMessage "No dynamic distribution groups found matching the provided criteria." -Level WARNING
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $counter = 0
            $total = $groups.Count

            foreach ($group in $groups) {
                $counter++
                $percentComplete = (($counter / $total) * 100)
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $total ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

                try {
                    $members = @(Get-DynamicDistributionGroupMember -Identity $group.Identity -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Failed to retrieve members for '$($group.DisplayName)': $($_.Exception.Message)" -Level WARNING
                    continue
                }

                if (-not $members -or $members.Count -eq 0) {
                    if ($exportAll) {
                        $results.Add([pscustomobject][ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                                'Member Display Name'         = $null
                                'Member FirstName'            = $null
                                'Member LastName'             = $null
                                'Member Primary Smtp Address' = $null
                                'Member Company'              = $null
                                'Member City'                 = $null
                            }) | Out-Null
                    }
                    continue
                }

                foreach ($member in $members) {
                    if ($exportAll) {
                        $row = [ordered]@{
                            'Group Name'                  = $group.DisplayName
                            'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                            'Member Display Name'         = $member.DisplayName
                            'Member FirstName'            = $member.FirstName
                            'Member LastName'             = $member.LastName
                            'Member Primary Smtp Address' = $member.PrimarySmtpAddress
                            'Member Company'              = $member.Company
                            'Member City'                 = $member.City
                        }
                    }
                    else {
                        $row = [ordered]@{
                            'Member Display Name'         = $member.DisplayName
                            'Member FirstName'            = $member.FirstName
                            'Member LastName'             = $member.LastName
                            'Member Primary Smtp Address' = $member.PrimarySmtpAddress
                            'Member Company'              = $member.Company
                            'Member City'                 = $member.City
                        }
                    }

                    $results.Add([pscustomobject]$row) | Out-Null
                }
            }

            if (-not $results -or $results.Count -eq 0) {
                Write-NCMessage "No members found for the specified dynamic distribution groups." -Level WARNING
                return
            }

            if ($GridView.IsPresent) {
                $results | Out-GridView -Title "M365 Dynamic Distribution Groups"
            }
            elseif ($emitCsv) {
                $csvPath = New-File("$($folder)\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-DynamicDistributionGroups-Report.csv")
                $results | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                Write-NCMessage "Dynamic distribution group membership exported to $csvPath." -Level SUCCESS
            }
            else {
                $results
            }
        }
        finally {
            Write-Progress -Activity "Processing dynamic distribution groups" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Get-DynamicDistributionGroupFilter {
    <#
    .SYNOPSIS
        Returns a simplified filter definition for a Dynamic Distribution Group.
    .DESCRIPTION
        Fetches the raw RecipientFilter, strips the boilerplate clauses automatically appended
        by Exchange Online, normalizes whitespace/quoting, and outputs a copyable expression that
        can be reused when editing the group filter.
    .PARAMETER DynamicDistributionGroup
        Dynamic distribution group identity (name, alias, or SMTP). Accepts pipeline input.
    .PARAMETER IncludeDefaults
        Include the default Exchange filter clauses (SystemMailbox, CAS_, Audit mailboxes, etc.).
    .PARAMETER AsObject
        Return metadata (name, container, clause list) instead of just the normalized filter string.
    .EXAMPLE
        Get-DynamicDistributionGroupFilter -DynamicDistributionGroup "Campus Assago"
    .EXAMPLE
        "campusassago@tenant.onmicrosoft.com" | Get-DynamicDistributionGroupFilter -AsObject
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('DDG', 'Identity')]
        [string[]]$DynamicDistributionGroup,
        [switch]$IncludeDefaults,
        [switch]$AsObject
    )

    begin {
        Set-ProgressAndInfoPreferences
        $requestedGroups = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($entry in $DynamicDistributionGroup) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$requestedGroups.Add($entry.Trim())
            }
        }
    }

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
                return
            }

            $targets = $requestedGroups | Select-Object -Unique
            if (-not $targets -or $targets.Count -eq 0) {
                Write-NCMessage "Specify at least one dynamic distribution group." -Level WARNING
                return
            }

            $defaultClauses = @(
                '-not Name -like "SystemMailbox{*"',
                '-not Name -like "CAS_{*"',
                '-not RecipientTypeDetailsValue -eq "MailboxPlan"',
                '-not RecipientTypeDetailsValue -eq "DiscoveryMailbox"',
                '-not RecipientTypeDetailsValue -eq "PublicFolderMailbox"',
                '-not RecipientTypeDetailsValue -eq "ArbitrationMailbox"',
                '-not RecipientTypeDetailsValue -eq "AuditLogMailbox"',
                '-not RecipientTypeDetailsValue -eq "AuxAuditLogMailbox"',
                '-not RecipientTypeDetailsValue -eq "SupervisoryReviewPolicyMailbox"'
            )

            $normalizeFilter = {
                param($filter)
                if ([string]::IsNullOrWhiteSpace($filter)) {
                    return $null
                }

                $value = $filter -replace "`r?`n", ' '
                $value = $value -replace '[()]', ' '
                $value = $value -replace '\s+-and\s+', ' -and '
                $value = $value -replace '\s+-or\s+', ' -or '
                $value = $value -replace '\s+-not\s+', ' -not '
                $value = $value -replace '\s+', ' '
                $value = $value.Replace("'", '"')
                return $value.Trim()
            }

            $splitClauses = {
                param($filter)
                if ([string]::IsNullOrWhiteSpace($filter)) {
                    return @()
                }

                return ($filter -split ' -and ' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }

            foreach ($identity in $targets) {
                try {
                    $group = Get-DynamicDistributionGroup -Identity $identity -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Dynamic distribution group '$identity' not found: $($_.Exception.Message)" -Level WARNING
                    continue
                }

                $rawFilter = $group.RecipientFilter
                if ([string]::IsNullOrWhiteSpace($rawFilter)) {
                    Write-NCMessage "Recipient filter not available for '$($group.DisplayName)'." -Level WARNING
                    continue
                }

                $normalizedFilter = & $normalizeFilter $rawFilter
                $clauses = & $splitClauses $normalizedFilter

                if (-not $IncludeDefaults.IsPresent -and $clauses.Count -gt 0) {
                    $clauses = $clauses | Where-Object { $defaultClauses -notcontains $_ }
                }

                if (-not $clauses -or $clauses.Count -eq 0) {
                    $clauses = & $splitClauses $normalizedFilter
                }

                $cleanedFilter = $clauses -join ' -and '

                $result = [pscustomobject][ordered]@{
                    Name               = $group.DisplayName
                    Identity           = $group.Identity
                    RecipientContainer = $group.RecipientContainer
                    Filter             = $cleanedFilter
                    RawFilter          = $rawFilter
                    FilterClauses      = $clauses
                    IncludeDefaults    = $IncludeDefaults.IsPresent
                }

                if ($AsObject.IsPresent) {
                    $result
                }
                else {
                    $cleanedFilter
                }
            }
        }
        finally {
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Export-M365Group {
    <#
    .SYNOPSIS
        Exports Microsoft 365 (Unified) group membership information.
    .DESCRIPTION
        Ensures an Exchange Online session, enumerates either all Microsoft 365 groups or the
        provided identities, and retrieves their members for CSV, GridView, or pipeline output.
    .PARAMETER M365Group
        Microsoft 365 group identity (name, alias, or SMTP). Accepts pipeline input.
    .PARAMETER Csv
        Export the results to CSV (implied when -All or -CsvFolder is specified).
    .PARAMETER CsvFolder
        Destination folder for the CSV report. Defaults to the current directory.
    .PARAMETER All
        Export every Microsoft 365 group in the tenant.
    .PARAMETER GridView
        Show the result in Out-GridView instead of returning objects.
    .EXAMPLE
        Export-M365Group -M365Group "Project Hub"
    .EXAMPLE
        Export-M365Group -All -CsvFolder C:\Temp
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Group', 'Identity')]
        [string[]]$M365Group,
        [switch]$Csv,
        [string]$CsvFolder,
        [switch]$All,
        [switch]$GridView
    )

    begin {
        Set-ProgressAndInfoPreferences
        $requestedGroups = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if ($M365Group) {
            foreach ($entry in $M365Group) {
                if (-not [string]::IsNullOrWhiteSpace($entry)) {
                    [void]$requestedGroups.Add($entry.Trim())
                }
            }
        }
    }

    end {
        try {
            if (-not (Test-EOLConnection)) {
                Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
                return
            }

            $exportAll = $All.IsPresent -or $requestedGroups.Count -eq 0
            $emitCsv = $Csv.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvFolder) -or $exportAll
            $folder = $null

            if ($emitCsv) {
                try {
                    $folder = Test-Folder $CsvFolder
                }
                catch {
                    Write-NCMessage "Invalid CSV folder. $($_.Exception.Message)" -Level ERROR
                    return
                }
            }

            $groups = @()
            if ($exportAll) {
                try {
                    $groups = Get-UnifiedGroup -ResultSize Unlimited -WarningAction SilentlyContinue
                }
                catch {
                    Write-NCMessage "Failed to retrieve Microsoft 365 groups: $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            else {
                foreach ($identity in ($requestedGroups.ToArray() | Select-Object -Unique)) {
                    try {
                        $groups += Get-UnifiedGroup -Identity $identity -ErrorAction Stop
                    }
                    catch {
                        Write-NCMessage "Microsoft 365 group '$identity' not found: $($_.Exception.Message)" -Level WARNING
                    }
                }
            }

            if (-not $groups -or $groups.Count -eq 0) {
                Write-NCMessage "No Microsoft 365 groups found matching the provided criteria." -Level WARNING
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            $counter = 0
            $total = $groups.Count

            foreach ($group in $groups) {
                $counter++
                $percentComplete = (($counter / $total) * 100)
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $total ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

                try {
                    $members = @(Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Member -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Failed to retrieve members for '$($group.DisplayName)': $($_.Exception.Message)" -Level WARNING
                    continue
                }

                if (-not $members -or $members.Count -eq 0) {
                    if ($exportAll) {
                        $results.Add([pscustomobject][ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                                'Member Display Name'         = $null
                                'Member FirstName'            = $null
                                'Member LastName'             = $null
                                'Member Primary Smtp Address' = $null
                                'Member Company'              = $null
                                'Member City'                 = $null
                            }) | Out-Null
                    }
                    continue
                }

                foreach ($member in $members) {
                    if ($exportAll) {
                        $row = [ordered]@{
                            'Group Name'                  = $group.DisplayName
                            'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                            'Member Display Name'         = $member.DisplayName
                            'Member FirstName'            = $member.FirstName
                            'Member LastName'             = $member.LastName
                            'Member Primary Smtp Address' = $member.PrimarySmtpAddress
                            'Member Company'              = $member.Company
                            'Member City'                 = $member.City
                        }
                    }
                    else {
                        $row = [ordered]@{
                            'Member Display Name'         = $member.DisplayName
                            'Member FirstName'            = $member.FirstName
                            'Member LastName'             = $member.LastName
                            'Member Primary Smtp Address' = $member.PrimarySmtpAddress
                            'Member Company'              = $member.Company
                            'Member City'                 = $member.City
                        }
                    }

                    $results.Add([pscustomobject]$row) | Out-Null
                }
            }

            if (-not $results -or $results.Count -eq 0) {
                Write-NCMessage "No members found for the specified Microsoft 365 groups." -Level WARNING
                return
            }

            if ($GridView.IsPresent) {
                $results | Out-GridView -Title "M365 Unified Groups"
            }
            elseif ($emitCsv) {
                $csvPath = New-File("$($folder)\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-UnifiedGroups-Report.csv")
                $results | Export-CSV -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $($NCVars.CSV_DefaultLimiter)
                Write-NCMessage "Microsoft 365 group membership exported to $csvPath." -Level SUCCESS
            }
            else {
                $results
            }
        }
        finally {
            Write-Progress -Activity "Processing Microsoft 365 groups" -Completed
            Restore-ProgressAndInfoPreferences
        }
    }
}

function Get-RoleGroupsMembers {
    <#
    .SYNOPSIS
        Lists Exchange Online role groups and their members.
    .DESCRIPTION
        Ensures an Exchange Online session, retrieves all role groups, collects members, and
        returns the data (optionally as a formatted table or GridView).
    .PARAMETER AsTable
        Display the output as a formatted table (default is to return objects).
    .PARAMETER GridView
        Show the result in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-RoleGroupsMembers
    .EXAMPLE
        Get-RoleGroupsMembers -AsTable
    #>
    [CmdletBinding()]
    param(
        [switch]$AsTable,
        [switch]$GridView
    )

    Set-ProgressAndInfoPreferences
    try {
        if (-not (Test-EOLConnection)) {
            Write-NCMessage "`nCan't connect or use Microsoft Exchange Online Management module. `nPlease check logs." -Level ERROR
            return
        }

        try {
            $roleGroups = Get-RoleGroup -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Failed to retrieve role groups: $($_.Exception.Message)" -Level ERROR
            return
        }

        if (-not $roleGroups -or $roleGroups.Count -eq 0) {
            Write-NCMessage "No role groups found." -Level WARNING
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $counter = 0
        $total = $roleGroups.Count

        foreach ($group in $roleGroups) {
            $counter++
            $percentComplete = (($counter / $total) * 100)
            Write-Progress -Activity "Processing $($group.Name)" -Status "$counter of $total ($($percentComplete.ToString('0.00'))%)" -PercentComplete $percentComplete

            try {
                $members = @(Get-RoleGroupMember -Identity $group.Identity -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Failed to retrieve members for role group '$($group.Name)': $($_.Exception.Message)" -Level WARNING
                continue
            }

            $results.Add([pscustomobject][ordered]@{
                    'Role Group' = $group.Name
                    Count        = if ($members) { $members.Count } else { 0 }
                    Members      = if ($members) { ($members.DisplayName -join "`n") } else { $null }
                }) | Out-Null
        }

        $sorted = $results | Sort-Object Count -Descending

        if ($GridView.IsPresent) {
            $sorted | Out-GridView -Title "Exchange Role Groups"
        }
        elseif ($AsTable.IsPresent) {
            Show-Table -Rows $sorted -AsTable
        }
        else {
            $sorted
        }
    }
    finally {
        Write-Progress -Activity "Processing role groups" -Completed
        Restore-ProgressAndInfoPreferences
    }
}

function Get-UserGroups {
    <#
    .SYNOPSIS
        Shows the Microsoft 365 groups that a user, contact, or distribution group belongs to.
    .DESCRIPTION
        Ensures Microsoft Graph (and Exchange Online) connectivity, resolves the provided identity
        via Get-Recipient, and uses Microsoft Graph to list every directory object membership.
    .PARAMETER UserPrincipalName
        User, contact, or group identity. Accepts display names, aliases, or e-mail addresses.
    .PARAMETER GridView
        Show additional details in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-UserGroups -UserPrincipalName user@contoso.com
    .EXAMPLE
        'user@contoso.com' | Get-UserGroups -GridView
    .NOTES
        Inspired by community samples originally published on infrasos.com.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'Identity', 'UPN')]
        [string]$UserPrincipalName,
        [switch]$GridView
    )

    begin {
        $graphConnected = $null
    }

    process {
        if ($null -eq $graphConnected) {
            $graphConnected = Test-MgGraphConnection
            if (-not $graphConnected) {
                Write-NCMessage "`nCan't connect or use Microsoft Graph modules. `nPlease check logs." -Level ERROR
                return
            }
        }

        $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $UserPrincipalName
        if (-not $resolvedPrincipal) {
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        }
        
        $recipientType = (Get-Recipient -Identity $resolvedPrincipal).RecipientTypeDetails
        $memberships = @()

        switch ($recipientType) {
            'MailContact' {
                try {
                    $contact = Get-MgContact -Filter "Mail eq '$resolvedPrincipal'" -All -ErrorAction Stop | Select-Object -First 1
                }
                catch {
                    Write-NCMessage "Unable to resolve contact $resolvedPrincipal in Microsoft Graph: $($_.Exception.Message)" -Level ERROR
                    return
                }

                if (-not $contact) {
                    Write-NCMessage "Microsoft Graph contact not found for $resolvedPrincipal." -Level WARNING
                    return
                }

                try {
                    $memberships = @(Get-MgContactMemberOf -OrgContactId $contact.Id -All -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Unable to read group memberships for contact ${resolvedPrincipal}: $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            'MailUniversalDistributionGroup' {
                try {
                    $group = Get-MgGroup -Filter "Mail eq '$resolvedPrincipal'" -All -ErrorAction Stop | Select-Object -First 1
                }
                catch {
                    Write-NCMessage "Unable to resolve group $resolvedPrincipal in Microsoft Graph: $($_.Exception.Message)" -Level ERROR
                    return
                }

                if (-not $group) {
                    Write-NCMessage "Microsoft Graph group not found for $resolvedPrincipal." -Level WARNING
                    return
                }

                try {
                    $memberships = @(Get-MgGroupMemberOf -GroupId $group.Id -All -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Unable to read memberships for group ${resolvedPrincipal}: $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
            default {
                $recipient = Get-Mailbox -Identity $resolvedPrincipal -ErrorAction Stop # To get WindowsLiveID when UPN differs / when Get-Recipient can't provide it
                $userId = if ($recipient.WindowsLiveID) { $recipient.WindowsLiveID } else { $resolvedPrincipal }

                try {
                    $user = Get-MgUser -UserId $userId -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Unable to resolve user $userId in Microsoft Graph: $($_.Exception.Message)" -Level ERROR
                    return
                }

                try {
                    $memberships = @(Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Unable to read group memberships for ${resolvedPrincipal}: $($_.Exception.Message)" -Level ERROR
                    return
                }
            }
        }

        Write-NCMessage "`n$recipientType ($resolvedPrincipal) - Groups found: $($memberships.Count)" -Level VERBOSE

        if (-not $memberships -or $memberships.Count -eq 0) {
            Write-NCMessage "No groups found for $resolvedPrincipal." -Level WARNING
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($membership in $memberships) {
            $props = if ($membership.AdditionalProperties) { $membership.AdditionalProperties } else { @{} }
            $row = [ordered]@{
                'Group Name' = if ($props.ContainsKey('displayName')) { $props.displayName } else { $null }
                'Group Mail' = if ($props.ContainsKey('mail')) { $props.mail } else { $null }
            }

            if ($GridView.IsPresent) {
                $row['Group Description'] = if ($props.ContainsKey('description')) { $props.description } else { $null }
                $row['Group Mail Nickname'] = if ($props.ContainsKey('mailNickname')) { $props.mailNickname } else { $null }
                $row['Group Mail Enabled'] = if ($props.ContainsKey('mailEnabled')) { $props.mailEnabled } else { $null }
                $row['Group Type'] = if ($props.ContainsKey('groupTypes')) { ($props.groupTypes -join ', ') } else { $null }
                $row['Group ID'] = $membership.Id
            }

            $results.Add([pscustomobject]$row) | Out-Null
        }

        if ($GridView.IsPresent) {
            $results | Out-GridView -Title "M365 User Groups - $resolvedPrincipal"
        }
        else {
            $results | Sort-Object 'Group Name'
        }
    }
}

Set-Alias -Name Export-DG -Value Export-DistributionGroups
Set-Alias -Name Export-DDG -Value Export-DynamicDistributionGroups
Set-Alias -Name Get-DDGRecipientFilter -Value Get-DynamicDistributionGroupFilter
