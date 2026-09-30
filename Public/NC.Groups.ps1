#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Groups helpers =======================================================================================================================

function Add-EntraGroupDevice {
    <#
    .SYNOPSIS
        Adds one or more devices to an Entra group.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then adds
        the provided devices by display name or object ID. Accepts pipeline input for devices.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER DeviceIdentifier
        Device display name or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER TreatInputAsId
        Treat every DeviceIdentifier as an object ID without attempting name resolution.
    .PARAMETER PassThru
        Emit a summary object for each processed device.
    .EXAMPLE
        "PC1", "PC2" | Add-EntraGroupDevice -GroupName "My Entra Group"
    .EXAMPLE
        Add-EntraGroupDevice -GroupId "00000000-0000-0000-0000-000000000000" -DeviceIdentifier "PC1"
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Alias('Group', 'DisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [string]$GroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Device', 'DeviceName', 'Id', 'DeviceId', 'Name')]
        [string[]]$DeviceIdentifier,

        [switch]$TreatInputAsId,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $devices = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not $graphConnected) { return }

        foreach ($entry in $DeviceIdentifier) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$devices.Add($entry.Trim())
            }
        }
    }

    end {
        if (-not $graphConnected) { return }
        if ($devices.Count -eq 0) {
            Write-NCMessage "No devices were specified." -Level WARNING
            return
        }

        $resolvedGroup = $null
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                $resolvedGroup = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Entra group with ID '$GroupId' not found: $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            $escapedName = $GroupName.Replace("'", "''")
            try {
                $resolvedGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-NCMessage "Unable to resolve group '$GroupName': $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $resolvedGroup) {
                try {
                    $resolvedGroup = Get-MgGroup -GroupId $GroupName -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Entra group '$GroupName' not found by name or ID" -Level ERROR
                    return
                }
            }
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $uniqueDevices = @($devices | Select-Object -Unique)
        $membersRefUrl = "/groups/$([uri]::EscapeDataString([string]$resolvedGroup.Id))/members/`$ref"
        Write-NCMessage "Processing $($uniqueDevices.Count) device(s) in Graph batches (20 per request) ..." -Level INFO

        for ($offset = 0; $offset -lt $uniqueDevices.Count; $offset += 20) {
            $chunk = @($uniqueDevices[$offset..([Math]::Min($offset + 20, $uniqueDevices.Count) - 1)])
            $targets = @(Resolve-NCEntraDeviceTargetBatch -DeviceIdentifier $chunk -TreatInputAsId:$TreatInputAsId)

            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($target in $targets) {
                if ($PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Add device '$($target.Label)'")) {
                    $approved.Add($target)
                }
            }
            if ($approved.Count -eq 0) { continue }

            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'POST'; Url = $membersRefUrl; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $approved[$i].Id) } }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Adding devices to $($resolvedGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $deviceId = $approved[$i].Id
                $deviceLabel = $approved[$i].Label
                $response = $responses[$i]
                $status = 'Added'

                if ($response.Success) {
                    Write-NCMessage "Added device '$deviceLabel' to group '$($resolvedGroup.DisplayName)'" -Level SUCCESS
                }
                elseif ($response.ErrorMessage -match 'added object references already exist') {
                    $status = 'Exists'
                    Write-NCMessage "Device '$deviceLabel' is already a member of '$($resolvedGroup.DisplayName)'" -Level WARNING
                }
                else {
                    $status = 'Failed'
                    Write-NCMessage "Failed to add device '$deviceLabel' to '$($resolvedGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }

                if ($PassThru.IsPresent) {
                    $results.Add([pscustomobject][ordered]@{
                            GroupName  = $resolvedGroup.DisplayName
                            GroupId    = $resolvedGroup.Id
                            MemberName = $deviceLabel
                            MemberId   = $deviceId
                            MemberType = 'Device'
                            Status     = $status
                        }) | Out-Null
                }
            }
        }

        if ($PassThru.IsPresent -and $results.Count -gt 0) {
            $results
        }
    }
}

function Add-EntraGroupOwner {
    <#
    .SYNOPSIS
        Adds one or more owners to an Entra group.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then adds
        the provided owners by UPN/display name or object ID. Accepts pipeline input for owners.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER OwnerIdentifier
        User principal name, display name, or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER TreatInputAsId
        Treat every OwnerIdentifier as an object ID without attempting name resolution.
    .PARAMETER PassThru
        Emit a summary object for each processed owner.
    .EXAMPLE
        "user1@contoso.com","user2@contoso.com" | Add-EntraGroupOwner -GroupName "My Entra Group"
    .EXAMPLE
        Add-EntraGroupOwner -GroupId "00000000-0000-0000-0000-000000000000" -OwnerIdentifier "user1@contoso.com"
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Alias('Group', 'DisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [string]$GroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Owner', 'UPN', 'Mail', 'Id', 'UserId', 'Name')]
        [string[]]$OwnerIdentifier,

        [switch]$TreatInputAsId,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $owners = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not $graphConnected) { return }

        foreach ($entry in $OwnerIdentifier) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$owners.Add($entry.Trim())
            }
        }
    }

    end {
        if (-not $graphConnected) { return }
        if ($owners.Count -eq 0) {
            Write-NCMessage "No owners were specified." -Level WARNING
            return
        }

        $resolvedGroup = Resolve-NCEntraGroup -GroupName $GroupName -GroupId $GroupId
        if (-not $resolvedGroup) {
            return
        }

        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $uniqueOwners = @($owners | Select-Object -Unique)
        $ownersRefUrl = "/groups/$([uri]::EscapeDataString([string]$resolvedGroup.Id))/owners/`$ref"
        Write-NCMessage "Processing $($uniqueOwners.Count) owner(s) in Graph batches (20 per request) ..." -Level INFO

        for ($offset = 0; $offset -lt $uniqueOwners.Count; $offset += 20) {
            $chunk = @($uniqueOwners[$offset..([Math]::Min($offset + 20, $uniqueOwners.Count) - 1)])
            $targets = @(Resolve-NCEntraOwnerBatch -OwnerIdentifier $chunk -TreatInputAsId:$TreatInputAsId)

            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($target in $targets) {
                if ($PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Add owner '$($target.Label)'")) {
                    $approved.Add($target)
                }
            }
            if ($approved.Count -eq 0) { continue }

            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'POST'; Url = $ownersRefUrl; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $approved[$i].Id) } }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Adding owners to $($resolvedGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $resolvedOwner = $approved[$i]
                $response = $responses[$i]
                $status = 'Added'

                if ($response.Success) {
                    Write-NCMessage "Added owner '$($resolvedOwner.Label)' to group '$($resolvedGroup.DisplayName)'." -Level SUCCESS
                }
                elseif ($response.ErrorMessage -match 'already exist' -or $response.ErrorMessage -match 'exists') {
                    $status = 'Exists'
                    Write-NCMessage "Owner '$($resolvedOwner.Label)' is already an owner of '$($resolvedGroup.DisplayName)'." -Level WARNING
                }
                else {
                    $status = 'Failed'
                    Write-NCMessage "Failed to add owner '$($resolvedOwner.Label)' to '$($resolvedGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }

                if ($PassThru.IsPresent) {
                    $results.Add([pscustomobject][ordered]@{
                            GroupName = $resolvedGroup.DisplayName
                            GroupId   = $resolvedGroup.Id
                            OwnerName = $resolvedOwner.Label
                            OwnerId   = $resolvedOwner.Id
                            Status    = $status
                        }) | Out-Null
                }
            }
        }

        if ($PassThru.IsPresent -and $results.Count -gt 0) {
            $results
        }
    }
}

function Remove-EntraGroupOwner {
    <#
    .SYNOPSIS
        Removes one or more owners from an Entra group.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then removes
        the provided owners by UPN/display name or object ID. Accepts pipeline input for owners.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER OwnerIdentifier
        User principal name, display name, or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER ClearAll
        Remove all owners from the Entra group.
    .PARAMETER TreatInputAsId
        Treat every OwnerIdentifier as an object ID without attempting name resolution.
    .PARAMETER PassThru
        Emit a summary object for each processed owner.
    .EXAMPLE
        "user1@contoso.com","user2@contoso.com" | Remove-EntraGroupOwner -GroupName "My Entra Group"
    .EXAMPLE
        Remove-EntraGroupOwner -GroupId "00000000-0000-0000-0000-000000000000" -OwnerIdentifier "user1@contoso.com"
    .EXAMPLE
        Remove-EntraGroupOwner -GroupName "My Entra Group" -ClearAll
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllByName')]
        [Alias('Group', 'DisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllById')]
        [string]$GroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Owner', 'UPN', 'Mail', 'Id', 'UserId', 'Name')]
        [string[]]$OwnerIdentifier,

        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllByName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllById')]
        [switch]$ClearAll,

        [switch]$TreatInputAsId,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $owners = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not $graphConnected) { return }
        if ($ClearAll.IsPresent) { return }

        foreach ($entry in $OwnerIdentifier) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$owners.Add($entry.Trim())
            }
        }
    }

    end {
        if (-not $graphConnected) { return }
        if (-not $ClearAll.IsPresent -and $owners.Count -eq 0) {
            Write-NCMessage "No owners were specified." -Level WARNING
            return
        }

        $resolvedGroup = Resolve-NCEntraGroup -GroupName $GroupName -GroupId $GroupId
        if (-not $resolvedGroup) {
            return
        }

        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $ownersToRemove = [System.Collections.Generic.List[object]]::new()

        if ($ClearAll.IsPresent) {
            $confirmMessage = "You are about to remove ALL owners from '$($resolvedGroup.DisplayName)'. This is a high-risk operation."
            if (-not $PSCmdlet.ShouldContinue($confirmMessage, "Confirm ClearAll")) {
                Write-NCMessage "ClearAll operation cancelled." -Level WARNING
                return
            }

            try {
                $ownerResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($resolvedGroup.Id)/owners?`$select=id,displayName,userPrincipalName,appDisplayName" -Method GET
                $ownerItems = @()
                if ($ownerResponse -and $ownerResponse.PSObject.Properties['value']) {
                    $ownerItems = @($ownerResponse.value)
                }
                elseif ($ownerResponse) {
                    $ownerItems = @($ownerResponse)
                }
            }
            catch {
                Write-NCMessage "Unable to read owners for group $($resolvedGroup.DisplayName): $($_.Exception.Message)" -Level ERROR
                return
            }

            foreach ($ownerItem in $ownerItems) {
                $ownerId = if ($ownerItem.PSObject.Properties['id']) { [string]$ownerItem.id } else { $null }
                if ([string]::IsNullOrWhiteSpace($ownerId)) {
                    continue
                }

                $ownersToRemove.Add([pscustomobject]@{
                        Id    = $ownerId
                        Label = (Get-NCGraphObjectLabel -InputObject $ownerItem)
                    }) | Out-Null
            }

            if ($ownersToRemove.Count -eq 0) {
                Write-NCMessage "No owners found for $($resolvedGroup.DisplayName)." -Level WARNING
                return
            }
        }
        else {
            $uniqueOwners = @($owners | Select-Object -Unique)

            foreach ($resolvedOwner in @(Resolve-NCEntraOwnerBatch -OwnerIdentifier $uniqueOwners -TreatInputAsId:$TreatInputAsId)) {
                $ownersToRemove.Add([pscustomobject]@{
                        Id    = $resolvedOwner.Id
                        Label = $resolvedOwner.Label
                    }) | Out-Null
            }
        }

        if ($ownersToRemove.Count -gt 0) {
            Write-NCMessage "Processing $($ownersToRemove.Count) owner(s) in Graph batches (20 per request) ..." -Level INFO
        }

        for ($offset = 0; $offset -lt $ownersToRemove.Count; $offset += 20) {
            $chunk = @($ownersToRemove[$offset..([Math]::Min($offset + 20, $ownersToRemove.Count) - 1)])

            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $chunk) {
                if ($PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Remove owner '$($entry.Label)'")) {
                    $approved.Add($entry)
                }
            }
            if ($approved.Count -eq 0) { continue }

            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'DELETE'; Url = "/groups/$([uri]::EscapeDataString([string]$resolvedGroup.Id))/owners/$([uri]::EscapeDataString([string]$approved[$i].Id))/`$ref" }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Removing owners from $($resolvedGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $entry = $approved[$i]
                $response = $responses[$i]
                $status = 'Removed'

                if ($response.Success) {
                    Write-NCMessage "Removed owner '$($entry.Label)' from group '$($resolvedGroup.DisplayName)'." -Level SUCCESS
                }
                elseif ($response.ErrorMessage -match 'could not find' -or $response.ErrorMessage -match 'does not exist') {
                    $status = 'NotFound'
                    Write-NCMessage "Owner '$($entry.Label)' is not an owner of '$($resolvedGroup.DisplayName)'." -Level WARNING
                }
                else {
                    $status = 'Failed'
                    Write-NCMessage "Failed to remove owner '$($entry.Label)' from '$($resolvedGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }

                if ($PassThru.IsPresent) {
                    $results.Add([pscustomobject][ordered]@{
                            GroupName = $resolvedGroup.DisplayName
                            GroupId   = $resolvedGroup.Id
                            OwnerName = $entry.Label
                            OwnerId   = $entry.Id
                            Status    = $status
                        }) | Out-Null
                }
            }
        }

        if ($PassThru.IsPresent -and $results.Count -gt 0) {
            $results
        }
    }
}

function Copy-EntraGroupOwner {
    <#
    .SYNOPSIS
        Copies owners from one Entra group to another.
    .DESCRIPTION
        Reads the owners of a source Entra group and adds any missing owners to the destination
        group without removing existing destination owners.
    .PARAMETER SourceGroupName
        Display name of the source Entra group.
    .PARAMETER SourceGroupId
        Object ID of the source Entra group.
    .PARAMETER DestinationGroupName
        Display name of the destination Entra group.
    .PARAMETER DestinationGroupId
        Object ID of the destination Entra group.
    .PARAMETER PassThru
        Emit a summary object for each copied owner.
    .EXAMPLE
        Copy-EntraGroupOwner -SourceGroupName "HR" -DestinationGroupName "HR - Test"
    .EXAMPLE
        Copy-EntraGroupOwner -SourceGroupId "00000000-0000-0000-0000-000000000000" -DestinationGroupId "11111111-1111-1111-1111-111111111111"
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'ByName')]
        [Alias('Source', 'From')]
        [string]$SourceGroupName,

        [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'ById')]
        [string]$SourceGroupId,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'ByName')]
        [Alias('Destination', 'To')]
        [string]$DestinationGroupName,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'ById')]
        [string]$DestinationGroupId,

        [switch]$PassThru
    )

    $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
    if (-not $graphConnected) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return
    }

    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
        return
    }

    $sourceGroup = if ($PSCmdlet.ParameterSetName -eq 'ById') {
        Resolve-NCEntraGroup -GroupName $SourceGroupId -GroupId $SourceGroupId
    }
    else {
        Resolve-NCEntraGroup -GroupName $SourceGroupName
    }

    if (-not $sourceGroup) {
        return
    }

    $destinationGroup = if ($PSCmdlet.ParameterSetName -eq 'ById') {
        Resolve-NCEntraGroup -GroupName $DestinationGroupId -GroupId $DestinationGroupId
    }
    else {
        Resolve-NCEntraGroup -GroupName $DestinationGroupName
    }

    if (-not $destinationGroup) {
        return
    }

    if ($sourceGroup.Id -eq $destinationGroup.Id) {
        Write-NCMessage "Source and destination groups are the same. Aborting." -Level ERROR
        return
    }

    try {
        $ownerResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($sourceGroup.Id)/owners?`$select=id,displayName,userPrincipalName,appDisplayName" -Method GET
        $sourceOwners = @()
        if ($ownerResponse -and $ownerResponse.PSObject.Properties['value']) {
            $sourceOwners = @($ownerResponse.value)
        }
        elseif ($ownerResponse) {
            $sourceOwners = @($ownerResponse)
        }
    }
    catch {
        Write-NCMessage "Unable to read owners for source group $($sourceGroup.DisplayName): $($_.Exception.Message)" -Level ERROR
        return
    }

    if ($sourceOwners.Count -eq 0) {
        Write-NCMessage "Source group $($sourceGroup.DisplayName) has no owners to copy." -Level WARNING
        return
    }

    try {
        $destinationOwnerResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($destinationGroup.Id)/owners?`$select=id" -Method GET
        $destinationOwners = @()
        if ($destinationOwnerResponse -and $destinationOwnerResponse.PSObject.Properties['value']) {
            $destinationOwners = @($destinationOwnerResponse.value)
        }
        elseif ($destinationOwnerResponse) {
            $destinationOwners = @($destinationOwnerResponse)
        }
    }
    catch {
        Write-NCMessage "Unable to read owners for destination group $($destinationGroup.DisplayName): $($_.Exception.Message)" -Level ERROR
        return
    }

    $destinationOwnerIds = @($destinationOwners | ForEach-Object { [string]$_.id })
    $results = [System.Collections.Generic.List[object]]::new()
    $entries = [System.Collections.Generic.List[object]]::new()

    foreach ($ownerItem in $sourceOwners) {
        $ownerId = if ($ownerItem.PSObject.Properties['id']) { [string]$ownerItem.id } else { $null }
        if ([string]::IsNullOrWhiteSpace($ownerId)) {
            continue
        }

        $entries.Add([pscustomobject]@{
                Id     = $ownerId
                Label  = (Get-NCGraphObjectLabel -InputObject $ownerItem)
                Status = if ($destinationOwnerIds -contains $ownerId) { 'Exists' } else { $null }
                Send   = $false
            })
    }

    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $entries) {
        if ($entry.Status -eq 'Exists') { continue }
        if ($PSCmdlet.ShouldProcess($destinationGroup.DisplayName, "Copy owner '$($entry.Label)' from '$($sourceGroup.DisplayName)'")) {
            $entry.Send = $true
            $pending.Add($entry)
        }
    }

    if ($pending.Count -gt 0) {
        $ownersRefUrl = "/groups/$([uri]::EscapeDataString([string]$destinationGroup.Id))/owners/`$ref"
        Write-NCMessage "Processing $($pending.Count) owner(s) in Graph batches (20 per request) ..." -Level INFO

        for ($offset = 0; $offset -lt $pending.Count; $offset += 20) {
            $approved = @($pending[$offset..([Math]::Min($offset + 20, $pending.Count) - 1)])
            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'POST'; Url = $ownersRefUrl; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $approved[$i].Id) } }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Copying owners to $($destinationGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $ownerLabel = $approved[$i].Label
                $response = $responses[$i]

                if ($response.Success) {
                    $approved[$i].Status = 'Added'
                    Write-NCMessage "Copied owner '$ownerLabel' to '$($destinationGroup.DisplayName)'." -Level SUCCESS
                }
                elseif ($response.ErrorMessage -match 'already exist' -or $response.ErrorMessage -match 'exists') {
                    $approved[$i].Status = 'Exists'
                    Write-NCMessage "Owner '$ownerLabel' is already an owner of '$($destinationGroup.DisplayName)'." -Level WARNING
                }
                else {
                    $approved[$i].Status = 'Failed'
                    Write-NCMessage "Failed to copy owner '$ownerLabel' to '$($destinationGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }
            }
        }
    }

    if ($PassThru.IsPresent) {
        foreach ($entry in $entries) {
            if ($entry.Status -eq 'Exists' -or $entry.Send) {
                $results.Add([pscustomobject][ordered]@{
                        SourceGroup      = $sourceGroup.DisplayName
                        DestinationGroup = $destinationGroup.DisplayName
                        OwnerName        = $entry.Label
                        OwnerId          = $entry.Id
                        Status           = $entry.Status
                    }) | Out-Null
            }
        }
    }

    if ($PassThru.IsPresent -and $results.Count -gt 0) {
        $results
    }
}

function Copy-EntraGroup {
    <#
    .SYNOPSIS
        Clones an Entra group into a new or existing group.
    .DESCRIPTION
        Copies the source group's core properties, owners, and members into the destination.
        If the destination group name does not already exist, a new group is created first.
        Dynamic membership groups are cloned as a static snapshot of their current members.
    .PARAMETER SourceGroupName
        Display name of the source Entra group.
    .PARAMETER SourceGroupId
        Object ID of the source Entra group.
    .PARAMETER DestinationGroupName
        Display name of the destination Entra group. If no group with this name exists, a new
        group is created with this name.
    .PARAMETER DestinationGroupId
        Object ID of an existing destination Entra group.
    .PARAMETER SkipMembers
        Do not copy members.
    .PARAMETER SkipOwners
        Do not copy owners.
    .PARAMETER SkipDescription
        Do not copy the source description.
    .PARAMETER PassThru
        Emit a summary object for the clone operation.
    .EXAMPLE
        Copy-EntraGroup -SourceGroupName "GitLab-Prod" -DestinationGroupName "GitLab-Prod-Test"
    .EXAMPLE
        Copy-EntraGroup -SourceGroupName "GitLab-Prod" -DestinationGroupName "GitLab-Prod-Test" -SkipOwners
    .EXAMPLE
        Copy-EntraGroup -SourceGroupId "00000000-0000-0000-0000-000000000000" -DestinationGroupId "11111111-1111-1111-1111-111111111111" -PassThru
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Alias('Source', 'From')]
        [string]$SourceGroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [string]$SourceGroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1)]
        [Alias('Destination', 'To')]
        [string]$DestinationGroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1)]
        [string]$DestinationGroupId,

        [switch]$SkipMembers,
        [switch]$SkipOwners,
        [switch]$SkipDescription,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            return
        }

        if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
            return
        }
    }

    process {
        $sourceResolved = if ($PSCmdlet.ParameterSetName -eq 'ById') {
            Resolve-NCEntraGroup -GroupName $SourceGroupId -GroupId $SourceGroupId
        }
        else {
            Resolve-NCEntraGroup -GroupName $SourceGroupName
        }

        if (-not $sourceResolved) {
            return
        }

        try {
            $sourceGroup = Get-MgGroup -GroupId $sourceResolved.Id -Property @(
                    'id',
                    'displayName',
                    'description',
                    'groupTypes',
                    'mailEnabled',
                    'mailNickname',
                    'securityEnabled',
                    'visibility',
                    'onPremisesSyncEnabled',
                    'isAssignableToRole'
                ) -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to retrieve full details for source group '$($sourceResolved.DisplayName)': $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($sourceGroup.IsAssignableToRole -eq $true) {
            Write-NCMessage "Role-assignable groups are not supported by Copy-EntraGroup yet." -Level ERROR
            return
        }

        if (($sourceGroup.MailEnabled -eq $true) -and ($sourceGroup.SecurityEnabled -eq $false) -and (-not $sourceGroup.GroupTypes -or ($sourceGroup.GroupTypes.Count -eq 0))) {
            Write-NCMessage "Distribution groups are not supported by Copy-EntraGroup. Use the distribution-group helpers instead." -Level ERROR
            return
        }

        $destinationCreated = $false
        $destinationGroup = $null

        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                $destinationGroup = Get-MgGroup -GroupId $DestinationGroupId -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Destination Entra group with ID '$DestinationGroupId' not found: $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            $escapedName = $DestinationGroupName.Replace("'", "''")
            try {
                $destinationGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-NCMessage "Unable to resolve destination group '$DestinationGroupName': $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $destinationGroup) {
                $destinationCreated = $true
            }
        }

        if ($destinationGroup -and $destinationGroup.Id -eq $sourceGroup.Id) {
            Write-NCMessage "Source and destination groups are the same. Aborting." -Level ERROR
            return
        }

        $destinationDisplayName = if ($destinationCreated) { $DestinationGroupName } else { $destinationGroup.DisplayName }
        $operationLabel = if ($destinationCreated) {
            "Create and clone group '$($sourceGroup.DisplayName)' into '$DestinationGroupName'"
        }
        else {
            "Clone group '$($sourceGroup.DisplayName)' into existing group '$($destinationDisplayName)'"
        }

        if (-not $PSCmdlet.ShouldProcess($destinationDisplayName, $operationLabel)) {
            return
        }

        if ($destinationCreated) {
            $createGroupTypes = @()
            if ($sourceGroup.GroupTypes) {
                $createGroupTypes = @($sourceGroup.GroupTypes | Where-Object { $_ -ne 'DynamicMembership' })
                if ($sourceGroup.GroupTypes -contains 'DynamicMembership') {
                    Write-NCMessage "Source group '$($sourceGroup.DisplayName)' is dynamic; the cloned group will be a static snapshot of its current members." -Level WARNING
                }
            }

            $mailNickname = [regex]::Replace($DestinationGroupName, '[^a-zA-Z0-9]', '')
            if ([string]::IsNullOrWhiteSpace($mailNickname)) {
                $mailNickname = "group$((Get-Date).ToString('yyyyMMddHHmmss'))"
            }

            $createBody = [ordered]@{
                displayName     = $DestinationGroupName
                mailEnabled     = [bool]$sourceGroup.MailEnabled
                mailNickname    = $mailNickname
                securityEnabled = [bool]$sourceGroup.SecurityEnabled
            }

            if (-not $SkipDescription.IsPresent -and -not [string]::IsNullOrWhiteSpace($sourceGroup.Description)) {
                $createBody.description = $sourceGroup.Description
            }

            if ($createGroupTypes.Count -gt 0) {
                $createBody.groupTypes = $createGroupTypes
            }

            if (-not [string]::IsNullOrWhiteSpace($sourceGroup.Visibility)) {
                $createBody.visibility = $sourceGroup.Visibility
            }

            try {
                $destinationGroup = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/groups' -Method POST -Body ($createBody | ConvertTo-Json -Depth 10) -ContentType 'application/json'
                $destinationCreated = $true
                Write-NCMessage "Created destination group '$DestinationGroupName' for clone operation." -Level SUCCESS
            }
            catch {
                Write-NCMessage "Failed to create destination group '$DestinationGroupName': $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            if (-not $SkipDescription.IsPresent -and -not [string]::IsNullOrWhiteSpace($sourceGroup.Description) -and $sourceGroup.Description -ne $destinationGroup.Description) {
                try {
                    Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($destinationGroup.Id)" -Method PATCH -Body (@{ description = $sourceGroup.Description } | ConvertTo-Json -Depth 10) -ContentType 'application/json' | Out-Null
                    Write-NCMessage "Copied description to '$($destinationGroup.DisplayName)'." -Level SUCCESS
                }
                catch {
                    Write-NCMessage "Unable to copy description to '$($destinationGroup.DisplayName)': $($_.Exception.Message)" -Level WARNING
                }
            }
        }

        if ($destinationGroup.OnPremisesSyncEnabled -eq $true) {
            Write-NCMessage "Destination group '$($destinationGroup.DisplayName)' is synchronized from on-premises AD and cannot be modified directly in Entra." -Level ERROR
            return
        }

        try {
            $destinationMembers = @(Get-MgGroupMember -GroupId $destinationGroup.Id -All -ErrorAction Stop)
        }
        catch {
            Write-NCMessage "Unable to read destination members for '$($destinationGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
            return
        }

        $destinationMemberIds = @($destinationMembers | ForEach-Object { [string]$_.Id })
        $destinationOwnerIds = @()

        if (-not $SkipOwners.IsPresent) {
            try {
                $destinationOwnerUri = "https://graph.microsoft.com/v1.0/groups/$($destinationGroup.Id)/owners?`$select=id,displayName,userPrincipalName,appDisplayName"
                $destinationOwnerItems = @(Invoke-NCGraphAllPagesCore -Uri $destinationOwnerUri)
                $destinationOwnerIds = @($destinationOwnerItems | ForEach-Object { [string]$_.id })
            }
            catch {
                Write-NCMessage "Unable to read destination owners for '$($destinationGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
                return
            }
        }

        $memberCopied = 0
        $memberSkipped = 0
        $ownerCopied = 0
        $ownerSkipped = 0

        if (-not $SkipOwners.IsPresent) {
            try {
                $sourceOwnerUri = "https://graph.microsoft.com/v1.0/groups/$($sourceGroup.Id)/owners?`$select=id,displayName,userPrincipalName,appDisplayName"
                $sourceOwners = @(Invoke-NCGraphAllPagesCore -Uri $sourceOwnerUri)
            }
            catch {
                Write-NCMessage "Unable to read source owners for '$($sourceGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
                return
            }

            $ownerPending = [System.Collections.Generic.List[object]]::new()
            foreach ($owner in $sourceOwners) {
                $ownerId = if ($owner.PSObject.Properties['id']) { [string]$owner.id } else { $null }
                if ([string]::IsNullOrWhiteSpace($ownerId)) {
                    continue
                }

                if ($destinationOwnerIds -contains $ownerId) {
                    $ownerSkipped++
                    continue
                }

                $ownerPending.Add([pscustomobject]@{ Id = $ownerId; Label = (Get-NCGraphObjectLabel -InputObject $owner) })
            }

            if ($ownerPending.Count -gt 0) {
                $ownersRefUrl = "/groups/$([uri]::EscapeDataString([string]$destinationGroup.Id))/owners/`$ref"
                Write-NCMessage "Processing $($ownerPending.Count) owner(s) in Graph batches (20 per request) ..." -Level INFO

                for ($offset = 0; $offset -lt $ownerPending.Count; $offset += 20) {
                    $approved = @($ownerPending[$offset..([Math]::Min($offset + 20, $ownerPending.Count) - 1)])
                    $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                            @{ Id = "m$i"; Method = 'POST'; Url = $ownersRefUrl; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $approved[$i].Id) } }
                        })
                    $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Copying owners to $($destinationGroup.DisplayName)")

                    for ($i = 0; $i -lt $approved.Count; $i++) {
                        $ownerLabel = $approved[$i].Label
                        $response = $responses[$i]

                        if ($response.Success) {
                            $ownerCopied++
                            Write-NCMessage "Copied owner '$ownerLabel' to '$($destinationGroup.DisplayName)'." -Level SUCCESS
                        }
                        elseif ($response.ErrorMessage -match 'already exist' -or $response.ErrorMessage -match 'exists') {
                            $ownerSkipped++
                            Write-NCMessage "Owner '$ownerLabel' is already an owner of '$($destinationGroup.DisplayName)'." -Level WARNING
                        }
                        else {
                            Write-NCMessage "Failed to copy owner '$ownerLabel' to '$($destinationGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                        }
                    }
                }
            }
        }

        if (-not $SkipMembers.IsPresent) {
            try {
                $sourceMembers = @(Get-MgGroupMember -GroupId $sourceGroup.Id -All -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to read source members for '$($sourceGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
                return
            }

            $resolveType = {
                param($odataType)
                if ([string]::IsNullOrWhiteSpace($odataType)) {
                    return 'DirectoryObject'
                }

                $value = $odataType.ToLowerInvariant()
                if ($value -match 'user') { return 'User' }
                if ($value -match 'device') { return 'Device' }
                if ($value -match 'group') { return 'Group' }
                if ($value -match 'serviceprincipal') { return 'ServicePrincipal' }
                if ($value -match 'orgcontact') { return 'Contact' }
                return 'DirectoryObject'
            }

            $memberPending = [System.Collections.Generic.List[object]]::new()
            foreach ($member in $sourceMembers) {
                $memberId = if ($member.PSObject.Properties['id']) { [string]$member.id } else { $null }
                if ([string]::IsNullOrWhiteSpace($memberId)) {
                    continue
                }

                if ($destinationMemberIds -contains $memberId) {
                    $memberSkipped++
                    continue
                }

                $memberProps = if ($member.AdditionalProperties) { $member.AdditionalProperties } else { @{} }
                $memberType = if ($memberProps.ContainsKey('@odata.type')) { & $resolveType $memberProps['@odata.type'] } else { 'DirectoryObject' }
                $memberLabel = if ($memberProps.ContainsKey('displayName')) { $memberProps.displayName } elseif ($memberProps.ContainsKey('userPrincipalName')) { $memberProps.userPrincipalName } else { $memberId }

                $memberPending.Add([pscustomobject]@{ Id = $memberId; Type = $memberType; Label = $memberLabel })
            }

            if ($memberPending.Count -gt 0) {
                $membersRefUrl = "/groups/$([uri]::EscapeDataString([string]$destinationGroup.Id))/members/`$ref"
                Write-NCMessage "Processing $($memberPending.Count) member(s) in Graph batches (20 per request) ..." -Level INFO

                for ($offset = 0; $offset -lt $memberPending.Count; $offset += 20) {
                    $approved = @($memberPending[$offset..([Math]::Min($offset + 20, $memberPending.Count) - 1)])
                    $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                            @{ Id = "m$i"; Method = 'POST'; Url = $membersRefUrl; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $approved[$i].Id) } }
                        })
                    $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Copying members to $($destinationGroup.DisplayName)")

                    for ($i = 0; $i -lt $approved.Count; $i++) {
                        $memberType = $approved[$i].Type
                        $memberLabel = $approved[$i].Label
                        $response = $responses[$i]

                        if ($response.Success) {
                            $memberCopied++
                            Write-NCMessage "Copied $memberType '$memberLabel' to '$($destinationGroup.DisplayName)'." -Level SUCCESS
                        }
                        elseif ($response.ErrorMessage -match 'already exist' -or $response.ErrorMessage -match 'exists') {
                            $memberSkipped++
                            Write-NCMessage "$memberType '$memberLabel' is already a member of '$($destinationGroup.DisplayName)'." -Level WARNING
                        }
                        else {
                            Write-NCMessage "Failed to copy $memberType '$memberLabel' to '$($destinationGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                        }
                    }
                }
            }
        }

        if ($PassThru.IsPresent) {
            [pscustomobject][ordered]@{
                SourceGroupName      = $sourceGroup.DisplayName
                SourceGroupId        = $sourceGroup.Id
                DestinationGroupName = $destinationGroup.DisplayName
                DestinationGroupId   = $destinationGroup.Id
                DestinationCreated   = $destinationCreated
                MembersCopied        = $memberCopied
                MembersSkipped       = $memberSkipped
                OwnersCopied         = $ownerCopied
                OwnersSkipped        = $ownerSkipped
                Status               = 'Completed'
            }
        }
    }
}

function Add-EntraGroupUser {
    <#
    .SYNOPSIS
        Adds one or more users to an Entra group.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then adds
        the provided users by UPN/display name or object ID. Accepts pipeline input for users.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER UserIdentifier
        User principal name, display name, or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER TreatInputAsId
        Treat every UserIdentifier as an object ID without attempting name resolution.
    .PARAMETER PassThru
        Emit a summary object for each processed user.
    .EXAMPLE
        "user1@contoso.com","user2@contoso.com" | Add-EntraGroupUser -GroupName "My Entra Group"
    .EXAMPLE
        Add-EntraGroupUser -GroupId "00000000-0000-0000-0000-000000000000" -UserIdentifier "user1@contoso.com"
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Alias('Group', 'DisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [string]$GroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Mail', 'Id', 'UserId')]
        [string[]]$UserIdentifier,

        [switch]$TreatInputAsId,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $users = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not $graphConnected) { return }

        foreach ($entry in $UserIdentifier) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$users.Add($entry.Trim())
            }
        }
    }

    end {
        if (-not $graphConnected) { return }
        if ($users.Count -eq 0) {
            Write-NCMessage "No users were specified." -Level WARNING
            return
        }

        $resolvedGroup = $null
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                $resolvedGroup = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Entra group with ID '$GroupId' not found: $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            $escapedName = $GroupName.Replace("'", "''")
            try {
                $resolvedGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-NCMessage "Unable to resolve group '$GroupName': $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $resolvedGroup) {
                try {
                    $resolvedGroup = Get-MgGroup -GroupId $GroupName -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Entra group '$GroupName' not found by name or ID" -Level ERROR
                    return
                }
            }
        }

        try {
            $resolvedGroup = Get-MgGroup -GroupId $resolvedGroup.Id -Property @('id', 'displayName', 'onPremisesSyncEnabled') -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to retrieve full details for group '$($resolvedGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
            return
        }

        if ($resolvedGroup.OnPremisesSyncEnabled -eq $true) {
            Write-NCMessage "Group '$($resolvedGroup.DisplayName)' is synchronized from on-premises AD, so membership can't be changed directly in Entra. Update the group in AD and let sync propagate the change." -Level ERROR
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $uniqueUsers = @($users | Select-Object -Unique)
        $membersRefUrl = "/groups/$($resolvedGroup.Id)/members/`$ref"
        Write-NCMessage "Processing $($uniqueUsers.Count) user(s) in Graph batches (20 per request) ..." -Level INFO

        for ($offset = 0; $offset -lt $uniqueUsers.Count; $offset += 20) {
            $chunk = @($uniqueUsers[$offset..([Math]::Min($offset + 20, $uniqueUsers.Count) - 1)])
            $targets = @(Resolve-NCEntraGroupUserTarget -UserIdentifier $chunk -TreatInputAsId:$TreatInputAsId)

            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($target in $targets) {
                if ($PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Add user '$($target.Label)'")) {
                    $approved.Add($target)
                }
            }
            if ($approved.Count -eq 0) { continue }

            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'POST'; Url = $membersRefUrl; Body = @{ '@odata.id' = (Get-NCGraphDirectoryObjectUri -Id $approved[$i].Id) } }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Adding users to $($resolvedGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $userId = $approved[$i].Id
                $userLabel = $approved[$i].Label
                $response = $responses[$i]
                $status = 'Added'

                if ($response.Success) {
                    Write-NCMessage "Added user '$userLabel' to group '$($resolvedGroup.DisplayName)'." -Level SUCCESS
                }
                elseif ($response.ErrorMessage -match 'added object references already exist') {
                    $status = 'Exists'
                    Write-NCMessage "User '$userLabel' is already a member of '$($resolvedGroup.DisplayName)'." -Level WARNING
                }
                elseif ($response.ErrorMessage -match 'on-premises mastered Directory Sync objects|currently undergoing migration') {
                    $status = 'Failed'
                    Write-NCMessage "Group '$($resolvedGroup.DisplayName)' is synchronized from on-premises AD, so membership can't be changed directly in Entra. Update the group in AD and let sync propagate the change." -Level ERROR
                }
                else {
                    $status = 'Failed'
                    Write-NCMessage "Failed to add user '$userLabel' to '$($resolvedGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }

                if ($PassThru.IsPresent) {
                    $results.Add([pscustomobject][ordered]@{
                            GroupName  = $resolvedGroup.DisplayName
                            GroupId    = $resolvedGroup.Id
                            MemberName = $userLabel
                            MemberId   = $userId
                            MemberType = 'User'
                            Status     = $status
                        }) | Out-Null
                }
            }
        }

        if ($PassThru.IsPresent -and $results.Count -gt 0) {
            $results
        }
    }
}

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
    .PARAMETER BatchSize
        Number of processed groups before flushing partial CSV output.
    .PARAMETER Resume
        Resume from the latest matching CSV in the target folder or from -CsvPath.
    .PARAMETER CsvPath
        Explicit CSV file to resume. When omitted, the most recent matching CSV in the target folder is used.
    .PARAMETER MaxConsecutiveErrors
        Stop after this many consecutive group-member retrieval failures.
    .PARAMETER BatchSize
        Number of processed groups before flushing partial CSV output.
    .PARAMETER Resume
        Resume from the latest matching CSV in the target folder or from -CsvPath.
    .PARAMETER CsvPath
        Explicit CSV file to resume. When omitted, the most recent matching CSV in the target folder is used.
    .PARAMETER MaxConsecutiveErrors
        Stop after this many consecutive group-member retrieval failures.
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
        [switch]$GridView,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25,
        [switch]$Resume,
        [string]$CsvPath,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5
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
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $exportAll = $All.IsPresent -or $requestedGroups.Count -eq 0
            $emitCsv = $Csv.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvFolder) -or $exportAll -or $Resume.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvPath)
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

            $normalizeText = {
                param($value)
                return Get-NormalizedText -Value $value
            }
            $buildMemberKey = {
                param($row)

                $groupIdentity = & $normalizeText $row.'Group Identity'
                if (-not $groupIdentity) {
                    $groupIdentity = & $normalizeText $row.'Group Name'
                }
                if (-not $groupIdentity) {
                    $groupIdentity = & $normalizeText $row.'Group Primary Smtp Address'
                }

                $memberPrimary = & $normalizeText $row.'Member Primary Smtp Address'
                $memberDisplay = & $normalizeText $row.'Member Display Name'
                return "{0}|{1}|{2}" -f $groupIdentity, $memberPrimary, $memberDisplay
            }
            $existingMemberKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $results = [System.Collections.Generic.List[object]]::new()
            $processedSinceFlush = 0
            $consecutiveErrors = 0
            $aborted = $false
            $csvPathResolved = $null
            $counter = 0
            $total = $groups.Count

            if ($emitCsv) {
                $folderForExport = $folder
                $defaultCsvPath = New-File (Join-Path -Path $folderForExport -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-DistributionGroups-Report.csv")
                $csvPathResolved = $defaultCsvPath

                if ($Resume) {
                    $resumePath = $null
                    if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                        $resumePath = $CsvPath
                    }
                    else {
                        $existingCsv = Get-ChildItem -LiteralPath $folderForExport -File -Filter "*_M365-DistributionGroups-Report.csv" |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1
                        if ($existingCsv) {
                            $resumePath = $existingCsv.FullName
                        }
                    }

                    if ($resumePath) {
                        $csvPathResolved = $resumePath
                        if (Test-Path -LiteralPath $csvPathResolved) {
                            try {
                                foreach ($row in (Import-CSV -LiteralPath $csvPathResolved -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                                    $null = $existingMemberKeys.Add((& $buildMemberKey $row))
                                }
                                Write-NCMessage ("Resuming distribution group export from {0}; {1} row(s) already recorded." -f $csvPathResolved, $existingMemberKeys.Count) -Level INFO
                            }
                            catch {
                                Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csvPathResolved, $_.Exception.Message) -Level WARNING
                                $existingMemberKeys.Clear()
                                $csvPathResolved = $defaultCsvPath
                            }
                        }
                        else {
                            Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csvPathResolved) -Level INFO
                        }
                    }
                    else {
                        Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $csvPathResolved) -Level INFO
                    }
                }

                Write-NCMessage ("Distribution group export will flush every {0} group(s). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
                Write-NCMessage "Saving report to $csvPathResolved" -Level DEBUG
            }

            $writeBuffer = {
                param([System.Collections.Generic.List[object]]$buffer)

                if ($buffer.Count -eq 0) {
                    return
                }

                $exportRows = $buffer | Select-Object 'Group Name', 'Group Identity', 'Group Primary Smtp Address', 'Member Display Name', 'Member FirstName', 'Member LastName', 'Member Primary Smtp Address', 'Member Company', 'Member City'
                if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    $exportRows | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                }
                else {
                    $exportRows | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                }
                $buffer.Clear()
            }

            foreach ($group in $groups) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $total
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $total - $Percentage%" -PercentComplete $Percentage

                try {
                    $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Failed to retrieve members for '$($group.DisplayName)': $($_.Exception.Message)" -Level WARNING
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        $aborted = $true
                        break
                    }
                    continue
                }

                if (-not $members -or $members.Count -eq 0) {
                    if ($exportAll) {
                        if ($emitCsv) {
                            $results.Add([pscustomobject][ordered]@{
                                    'Group Name'                  = $group.DisplayName
                                    'Group Identity'              = $group.Identity
                                    'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                                    'Member Display Name'         = $null
                                    'Member FirstName'            = $null
                                    'Member LastName'             = $null
                                    'Member Primary Smtp Address' = $null
                                    'Member Company'              = $null
                                    'Member City'                 = $null
                                }) | Out-Null
                        }
                        else {
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
                    }
                    $consecutiveErrors = 0
                    continue
                }

                foreach ($member in $members) {
                    if ($exportAll) {
                        if ($emitCsv) {
                            $row = [ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Identity'              = $group.Identity
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
                    }
                    else {
                        if ($emitCsv) {
                            $row = [ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Identity'              = $group.Identity
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
                    }

                    $rowObject = [pscustomobject]$row
                    if ($emitCsv) {
                        $rowKey = & $buildMemberKey $rowObject
                        if ($Resume -and $existingMemberKeys.Contains($rowKey)) {
                            continue
                        }
                        $null = $existingMemberKeys.Add($rowKey)
                    }

                    $results.Add($rowObject) | Out-Null
                }

                $consecutiveErrors = 0
                if ($emitCsv) {
                    $processedSinceFlush++
                    if ($BatchSize -gt 0 -and $results.Count -gt 0 -and $processedSinceFlush -ge $BatchSize) {
                        & $writeBuffer $results
                        $processedSinceFlush = 0
                    }
                }

                if ($aborted) {
                    break
                }
            }

            if ($aborted -and $emitCsv -and $results.Count -gt 0) {
                & $writeBuffer $results
            }

            if (-not $results -or $results.Count -eq 0) {
                if ($emitCsv -and (Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    Write-NCMessage "No new distribution group members found. Existing CSV at $csvPathResolved already contains the requested rows." -Level INFO
                }
                else {
                    Write-NCMessage "No members found for the specified distribution groups." -Level WARNING
                }
                return
            }

            if ($GridView.IsPresent) {
                $results | Out-GridView -Title "M365 Distribution Groups"
            }
            elseif ($emitCsv) {
                & $writeBuffer $results
                if ($aborted) {
                    Write-NCMessage "Distribution group export stopped early. Partial data kept at $csvPathResolved." -Level ERROR
                }
                else {
                    Write-NCMessage "Distribution group membership exported to $csvPathResolved." -Level SUCCESS
                }
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

Set-Alias -Name Export-DG -Value Export-DistributionGroups

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
        [switch]$GridView,
        [ValidateRange(1, 500)]
        [int]$BatchSize = 25,
        [switch]$Resume,
        [string]$CsvPath,
        [ValidateRange(1, 100)]
        [int]$MaxConsecutiveErrors = 5
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
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $exportAll = $All.IsPresent -or $requestedGroups.Count -eq 0
            $emitCsv = $Csv.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvFolder) -or $exportAll -or $Resume.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvPath)
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

            $normalizeText = {
                param($value)
                return Get-NormalizedText -Value $value
            }
            $buildMemberKey = {
                param($row)

                $groupIdentity = & $normalizeText $row.'Group Identity'
                if (-not $groupIdentity) {
                    $groupIdentity = & $normalizeText $row.'Group Name'
                }
                if (-not $groupIdentity) {
                    $groupIdentity = & $normalizeText $row.'Group Primary Smtp Address'
                }

                $memberPrimary = & $normalizeText $row.'Member Primary Smtp Address'
                $memberDisplay = & $normalizeText $row.'Member Display Name'
                return "{0}|{1}|{2}" -f $groupIdentity, $memberPrimary, $memberDisplay
            }
            $existingMemberKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $results = [System.Collections.Generic.List[object]]::new()
            $processedSinceFlush = 0
            $consecutiveErrors = 0
            $aborted = $false
            $csvPathResolved = $null
            $counter = 0
            $total = $groups.Count

            if ($emitCsv) {
                $folderForExport = $folder
                $defaultCsvPath = New-File (Join-Path -Path $folderForExport -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-DynamicDistributionGroups-Report.csv")
                $csvPathResolved = $defaultCsvPath

                if ($Resume) {
                    $resumePath = $null
                    if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                        $resumePath = $CsvPath
                    }
                    else {
                        $existingCsv = Get-ChildItem -LiteralPath $folderForExport -File -Filter "*_M365-DynamicDistributionGroups-Report.csv" |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1
                        if ($existingCsv) {
                            $resumePath = $existingCsv.FullName
                        }
                    }

                    if ($resumePath) {
                        $csvPathResolved = $resumePath
                        if (Test-Path -LiteralPath $csvPathResolved) {
                            try {
                                foreach ($row in (Import-CSV -LiteralPath $csvPathResolved -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                                    $null = $existingMemberKeys.Add((& $buildMemberKey $row))
                                }
                                Write-NCMessage ("Resuming dynamic distribution group export from {0}; {1} row(s) already recorded." -f $csvPathResolved, $existingMemberKeys.Count) -Level INFO
                            }
                            catch {
                                Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csvPathResolved, $_.Exception.Message) -Level WARNING
                                $existingMemberKeys.Clear()
                                $csvPathResolved = $defaultCsvPath
                            }
                        }
                        else {
                            Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csvPathResolved) -Level INFO
                        }
                    }
                    else {
                        Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $csvPathResolved) -Level INFO
                    }
                }

                Write-NCMessage ("Dynamic distribution group export will flush every {0} group(s). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
                Write-NCMessage "Saving report to $csvPathResolved" -Level DEBUG
            }

            $writeBuffer = {
                param([System.Collections.Generic.List[object]]$buffer)

                if ($buffer.Count -eq 0) {
                    return
                }

                $exportRows = $buffer | Select-Object 'Group Name', 'Group Identity', 'Group Primary Smtp Address', 'Member Display Name', 'Member FirstName', 'Member LastName', 'Member Primary Smtp Address', 'Member Company', 'Member City'
                if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    $exportRows | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                }
                else {
                    $exportRows | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                }
                $buffer.Clear()
            }

            foreach ($group in $groups) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $total
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $total - $Percentage%" -PercentComplete $Percentage

                try {
                    $members = @(Get-DynamicDistributionGroupMember -Identity $group.Identity -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Failed to retrieve members for '$($group.DisplayName)': $($_.Exception.Message)" -Level WARNING
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        $aborted = $true
                        break
                    }
                    continue
                }

                if (-not $members -or $members.Count -eq 0) {
                    if ($exportAll) {
                        if ($emitCsv) {
                            $results.Add([pscustomobject][ordered]@{
                                    'Group Name'                  = $group.DisplayName
                                    'Group Identity'              = $group.Identity
                                    'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                                    'Member Display Name'         = $null
                                    'Member FirstName'            = $null
                                    'Member LastName'             = $null
                                    'Member Primary Smtp Address' = $null
                                    'Member Company'              = $null
                                    'Member City'                 = $null
                                }) | Out-Null
                        }
                        else {
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
                    }
                    $consecutiveErrors = 0
                    continue
                }

                foreach ($member in $members) {
                    if ($exportAll) {
                        if ($emitCsv) {
                            $row = [ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Identity'              = $group.Identity
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
                    }
                    else {
                        if ($emitCsv) {
                            $row = [ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Identity'              = $group.Identity
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
                    }

                    $rowObject = [pscustomobject]$row
                    if ($emitCsv) {
                        $rowKey = & $buildMemberKey $rowObject
                        if ($Resume -and $existingMemberKeys.Contains($rowKey)) {
                            continue
                        }
                        $null = $existingMemberKeys.Add($rowKey)
                    }

                    $results.Add($rowObject) | Out-Null
                }

                $consecutiveErrors = 0
                if ($emitCsv) {
                    $processedSinceFlush++
                    if ($BatchSize -gt 0 -and $results.Count -gt 0 -and $processedSinceFlush -ge $BatchSize) {
                        & $writeBuffer $results
                        $processedSinceFlush = 0
                    }
                }

                if ($aborted) {
                    break
                }
            }

            if ($aborted -and $emitCsv -and $results.Count -gt 0) {
                & $writeBuffer $results
            }

            if (-not $results -or $results.Count -eq 0) {
                if ($emitCsv -and (Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    Write-NCMessage "No new dynamic distribution group members found. Existing CSV at $csvPathResolved already contains the requested rows." -Level INFO
                }
                else {
                    Write-NCMessage "No members found for the specified dynamic distribution groups." -Level WARNING
                }
                return
            }

            if ($GridView.IsPresent) {
                $results | Out-GridView -Title "M365 Dynamic Distribution Groups"
            }
            elseif ($emitCsv) {
                & $writeBuffer $results
                if ($aborted) {
                    Write-NCMessage "Dynamic distribution group export stopped early. Partial data kept at $csvPathResolved." -Level ERROR
                }
                else {
                    Write-NCMessage "Dynamic distribution group membership exported to $csvPathResolved." -Level SUCCESS
                }
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

Set-Alias -Name Export-DDG -Value Export-DynamicDistributionGroups

function Export-EmptyEntraGroups {
    <#
    .SYNOPSIS
        Exports Entra groups that have no members.
    .DESCRIPTION
        Connects to Microsoft Graph, enumerates groups, checks membership for each one, and exports
        the groups with zero members to CSV by default.
    .PARAMETER CsvFolder
        Destination folder for the CSV file when exporting the report.
    .PARAMETER Csv
        When present, export the report to CSV. Defaults to on.
    .EXAMPLE
        Export-EmptyEntraGroups
    .EXAMPLE
        Export-EmptyEntraGroups -CsvFolder 'C:\Reports\Groups'
    #>
    [CmdletBinding()]
    param(
        [string]$CsvFolder,
        [bool]$Csv = $true
    )

    begin {
        Set-ProgressAndInfoPreferences
        $report = [System.Collections.Generic.List[object]]::new()
    }

    process {
        try {
            if (-not (Test-MgGraphConnection -Scopes @('Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false)) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
                return
            }

            $groups = @(Get-MgGroup -All -ErrorAction Stop)
            if (-not $groups -or $groups.Count -eq 0) {
                Write-NCMessage "No Entra groups found." -Level WARNING
                return
            }

            $totalGroups = $groups.Count
            $processedCount = 0

            Write-NCMessage "Processing $totalGroups group(s) in Graph batches (20 per request) ..." -Level INFO
            $memberRequests = @(for ($i = 0; $i -lt $totalGroups; $i++) {
                    @{ Id = "g$i"; Method = 'GET'; Url = "/groups/$([uri]::EscapeDataString([string]$groups[$i].Id))/members?`$select=id&`$top=1" }
                })
            $memberLookup = @{}
            foreach ($memberResult in @(Invoke-NCGraphBatch -Requests $memberRequests -Activity 'Checking group members')) {
                $memberLookup[$memberResult.Id] = $memberResult
            }

            foreach ($group in $groups) {
                $processedCount++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $totalGroups
                Write-Progress -Activity "Checking $($group.DisplayName)" -Status "$processedCount of $totalGroups - $Percentage%" -PercentComplete $Percentage

                $memberResult = $memberLookup["g$($processedCount - 1)"]
                if (-not $memberResult.Success) {
                    Write-NCMessage "Unable to read members for group '$($group.DisplayName)'. $($memberResult.ErrorMessage)" -Level WARNING
                    continue
                }

                $members = if ($null -ne $memberResult.Body -and $null -ne $memberResult.Body.value) { @($memberResult.Body.value) } else { @() }
                if ($members.Count -gt 0) {
                    continue
                }

                $groupType = if ($group.GroupTypes -contains 'Unified' -and $group.SecurityEnabled) {
                    'Microsoft 365 (security-enabled)'
                }
                elseif ($group.GroupTypes -contains 'Unified' -and -not $group.SecurityEnabled) {
                    'Microsoft 365'
                }
                elseif (-not ($group.GroupTypes -contains 'Unified') -and $group.SecurityEnabled -and $group.MailEnabled) {
                    'Mail-enabled security'
                }
                elseif (-not ($group.GroupTypes -contains 'Unified') -and $group.SecurityEnabled) {
                    'Security'
                }
                elseif (-not ($group.GroupTypes -contains 'Unified') -and $group.MailEnabled) {
                    'Distribution'
                }
                else {
                    'N/A'
                }

                $report.Add([pscustomobject][ordered]@{
                        DisplayName     = $group.DisplayName
                        Id              = $group.Id
                        GroupType       = $groupType
                        MemberCount     = 0
                        MailEnabled     = $group.MailEnabled
                        SecurityEnabled = $group.SecurityEnabled
                    }) | Out-Null
            }

            if ($Csv) {
                $folder = if ($CsvFolder) { Test-Folder $CsvFolder } else { Test-Folder $null }
                $csvPath = New-File "$folder\$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-EmptyGroups.csv"
                $report | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                Write-NCMessage "Empty groups report exported to $csvPath." -Level SUCCESS
            }
            else {
                $report | Sort-Object DisplayName
            }
        }
        catch {
            Write-NCMessage "Unable to export empty Entra groups. $($_.Exception.Message)" -Level ERROR
        }
        finally {
            Write-Progress -Activity "Checking empty Entra groups" -Completed
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
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
                return
            }

            $exportAll = $All.IsPresent -or $requestedGroups.Count -eq 0
            $emitCsv = $Csv.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvFolder) -or $exportAll -or $Resume.IsPresent -or -not [string]::IsNullOrWhiteSpace($CsvPath)
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

            $normalizeText = {
                param($value)
                return Get-NormalizedText -Value $value
            }
            $buildMemberKey = {
                param($row)

                $groupIdentity = & $normalizeText $row.'Group Identity'
                if (-not $groupIdentity) {
                    $groupIdentity = & $normalizeText $row.'Group Name'
                }
                if (-not $groupIdentity) {
                    $groupIdentity = & $normalizeText $row.'Group Primary Smtp Address'
                }

                $memberPrimary = & $normalizeText $row.'Member Primary Smtp Address'
                $memberDisplay = & $normalizeText $row.'Member Display Name'
                return "{0}|{1}|{2}" -f $groupIdentity, $memberPrimary, $memberDisplay
            }
            $existingMemberKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $results = [System.Collections.Generic.List[object]]::new()
            $processedSinceFlush = 0
            $consecutiveErrors = 0
            $aborted = $false
            $csvPathResolved = $null
            $counter = 0
            $total = $groups.Count

            if ($emitCsv) {
                $folderForExport = $folder
                $defaultCsvPath = New-File (Join-Path -Path $folderForExport -ChildPath "$((Get-Date -Format $NCVars.DateTimeString_CSV))_M365-UnifiedGroups-Report.csv")
                $csvPathResolved = $defaultCsvPath

                if ($Resume) {
                    $resumePath = $null
                    if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
                        $resumePath = $CsvPath
                    }
                    else {
                        $existingCsv = Get-ChildItem -LiteralPath $folderForExport -File -Filter "*_M365-UnifiedGroups-Report.csv" |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1
                        if ($existingCsv) {
                            $resumePath = $existingCsv.FullName
                        }
                    }

                    if ($resumePath) {
                        $csvPathResolved = $resumePath
                        if (Test-Path -LiteralPath $csvPathResolved) {
                            try {
                                foreach ($row in (Import-CSV -LiteralPath $csvPathResolved -Delimiter $NCVars.CSV_DefaultLimiter -ErrorAction Stop)) {
                                    $null = $existingMemberKeys.Add((& $buildMemberKey $row))
                                }
                                Write-NCMessage ("Resuming Microsoft 365 group export from {0}; {1} row(s) already recorded." -f $csvPathResolved, $existingMemberKeys.Count) -Level INFO
                            }
                            catch {
                                Write-NCMessage ("Unable to read existing CSV '{0}' for resume. {1}" -f $csvPathResolved, $_.Exception.Message) -Level WARNING
                                $existingMemberKeys.Clear()
                                $csvPathResolved = $defaultCsvPath
                            }
                        }
                        else {
                            Write-NCMessage ("Resume requested for '{0}', but the file does not exist. Starting a new report at that path." -f $csvPathResolved) -Level INFO
                        }
                    }
                    else {
                        Write-NCMessage ("Resume requested, but no existing CSV was found. Starting a new report at {0}." -f $csvPathResolved) -Level INFO
                    }
                }

                Write-NCMessage ("Microsoft 365 group export will flush every {0} group(s). Resume: {1}. Stop after {2} consecutive error(s)." -f $BatchSize, $Resume.IsPresent, $MaxConsecutiveErrors) -Level INFO
                Write-NCMessage "Saving report to $csvPathResolved" -Level DEBUG
            }

            $writeBuffer = {
                param([System.Collections.Generic.List[object]]$buffer)

                if ($buffer.Count -eq 0) {
                    return
                }

                $exportRows = $buffer | Select-Object 'Group Name', 'Group Identity', 'Group Primary Smtp Address', 'Member Display Name', 'Member FirstName', 'Member LastName', 'Member Primary Smtp Address', 'Member Company', 'Member City'
                if ((Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    $exportRows | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter -Append
                }
                else {
                    $exportRows | Export-Csv -LiteralPath $csvPathResolved -NoTypeInformation -Encoding $NCVars.CSV_Encoding -Delimiter $NCVars.CSV_DefaultLimiter
                }
                $buffer.Clear()
            }

            foreach ($group in $groups) {
                $counter++
                $Percentage = Get-NCProgressPercent -Current $counter -Total $total
                Write-Progress -Activity "Processing $($group.DisplayName)" -Status "$counter of $total - $Percentage%" -PercentComplete $Percentage

                try {
                    $members = @(Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Member -ErrorAction Stop)
                }
                catch {
                    Write-NCMessage "Failed to retrieve members for '$($group.DisplayName)': $($_.Exception.Message)" -Level WARNING
                    $consecutiveErrors++
                    if ($MaxConsecutiveErrors -gt 0 -and $consecutiveErrors -ge $MaxConsecutiveErrors) {
                        $aborted = $true
                        break
                    }
                    continue
                }

                if (-not $members -or $members.Count -eq 0) {
                    if ($exportAll) {
                        if ($emitCsv) {
                            $results.Add([pscustomobject][ordered]@{
                                    'Group Name'                  = $group.DisplayName
                                    'Group Identity'              = $group.Identity
                                    'Group Primary Smtp Address'  = $group.PrimarySmtpAddress
                                    'Member Display Name'         = $null
                                    'Member FirstName'            = $null
                                    'Member LastName'             = $null
                                    'Member Primary Smtp Address' = $null
                                    'Member Company'              = $null
                                    'Member City'                 = $null
                                }) | Out-Null
                        }
                        else {
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
                    }
                    $consecutiveErrors = 0
                    continue
                }

                foreach ($member in $members) {
                    if ($exportAll) {
                        if ($emitCsv) {
                            $row = [ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Identity'              = $group.Identity
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
                    }
                    else {
                        if ($emitCsv) {
                            $row = [ordered]@{
                                'Group Name'                  = $group.DisplayName
                                'Group Identity'              = $group.Identity
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
                    }

                    $rowObject = [pscustomobject]$row
                    if ($emitCsv) {
                        $rowKey = & $buildMemberKey $rowObject
                        if ($Resume -and $existingMemberKeys.Contains($rowKey)) {
                            continue
                        }
                        $null = $existingMemberKeys.Add($rowKey)
                    }

                    $results.Add($rowObject) | Out-Null
                }

                $consecutiveErrors = 0
                if ($emitCsv) {
                    $processedSinceFlush++
                    if ($BatchSize -gt 0 -and $results.Count -gt 0 -and $processedSinceFlush -ge $BatchSize) {
                        & $writeBuffer $results
                        $processedSinceFlush = 0
                    }
                }

                if ($aborted) {
                    break
                }
            }

            if ($aborted -and $emitCsv -and $results.Count -gt 0) {
                & $writeBuffer $results
            }

            if (-not $results -or $results.Count -eq 0) {
                if ($emitCsv -and (Test-Path -LiteralPath $csvPathResolved) -and ((Get-Item -LiteralPath $csvPathResolved).Length -gt 0)) {
                    Write-NCMessage "No new Microsoft 365 group members found. Existing CSV at $csvPathResolved already contains the requested rows." -Level INFO
                }
                else {
                    Write-NCMessage "No members found for the specified Microsoft 365 groups." -Level WARNING
                }
                return
            }

            if ($GridView.IsPresent) {
                $results | Out-GridView -Title "M365 Unified Groups"
            }
            elseif ($emitCsv) {
                & $writeBuffer $results
                if ($aborted) {
                    Write-NCMessage "Microsoft 365 group export stopped early. Partial data kept at $csvPathResolved." -Level ERROR
                }
                else {
                    Write-NCMessage "Microsoft 365 group membership exported to $csvPathResolved." -Level SUCCESS
                }
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
        "group@contoso.com" | Get-DynamicDistributionGroupFilter -AsObject
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
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
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

Set-Alias -Name Get-DDGRecipientFilter -Value Get-DynamicDistributionGroupFilter

function Get-EntraGroupDevice {
    <#
    .SYNOPSIS
        Shows the Entra groups that a device belongs to.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target device by display name or ID, then
        lists every directory object membership for that device.
    .PARAMETER DeviceIdentifier
        Device display name or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER TreatInputAsId
        Treat the provided DeviceIdentifier as an object ID without attempting name resolution.
    .PARAMETER GridView
        Show additional details in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-EntraGroupDevice -DeviceIdentifier "PC123"
    .EXAMPLE
        "00000000-0000-0000-0000-000000000000" | Get-EntraGroupDevice -TreatInputAsId -GridView
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Device', 'DeviceName', 'Id', 'DeviceId', 'Name', 'Identity', 'DisplayName')]
        [string]$DeviceIdentifier,
        [switch]$TreatInputAsId,
        [switch]$GridView
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $guidPattern = '^[0-9a-fA-F-]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        $queue = [System.Collections.Generic.List[string]]::new()
        $state = @{ Started = $false }

        $flush = {
            $inputs = @($queue)
            $queue.Clear()
            if ($inputs.Count -eq 0) { return }

            if (-not $state.Started) {
                $state.Started = $true
                Write-NCMessage "Processing devices in Graph batches (20 per request) ..." -Level INFO
            }

            $targets = @(Resolve-NCEntraDeviceTargetBatch -DeviceIdentifier $inputs -TreatInputAsId:$TreatInputAsId.IsPresent)
            if ($targets.Count -eq 0) { return }

            $requests = @(for ($i = 0; $i -lt $targets.Count; $i++) {
                    @{ Id = "m$i"; Method = 'GET'; Url = "/devices/$([uri]::EscapeDataString([string]$targets[$i].Id))/memberOf" }
                })
            $lookup = @{}
            foreach ($result in @(Invoke-NCGraphBatchCollection -Requests $requests -Activity 'Reading device group memberships')) {
                $lookup[$result.Id] = $result
            }

            for ($i = 0; $i -lt $targets.Count; $i++) {
                $deviceLabel = $targets[$i].Label
                $result = $lookup["m$i"]

                if (-not $result.Success) {
                    $inputWasId = $TreatInputAsId.IsPresent -or $targets[$i].Input -match $guidPattern
                    if ($inputWasId) {
                        Write-NCMessage "Entra device with ID '$($targets[$i].Input)' not found: $($result.ErrorMessage)" -Level ERROR
                    }
                    else {
                        Write-NCMessage "Unable to read group memberships for device ${deviceLabel}: $($result.ErrorMessage)" -Level ERROR
                    }
                    continue
                }

                $memberships = @($result.Items | ForEach-Object { ConvertTo-NCGraphDirectoryObject -Item $_ })

                Add-EmptyLine
                Write-Verbose "Device ($deviceLabel) - Groups found: $($memberships.Count)"

                if (-not $memberships -or $memberships.Count -eq 0) {
                    Write-NCMessage "No groups found for $deviceLabel." -Level WARNING
                    continue
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
                    $results | Out-GridView -Title "Entra Device Groups - $deviceLabel"
                }
                else {
                    $results | Sort-Object 'Group Name'
                }
            }
        }
    }

    process {
        if (-not $graphConnected) {
            return
        }

        $queue.Add($DeviceIdentifier)
        if ($queue.Count -ge 20) {
            & $flush
        }
    }

    end {
        if ($graphConnected -and $queue.Count -gt 0) {
            & $flush
        }
    }
}

function Get-EntraGroupMembers {
    <#
    .SYNOPSIS
        Shows the members of an Entra group (users, devices, and other directory objects).
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then
        lists every member of that group regardless of type.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER IncludeDeviceUsers
        When members are devices, resolve registered owners and users (may require additional Graph calls).
    .PARAMETER GridView
        Show additional details in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-EntraGroupMembers -GroupName "My Entra Group"
    .EXAMPLE
        Get-EntraGroupMembers -GroupId "00000000-0000-0000-0000-000000000000" -GridView
    .EXAMPLE
        "My Entra Group" | Get-EntraGroupMembers
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Group', 'DisplayName', 'Name', 'Identity')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
        [string]$GroupId,

        [switch]$IncludeDeviceUsers,
        [switch]$GridView
    )

    $graphConnected = Test-MgGraphConnection -Scopes @('Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
    if (-not $graphConnected) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return
    }

    $resolvedGroup = $null
    if ($PSCmdlet.ParameterSetName -eq 'ById') {
        try {
            $resolvedGroup = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Entra group with ID '$GroupId' not found: $($_.Exception.Message)" -Level ERROR
            return
        }
    }
    else {
        $escapedName = $GroupName.Replace("'", "''")
        try {
            $resolvedGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
        }
        catch {
            Write-NCMessage "Unable to resolve group '$GroupName': $($_.Exception.Message)" -Level ERROR
            return
        }

        if (-not $resolvedGroup) {
            try {
                $resolvedGroup = Get-MgGroup -GroupId $GroupName -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Entra group '$GroupName' not found by name or ID" -Level ERROR
                return
            }
        }
    }

    try {
        $members = @(Get-MgGroupMember -GroupId $resolvedGroup.Id -All -ErrorAction Stop)
    }
    catch {
        Write-NCMessage "Unable to read members for group $($resolvedGroup.DisplayName): $($_.Exception.Message)" -Level ERROR
        return
    }

    Add-EmptyLine
    Write-Verbose "Group ($($resolvedGroup.DisplayName)) - Members found: $($members.Count)"

    if (-not $members -or $members.Count -eq 0) {
        Write-NCMessage "No members found for $($resolvedGroup.DisplayName)." -Level WARNING
        return
    }

    $resolveType = {
        param($odataType)
        if ([string]::IsNullOrWhiteSpace($odataType)) {
            return 'DirectoryObject'
        }

        $value = $odataType.ToLowerInvariant()
        if ($value -match 'user') { return 'User' }
        if ($value -match 'device') { return 'Device' }
        if ($value -match 'group') { return 'Group' }
        if ($value -match 'serviceprincipal') { return 'ServicePrincipal' }
        if ($value -match 'orgcontact') { return 'Contact' }
        return 'DirectoryObject'
    }

    $resolveOdataType = {
        param($member)
        $memberProps = if ($member.AdditionalProperties) { $member.AdditionalProperties } else { @{} }
        if ($memberProps.ContainsKey('@odata.type')) { $memberProps['@odata.type'] } else { $null }
    }

    $deviceLookup = @{}
    if ($IncludeDeviceUsers.IsPresent) {
        $deviceRequests = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $members.Count; $i++) {
            if ((& $resolveType (& $resolveOdataType $members[$i])) -ne 'Device') { continue }
            $deviceSegment = [uri]::EscapeDataString([string]$members[$i].Id)
            $deviceRequests.Add(@{ Id = "o$i"; Method = 'GET'; Url = "/devices/$deviceSegment/registeredOwners" })
            $deviceRequests.Add(@{ Id = "u$i"; Method = 'GET'; Url = "/devices/$deviceSegment/registeredUsers" })
        }

        if ($deviceRequests.Count -gt 0) {
            Write-NCMessage "Processing $($deviceRequests.Count / 2) device(s) in Graph batches (20 per request) ..." -Level INFO
            foreach ($result in @(Invoke-NCGraphBatchCollection -Requests @($deviceRequests) -Activity 'Reading device registered owners and users')) {
                $deviceLookup[$result.Id] = $result
            }
        }
    }

    $results = [System.Collections.Generic.List[object]]::new()
    for ($memberIndex = 0; $memberIndex -lt $members.Count; $memberIndex++) {
        $member = $members[$memberIndex]
        $props = if ($member.AdditionalProperties) { $member.AdditionalProperties } else { @{} }
        $odataType = if ($props.ContainsKey('@odata.type')) { $props['@odata.type'] } else { $null }
        $memberType = & $resolveType $odataType
        $displayName = if ($props.ContainsKey('displayName')) { $props.displayName } else { $null }
        $upn = if ($props.ContainsKey('userPrincipalName')) { $props.userPrincipalName } else { $null }
        $mail = if ($props.ContainsKey('mail')) { $props.mail } else { $null }

        $row = [ordered]@{
            'Member Name' = if ($displayName) { $displayName } elseif ($upn) { $upn } else { $null }
            'Member Type' = $memberType
            'Member Id'   = $member.Id
        }

        if ($IncludeDeviceUsers.IsPresent -and $memberType -eq 'Device') {
            $owners = @()
            $users = @()
            $ownerResult = $deviceLookup["o$memberIndex"]
            if ($ownerResult.Success) {
                $owners = @($ownerResult.Items | ForEach-Object { ConvertTo-NCGraphDirectoryObject -Item $_ })
            }
            else {
                Write-NCMessage "Unable to read registered owners for device $($displayName): $($ownerResult.ErrorMessage)" -Level WARNING
            }

            $userResult = $deviceLookup["u$memberIndex"]
            if ($userResult.Success) {
                $users = @($userResult.Items | ForEach-Object { ConvertTo-NCGraphDirectoryObject -Item $_ })
            }
            else {
                Write-NCMessage "Unable to read registered users for device $($displayName): $($userResult.ErrorMessage)" -Level WARNING
            }

            $ownerLabels = $owners | ForEach-Object {
                $ownerProps = if ($_.AdditionalProperties) { $_.AdditionalProperties } else { @{} }
                if ($ownerProps.ContainsKey('userPrincipalName')) { $ownerProps.userPrincipalName } elseif ($ownerProps.ContainsKey('displayName')) { $ownerProps.displayName } else { $_.Id }
            }
            $userLabels = $users | ForEach-Object {
                $userProps = if ($_.AdditionalProperties) { $_.AdditionalProperties } else { @{} }
                if ($userProps.ContainsKey('userPrincipalName')) { $userProps.userPrincipalName } elseif ($userProps.ContainsKey('displayName')) { $userProps.displayName } else { $_.Id }
            }

            $ownerValue = if ($ownerLabels) { ($ownerLabels -join '; ') } else { $null }
            $userValue = if ($userLabels) { ($userLabels -join '; ') } else { $null }
            $combinedValue = $null

            if ($ownerValue -and $userValue -and ($ownerValue -eq $userValue)) {
                $combinedValue = $ownerValue
            }
            elseif ($ownerValue -and $userValue) {
                $combinedValue = "Owners: $ownerValue | Users: $userValue"
            }
            elseif ($ownerValue) {
                $combinedValue = "Owners: $ownerValue"
            }
            elseif ($userValue) {
                $combinedValue = "Users: $userValue"
            }

            $row['Device Owners/Users'] = $combinedValue
        }

        if ($GridView.IsPresent) {
            $row['Member UPN'] = $upn
            $row['Member Mail'] = $mail
            $row['Member OData Type'] = $odataType
        }

        $results.Add([pscustomobject]$row) | Out-Null
    }

    $sorted = $results | Sort-Object 'Member Type', 'Member Name'
    if ($GridView.IsPresent) {
        $sorted | Out-GridView -Title "Entra Group Members - $($resolvedGroup.DisplayName)"
    }
    else {
        $sorted
    }
}

function Get-EntraGroupUser {
    <#
    .SYNOPSIS
        Shows the Entra groups that a user belongs to.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target user by UPN/display name or ID, then
        lists every directory object membership for that user.
    .PARAMETER UserIdentifier
        User principal name, display name, or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER TreatInputAsId
        Treat the provided UserIdentifier as an object ID without attempting name resolution.
    .PARAMETER GridView
        Show additional details in Out-GridView instead of returning objects.
    .EXAMPLE
        Get-EntraGroupUser -UserIdentifier "user@contoso.com"
    .EXAMPLE
        "00000000-0000-0000-0000-000000000000" | Get-EntraGroupUser -TreatInputAsId -GridView
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Mail', 'Id', 'UserId', 'DisplayName', 'Identity')]
        [string]$UserIdentifier,
        [switch]$TreatInputAsId,
        [switch]$GridView
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $guidPattern = '^[0-9a-fA-F-]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        $queue = [System.Collections.Generic.List[string]]::new()
        $state = @{ Started = $false }

        $flush = {
            $inputs = @($queue)
            $queue.Clear()
            if ($inputs.Count -eq 0) { return }

            if (-not $state.Started) {
                $state.Started = $true
                Write-NCMessage "Processing users in Graph batches (20 per request) ..." -Level INFO
            }

            # (a) Resolve identities: object IDs with one batched GET, everything else through the user resolver.
            $idLookup = @{}
            $idRequests = @(for ($i = 0; $i -lt $inputs.Count; $i++) {
                    if ($TreatInputAsId.IsPresent -or $inputs[$i] -match $guidPattern) {
                        @{ Id = "u$i"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString($inputs[$i]))?`$select=id,userPrincipalName,displayName" }
                    }
                })
            if ($idRequests.Count -gt 0) {
                foreach ($result in @(Invoke-NCGraphBatch -Requests $idRequests -Activity 'Resolving users')) {
                    $idLookup[$result.Id] = $result
                }
            }

            $nameInputs = @(for ($i = 0; $i -lt $inputs.Count; $i++) {
                    if (-not ($TreatInputAsId.IsPresent -or $inputs[$i] -match $guidPattern)) { $inputs[$i] }
                })
            $failedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolved = if ($nameInputs.Count -gt 0) {
                Resolve-NCGraphUserBatch -Identifier $nameInputs -Property @('id', 'userPrincipalName', 'displayName') -FailedIdentifier $failedNames
            }
            else {
                @{}
            }

            # (b) Build the list of users to read, keeping input order.
            $targets = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $inputs.Count; $i++) {
                $identifier = $inputs[$i]
                $user = $null

                if ($TreatInputAsId.IsPresent -or $identifier -match $guidPattern) {
                    $result = $idLookup["u$i"]
                    if (-not $result.Success -or -not $result.Body) {
                        Write-NCMessage "Entra user with ID '$identifier' not found: $($result.ErrorMessage)" -Level ERROR
                        continue
                    }
                    $user = [pscustomobject]$result.Body
                }
                else {
                    $user = $resolved[$identifier.Trim()]
                    if (-not $user -and $failedNames.Contains($identifier.Trim())) {
                        # The resolver already reported this lookup failure.
                        continue
                    }
                    if (-not $user) {
                        $escapedUser = $identifier.Replace("'", "''")
                        try {
                            $userMatches = Get-MgUser -Filter "displayName eq '$escapedUser'" -All -ErrorAction Stop
                        }
                        catch {
                            Write-NCMessage "Unable to resolve user '$identifier': $($_.Exception.Message)" -Level ERROR
                            continue
                        }

                        if (-not $userMatches -or $userMatches.Count -eq 0) {
                            Write-NCMessage "User '$identifier' not found" -Level WARNING
                            continue
                        }

                        if ($userMatches.Count -gt 1) {
                            Write-NCMessage "Multiple users matched '$identifier'. Using the first result ($($userMatches[0].UserPrincipalName))." -Level WARNING
                        }

                        $matched = $userMatches | Select-Object -First 1
                        $user = [pscustomobject]@{ id = $matched.Id; userPrincipalName = $matched.UserPrincipalName; displayName = $matched.DisplayName }
                    }
                }

                if (-not $user -or -not $user.id) {
                    continue
                }

                $userLabel = if ($user.userPrincipalName) { $user.userPrincipalName } else { $user.displayName }
                $targets.Add([pscustomobject]@{ Id = [string]$user.id; Label = $userLabel })
            }

            if ($targets.Count -eq 0) { return }

            # (c) Read memberships for every resolved user in batches.
            $requests = @(for ($i = 0; $i -lt $targets.Count; $i++) {
                    @{ Id = "m$i"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString($targets[$i].Id))/memberOf?`$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,description,mailNickname" }
                })
            $lookup = @{}
            foreach ($result in @(Invoke-NCGraphBatchCollection -Requests $requests -Activity 'Reading user group memberships')) {
                $lookup[$result.Id] = $result
            }

            # (d) Emit per user, in input order.
            for ($i = 0; $i -lt $targets.Count; $i++) {
                $userLabel = $targets[$i].Label
                $result = $lookup["m$i"]

                if (-not $result.Success) {
                    Write-NCMessage "Unable to read group memberships for user ${userLabel}: $($result.ErrorMessage)" -Level ERROR
                    continue
                }

                $memberships = @($result.Items | ForEach-Object { ConvertTo-NCGraphDirectoryObject -Item $_ })

                Add-EmptyLine
                Write-Verbose "User ($userLabel) - Groups found: $($memberships.Count)"

                if (-not $memberships -or $memberships.Count -eq 0) {
                    Write-NCMessage "No groups found for $userLabel." -Level WARNING
                    continue
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
                    $results | Out-GridView -Title "Entra User Groups - $userLabel"
                }
                else {
                    $results | Sort-Object 'Group Name'
                }
            }
        }
    }

    process {
        if (-not $graphConnected) {
            return
        }

        $queue.Add($UserIdentifier)
        if ($queue.Count -ge 20) {
            & $flush
        }
    }

    end {
        if ($graphConnected -and $queue.Count -gt 0) {
            & $flush
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
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Exchange Online Management module. Please check logs." -Level ERROR
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
            $Percentage = Get-NCProgressPercent -Current $counter -Total $total
            Write-Progress -Activity "Processing $($group.Name)" -Status "$counter of $total - $Percentage%" -PercentComplete $Percentage

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
        $graphConnected = Test-MgGraphConnection
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $queue = [System.Collections.Generic.List[object]]::new()
        $state = @{ Started = $false }

        # Builds and emits the output rows for one input.
        $emit = {
            param([string]$ResolvedPrincipal, [string]$RecipientType, [object[]]$Memberships)

            Add-EmptyLine
            Write-Verbose "$RecipientType ($ResolvedPrincipal) - Groups found: $($Memberships.Count)"

            if (-not $Memberships -or $Memberships.Count -eq 0) {
                Write-NCMessage "No groups found for $ResolvedPrincipal." -Level WARNING
                return
            }

            $results = [System.Collections.Generic.List[object]]::new()
            foreach ($membership in $Memberships) {
                $props = if ($membership.AdditionalProperties) { $membership.AdditionalProperties } else { @{} }
                $row = [ordered]@{
                    GroupName = if ($props.ContainsKey('displayName')) { $props.displayName } else { $null }
                    GroupMail = if ($props.ContainsKey('mail')) { $props.mail } else { $null }
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
                $results | Out-GridView -Title "M365 User Groups - $ResolvedPrincipal"
            }
            else {
                $results | Sort-Object GroupName
            }
        }

        # Resolves the queued user-path inputs and reads their memberships in Graph batches.
        $flush = {
            $entries = @($queue)
            $queue.Clear()
            if ($entries.Count -eq 0) { return }

            if (-not $state.Started) {
                $state.Started = $true
                Write-NCMessage "Processing users in Graph batches (20 per request) ..." -Level INFO
            }

            # (a) Resolve every queued identity.
            $failedUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolved = Resolve-NCGraphUserBatch -Identifier @($entries | ForEach-Object { $_.UserId }) -Property @('id', 'userPrincipalName', 'displayName') -FailedIdentifier $failedUsers

            # (b) Keep the users that resolved, in input order.
            $targets = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $entries) {
                $user = $resolved[$entry.UserId.Trim()]
                if (-not $user -or -not $user.id) {
                    if (-not $failedUsers.Contains($entry.UserId.Trim())) {
                        Write-NCMessage "Unable to resolve user $($entry.UserId) in Microsoft Graph: user not found." -Level ERROR
                    }
                    continue
                }
                $targets.Add([pscustomobject]@{ Entry = $entry; GraphId = [string]$user.id })
            }

            if ($targets.Count -eq 0) { return }

            # (c) Read memberships for every resolved user.
            $requests = @(for ($i = 0; $i -lt $targets.Count; $i++) {
                    @{ Id = "m$i"; Method = 'GET'; Url = "/users/$([uri]::EscapeDataString($targets[$i].GraphId))/memberOf?`$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,description,mailNickname" }
                })
            $lookup = @{}
            foreach ($result in @(Invoke-NCGraphBatchCollection -Requests $requests -Activity 'Reading user group memberships')) {
                $lookup[$result.Id] = $result
            }

            # (d) Emit per input, in input order.
            for ($i = 0; $i -lt $targets.Count; $i++) {
                $entry = $targets[$i].Entry
                $result = $lookup["m$i"]

                if (-not $result.Success) {
                    Write-NCMessage "Unable to read group memberships for $($entry.Principal): $($result.ErrorMessage)" -Level ERROR
                    continue
                }

                $memberships = @($result.Items | ForEach-Object { ConvertTo-NCGraphDirectoryObject -Item $_ })
                & $emit $entry.Principal $entry.RecipientType $memberships
            }
        }
    }

    process {
        if (-not $graphConnected) {
            return
        }

        $resolvedPrincipal = Find-UserRecipient -UserPrincipalName $UserPrincipalName
        if (-not $resolvedPrincipal) {
            & $flush
            Write-NCMessage "Unable to resolve user recipient for $UserPrincipalName" -Level ERROR
            return
        }

        $recipientType = (Get-Recipient -Identity $resolvedPrincipal).RecipientTypeDetails
        $memberships = @()

        switch ($recipientType) {
            'MailContact' {
                & $flush # Keep output in input order: earlier queued users are emitted first.
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

                & $emit $resolvedPrincipal $recipientType $memberships
            }
            'MailUniversalDistributionGroup' {
                & $flush # Keep output in input order: earlier queued users are emitted first.
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

                & $emit $resolvedPrincipal $recipientType $memberships
            }
            default {
                try {
                    $recipient = Get-Mailbox -Identity $resolvedPrincipal -ErrorAction Stop # Preserve the Exchange-first path for regular mailboxes.
                    $userId = if ($recipient.WindowsLiveID) { $recipient.WindowsLiveID } elseif ($recipient.PrimarySmtpAddress) { $recipient.PrimarySmtpAddress } else { $resolvedPrincipal }
                }
                catch {
                    # Not a regular mailbox: the batch resolver tries the identity directly and falls back to Find-UserRecipient -PreferGraphIdentity on 404.
                    $userId = $resolvedPrincipal
                }

                $queue.Add([pscustomobject]@{ Principal = [string]$resolvedPrincipal; RecipientType = $recipientType; UserId = [string]$userId })
                if ($queue.Count -ge 20) {
                    & $flush
                }
            }
        }
    }

    end {
        if ($graphConnected -and $queue.Count -gt 0) {
            & $flush
        }
    }
}

function New-EntraSecurityGroup {
    <#
    .SYNOPSIS
        Creates a new Entra security group.
    .DESCRIPTION
        Connects to Microsoft Graph and creates a security-enabled, mail-disabled group. The
        mail nickname is generated automatically from the display name unless overridden.
    .PARAMETER GroupName
        Display name of the new Entra security group.
    .PARAMETER Description
        Optional group description.
    .PARAMETER MailNickname
        Optional mail nickname. When omitted, a sanitized value is generated from GroupName.
    .PARAMETER PassThru
        Emit the created group object.
    .EXAMPLE
        New-EntraSecurityGroup -GroupName "Sec - Finance"
    .EXAMPLE
        New-EntraSecurityGroup -GroupName "Sec - Finance" -Description "Finance security group"
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias('DisplayName', 'Name')]
        [string]$GroupName,

        [string]$Description,
        [string]$MailNickname,
        [switch]$PassThru
    )

    $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
    if (-not $graphConnected) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return
    }

    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
        return
    }

    if ([string]::IsNullOrWhiteSpace($GroupName)) {
        Write-NCMessage "GroupName cannot be empty." -Level WARNING
        return
    }

    $sanitizedNickname = if ([string]::IsNullOrWhiteSpace($MailNickname)) {
        [regex]::Replace($GroupName, '[^a-zA-Z0-9]', '')
    }
    else {
        $MailNickname.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($sanitizedNickname)) {
        $sanitizedNickname = "group$((Get-Date).ToString('yyyyMMddHHmmss'))"
    }

    $groupBody = @{
        displayName     = $GroupName
        mailEnabled     = $false
        mailNickname    = $sanitizedNickname
        securityEnabled = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($Description)) {
        $groupBody.description = $Description
    }

    if (-not $PSCmdlet.ShouldProcess($GroupName, 'Create security group')) {
        return
    }

    try {
        $createdGroup = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/groups' -Method POST -Body ($groupBody | ConvertTo-Json -Depth 10) -ContentType 'application/json'
        Write-NCMessage "Created security group '$GroupName'." -Level SUCCESS

        if ($PassThru.IsPresent) {
            [pscustomobject]@{
                GroupName       = $createdGroup.displayName
                GroupId         = $createdGroup.id
                MailNickname    = $createdGroup.mailNickname
                Description     = $createdGroup.description
                SecurityEnabled = $createdGroup.securityEnabled
                MailEnabled     = $createdGroup.mailEnabled
            }
        }
    }
    catch {
        Write-NCMessage "Failed to create security group '$GroupName': $($_.Exception.Message)" -Level ERROR
    }
}

function Remove-EntraGroupDevice {
    <#
    .SYNOPSIS
        Removes one or more devices from an Entra group.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then removes
        the provided devices by display name or object ID. Accepts pipeline input for devices.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER DeviceIdentifier
        Device display name or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER ClearAll
        Remove all device members from the Entra group (users and other objects are left untouched).
    .PARAMETER TreatInputAsId
        Treat every DeviceIdentifier as an object ID without attempting name resolution.
    .PARAMETER PassThru
        Emit a summary object for each processed device.
    .EXAMPLE
        "PC1", "PC2" | Remove-EntraGroupDevice -GroupName "My Entra Group"
    .EXAMPLE
        Remove-EntraGroupDevice -GroupId "00000000-0000-0000-0000-000000000000" -DeviceIdentifier "PC1"
    .EXAMPLE
        Remove-EntraGroupDevice -GroupName "My Entra Group" -ClearAll
    .EXAMPLE
        Remove-EntraGroupDevice -GroupName "My Entra Group" -ClearAll -WhatIf
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllByName')]
        [Alias('Group', 'DisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllById')]
        [string]$GroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Device', 'DeviceName', 'Id', 'DeviceId', 'Name')]
        [string[]]$DeviceIdentifier,

        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllByName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllById')]
        [switch]$ClearAll,

        [switch]$TreatInputAsId,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $devices = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not $graphConnected) { return }
        if ($ClearAll.IsPresent) { return }

        foreach ($entry in $DeviceIdentifier) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$devices.Add($entry.Trim())
            }
        }
    }

    end {
        if (-not $graphConnected) { return }
        if (-not $ClearAll.IsPresent -and $devices.Count -eq 0) {
            Write-NCMessage "No devices were specified." -Level WARNING
            return
        }

        $resolvedGroup = $null
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                $resolvedGroup = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Entra group with ID '$GroupId' not found: $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            $escapedName = $GroupName.Replace("'", "''")
            try {
                $resolvedGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-NCMessage "Unable to resolve group '$GroupName': $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $resolvedGroup) {
                try {
                    $resolvedGroup = Get-MgGroup -GroupId $GroupName -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Entra group '$GroupName' not found by name or ID" -Level ERROR
                    return
                }
            }
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $devicesToRemove = [System.Collections.Generic.List[object]]::new()

        if ($ClearAll.IsPresent) {
            $confirmMessage = "You are about to remove ALL device members from '$($resolvedGroup.DisplayName)'. This is a high-risk operation."
            if (-not $PSCmdlet.ShouldContinue($confirmMessage, "Confirm ClearAll")) {
                Write-NCMessage "ClearAll operation cancelled." -Level WARNING
                return
            }

            try {
                $members = @(Get-MgGroupMember -GroupId $resolvedGroup.Id -All -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to read members for group $($resolvedGroup.DisplayName): $($_.Exception.Message)" -Level ERROR
                return
            }

            $resolveType = {
                param($odataType)
                if ([string]::IsNullOrWhiteSpace($odataType)) {
                    return 'DirectoryObject'
                }

                $value = $odataType.ToLowerInvariant()
                if ($value -match 'user') { return 'User' }
                if ($value -match 'device') { return 'Device' }
                if ($value -match 'group') { return 'Group' }
                if ($value -match 'serviceprincipal') { return 'ServicePrincipal' }
                if ($value -match 'orgcontact') { return 'Contact' }
                return 'DirectoryObject'
            }

            foreach ($member in $members) {
                $props = if ($member.AdditionalProperties) { $member.AdditionalProperties } else { @{} }
                $odataType = if ($props.ContainsKey('@odata.type')) { $props['@odata.type'] } else { $null }
                $memberType = & $resolveType $odataType
                if ($memberType -ne 'Device') {
                    continue
                }

                $label = if ($props.ContainsKey('displayName')) {
                    $props.displayName
                }
                else {
                    $member.Id
                }

                $devicesToRemove.Add([pscustomobject]@{
                        Id    = $member.Id
                        Label = $label
                    }) | Out-Null
            }

            if ($devicesToRemove.Count -eq 0) {
                Write-NCMessage "No device members found for $($resolvedGroup.DisplayName)." -Level WARNING
                return
            }
        }
        else {
            $uniqueDevices = @($devices | Select-Object -Unique)
            foreach ($target in @(Resolve-NCEntraDeviceTargetBatch -DeviceIdentifier $uniqueDevices -TreatInputAsId:$TreatInputAsId)) {
                $devicesToRemove.Add([pscustomobject]@{
                        Id    = $target.Id
                        Label = $target.Label
                    }) | Out-Null
            }
        }

        if ($devicesToRemove.Count -gt 0) {
            Write-NCMessage "Processing $($devicesToRemove.Count) device(s) in Graph batches (20 per request) ..." -Level INFO
        }
        $groupPath = "/groups/$([uri]::EscapeDataString([string]$resolvedGroup.Id))"

        for ($offset = 0; $offset -lt $devicesToRemove.Count; $offset += 20) {
            $chunk = @($devicesToRemove[$offset..([Math]::Min($offset + 20, $devicesToRemove.Count) - 1)])

            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $chunk) {
                if ($PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Remove device '$($entry.Label)'")) {
                    $approved.Add($entry)
                }
            }
            if ($approved.Count -eq 0) { continue }

            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'DELETE'; Url = "$groupPath/members/$([uri]::EscapeDataString([string]$approved[$i].Id))/`$ref" }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Removing devices from $($resolvedGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $deviceId = $approved[$i].Id
                $deviceLabel = $approved[$i].Label
                $response = $responses[$i]
                $status = 'Removed'

                if ($response.Success) {
                    Write-NCMessage "Removed device '$deviceLabel' from group '$($resolvedGroup.DisplayName)'" -Level SUCCESS
                }
                elseif ($response.Status -eq 404 -or $response.ErrorMessage -match 'could not find member' -or $response.ErrorMessage -match 'does not exist') {
                    $status = 'NotFound'
                    Write-NCMessage "Device '$deviceLabel' is not a member of '$($resolvedGroup.DisplayName)'" -Level WARNING
                }
                else {
                    $status = 'Failed'
                    Write-NCMessage "Failed to remove device '$deviceLabel' from '$($resolvedGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }

                if ($PassThru.IsPresent) {
                    $results.Add([pscustomobject][ordered]@{
                            GroupName  = $resolvedGroup.DisplayName
                            GroupId    = $resolvedGroup.Id
                            MemberName = $deviceLabel
                            MemberId   = $deviceId
                            MemberType = 'Device'
                            Status     = $status
                        }) | Out-Null
                }
            }
        }

        if ($PassThru.IsPresent -and $results.Count -gt 0) {
            $results
        }
    }
}

function Remove-EntraGroupUser {
    <#
    .SYNOPSIS
        Removes one or more users from an Entra group.
    .DESCRIPTION
        Connects to Microsoft Graph, resolves the target group by display name or ID, then removes
        the provided users by UPN/display name or object ID. Accepts pipeline input for users.
    .PARAMETER GroupName
        Display name of the Entra group.
    .PARAMETER GroupId
        Object ID of the Entra group.
    .PARAMETER UserIdentifier
        User principal name, display name, or object ID. Accepts pipeline input and common Id/DisplayName property names.
    .PARAMETER ClearAll
        Remove all user members from the Entra group (devices and other objects are left untouched).
    .PARAMETER TreatInputAsId
        Treat every UserIdentifier as an object ID without attempting name resolution.
    .PARAMETER PassThru
        Emit a summary object for each processed user.
    .EXAMPLE
        "user1@contoso.com","user2@contoso.com" | Remove-EntraGroupUser -GroupName "My Entra Group"
    .EXAMPLE
        Remove-EntraGroupUser -GroupId "00000000-0000-0000-0000-000000000000" -UserIdentifier "user1@contoso.com"
    .EXAMPLE
        Remove-EntraGroupUser -GroupName "My Entra Group" -ClearAll
    .EXAMPLE
        Remove-EntraGroupUser -GroupName "My Entra Group" -ClearAll -WhatIf
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllByName')]
        [Alias('Group', 'DisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllById')]
        [string]$GroupId,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('User', 'UPN', 'Mail', 'Id', 'UserId')]
        [string[]]$UserIdentifier,

        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllByName')]
        [Parameter(Mandatory = $true, ParameterSetName = 'ClearAllById')]
        [switch]$ClearAll,

        [switch]$TreatInputAsId,
        [switch]$PassThru
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }

        $users = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if (-not $graphConnected) { return }
        if ($ClearAll.IsPresent) { return }

        foreach ($entry in $UserIdentifier) {
            if (-not [string]::IsNullOrWhiteSpace($entry)) {
                [void]$users.Add($entry.Trim())
            }
        }
    }

    end {
        if (-not $graphConnected) { return }
        if (-not $ClearAll.IsPresent -and $users.Count -eq 0) {
            Write-NCMessage "No users were specified." -Level WARNING
            return
        }

        $resolvedGroup = $null
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                $resolvedGroup = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
            }
            catch {
                Write-NCMessage "Entra group with ID '$GroupId' not found: $($_.Exception.Message)" -Level ERROR
                return
            }
        }
        else {
            $escapedName = $GroupName.Replace("'", "''")
            try {
                $resolvedGroup = Get-MgGroup -Filter "displayName eq '$escapedName'" -All -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-NCMessage "Unable to resolve group '$GroupName': $($_.Exception.Message)" -Level ERROR
                return
            }

            if (-not $resolvedGroup) {
                try {
                    $resolvedGroup = Get-MgGroup -GroupId $GroupName -ErrorAction Stop
                }
                catch {
                    Write-NCMessage "Entra group '$GroupName' not found by name or ID" -Level ERROR
                    return
                }
            }
        }

        $results = [System.Collections.Generic.List[object]]::new()
        $usersToRemove = [System.Collections.Generic.List[object]]::new()

        if ($ClearAll.IsPresent) {
            $confirmMessage = "You are about to remove ALL user members from '$($resolvedGroup.DisplayName)'. This is a high-risk operation."
            if (-not $PSCmdlet.ShouldContinue($confirmMessage, "Confirm ClearAll")) {
                Write-NCMessage "ClearAll operation cancelled." -Level WARNING
                return
            }

            try {
                $members = @(Get-MgGroupMember -GroupId $resolvedGroup.Id -All -ErrorAction Stop)
            }
            catch {
                Write-NCMessage "Unable to read members for group $($resolvedGroup.DisplayName): $($_.Exception.Message)" -Level ERROR
                return
            }

            $resolveType = {
                param($odataType)
                if ([string]::IsNullOrWhiteSpace($odataType)) {
                    return 'DirectoryObject'
                }

                $value = $odataType.ToLowerInvariant()
                if ($value -match 'user') { return 'User' }
                if ($value -match 'device') { return 'Device' }
                if ($value -match 'group') { return 'Group' }
                if ($value -match 'serviceprincipal') { return 'ServicePrincipal' }
                if ($value -match 'orgcontact') { return 'Contact' }
                return 'DirectoryObject'
            }

            foreach ($member in $members) {
                $props = if ($member.AdditionalProperties) { $member.AdditionalProperties } else { @{} }
                $odataType = if ($props.ContainsKey('@odata.type')) { $props['@odata.type'] } else { $null }
                $memberType = & $resolveType $odataType
                if ($memberType -ne 'User') {
                    continue
                }

                $label = if ($props.ContainsKey('userPrincipalName')) {
                    $props.userPrincipalName
                }
                elseif ($props.ContainsKey('displayName')) {
                    $props.displayName
                }
                else {
                    $member.Id
                }

                $usersToRemove.Add([pscustomobject]@{
                        Id    = $member.Id
                        Label = $label
                    }) | Out-Null
            }

            if ($usersToRemove.Count -eq 0) {
                Write-NCMessage "No user members found for $($resolvedGroup.DisplayName)." -Level WARNING
                return
            }
        }
        else {
            $uniqueUsers = @($users | Select-Object -Unique)
            foreach ($target in @(Resolve-NCEntraGroupUserTarget -UserIdentifier $uniqueUsers -TreatInputAsId:$TreatInputAsId)) {
                $usersToRemove.Add([pscustomobject]@{
                        Id    = $target.Id
                        Label = $target.Label
                    }) | Out-Null
            }
        }

        Write-NCMessage "Processing $($usersToRemove.Count) user(s) in Graph batches (20 per request) ..." -Level INFO

        for ($offset = 0; $offset -lt $usersToRemove.Count; $offset += 20) {
            $chunk = @($usersToRemove[$offset..([Math]::Min($offset + 20, $usersToRemove.Count) - 1)])

            $approved = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in $chunk) {
                if ($PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Remove user '$($entry.Label)'")) {
                    $approved.Add($entry)
                }
            }
            if ($approved.Count -eq 0) { continue }

            $requests = @(for ($i = 0; $i -lt $approved.Count; $i++) {
                    @{ Id = "m$i"; Method = 'DELETE'; Url = "/groups/$($resolvedGroup.Id)/members/$([uri]::EscapeDataString($approved[$i].Id))/`$ref" }
                })
            $responses = @(Invoke-NCGraphBatch -Requests $requests -Activity "Removing users from $($resolvedGroup.DisplayName)")

            for ($i = 0; $i -lt $approved.Count; $i++) {
                $userId = $approved[$i].Id
                $userLabel = $approved[$i].Label
                $response = $responses[$i]
                $status = 'Removed'

                if ($response.Success) {
                    Write-NCMessage "Removed user '$userLabel' from group '$($resolvedGroup.DisplayName)'." -Level SUCCESS
                }
                elseif ($response.Status -eq 404 -or $response.ErrorMessage -match 'could not find member' -or $response.ErrorMessage -match 'does not exist') {
                    $status = 'NotFound'
                    Write-NCMessage "User '$userLabel' is not a member of '$($resolvedGroup.DisplayName)'" -Level WARNING
                }
                else {
                    $status = 'Failed'
                    Write-NCMessage "Failed to remove user '$userLabel' from '$($resolvedGroup.DisplayName)': $($response.ErrorMessage)" -Level ERROR
                }

                if ($PassThru.IsPresent) {
                    $results.Add([pscustomobject][ordered]@{
                            GroupName  = $resolvedGroup.DisplayName
                            GroupId    = $resolvedGroup.Id
                            MemberName = $userLabel
                            MemberId   = $userId
                            MemberType = 'User'
                            Status     = $status
                        }) | Out-Null
                }
            }
        }

        if ($PassThru.IsPresent -and $results.Count -gt 0) {
            $results
        }
    }
}

function Search-EntraGroup {
    <#
    .SYNOPSIS
        Finds Entra groups by display name or description.
    .DESCRIPTION
        Uses Microsoft Graph search to find groups by display name or description and returns
        key properties for easy identification.
    .PARAMETER SearchText
        Text to search for in display name and/or description. Accepts pipeline input.
    .PARAMETER SearchIn
        Where to search: DisplayName, Description, or Any (both).
    .PARAMETER GridView
        Show additional details in Out-GridView instead of returning objects.
    .EXAMPLE
        Search-EntraGroup -SearchText "java"
    .EXAMPLE
        Search-EntraGroup -SearchText "legacy apps" -SearchIn Description
    .EXAMPLE
        "marketing" | Search-EntraGroup -SearchIn Any -GridView
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Search', 'Query', 'Text', 'Name', 'DisplayName', 'Description')]
        [string]$SearchText,

        [ValidateSet('DisplayName', 'Description', 'Any')]
        [string]$SearchIn = 'DisplayName',

        [switch]$GridView
    )

    begin {
        $graphConnected = Test-MgGraphConnection -Scopes @('Group.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphConnected) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }
    }

    process {
        if (-not $graphConnected) {
            return
        }

        if ([string]::IsNullOrWhiteSpace($SearchText)) {
            Write-NCMessage "SearchText cannot be empty." -Level WARNING
            return
        }

        $escapedText = $SearchText.Replace('"', '""').Trim()
        if ([string]::IsNullOrWhiteSpace($escapedText)) {
            Write-NCMessage "SearchText cannot be empty." -Level WARNING
            return
        }

        $groups = @()

        try {
            switch ($SearchIn) {
                'DisplayName' {
                    $searchClause = "`"displayName:$escapedText`""
                    $groups = @(Get-MgGroup -Search $searchClause -ConsistencyLevel eventual -CountVariable count -All -ErrorAction Stop)
                }
                'Description' {
                    $searchClause = "`"description:$escapedText`""
                    $groups = @(Get-MgGroup -Search $searchClause -ConsistencyLevel eventual -CountVariable count -All -ErrorAction Stop)
                }
                'Any' {
                    $searchDisplay = "`"displayName:$escapedText`""
                    $searchDescription = "`"description:$escapedText`""
                    $byName = @(Get-MgGroup -Search $searchDisplay -ConsistencyLevel eventual -CountVariable countName -All -ErrorAction Stop)
                    $byDescription = @(Get-MgGroup -Search $searchDescription -ConsistencyLevel eventual -CountVariable countDesc -All -ErrorAction Stop)
                    $groups = @($byName + $byDescription | Sort-Object Id -Unique)
                }
            }
        }
        catch {
            Write-NCMessage "Unable to search groups with '$SearchText': $($_.Exception.Message)" -Level ERROR
            return
        }

        Add-EmptyLine
        Write-Verbose "Groups found: $($groups.Count) for '$SearchText'."

        if (-not $groups -or $groups.Count -eq 0) {
            Write-NCMessage "No groups found for '$SearchText'." -Level WARNING
            return
        }

        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($group in $groups) {
            $row = [ordered]@{
                'Group Name'        = $group.DisplayName
                'Group Id'          = $group.Id
                'Group Description' = $group.Description
            }

            if ($GridView.IsPresent) {
                $row['Group Mail Nickname'] = $group.MailNickname
                $row['Group Mail Enabled'] = $group.MailEnabled
                $row['Group Security Enabled'] = $group.SecurityEnabled
                $row['Group Types'] = if ($group.GroupTypes) { ($group.GroupTypes -join ', ') } else { $null }
            }

            $results.Add([pscustomobject]$row) | Out-Null
        }

        if ($GridView.IsPresent) {
            $results | Out-GridView -Title "Entra Groups - Search: $SearchText"
        }
        else {
            $results | Sort-Object 'Group Name'
        }
    }
}

function Set-EntraGroupDescription {
    <#
    .SYNOPSIS
        Updates the description of an Entra group.
    .DESCRIPTION
        Resolves an Entra group by display name or object ID and updates only the description
        property in Microsoft Graph.
    .PARAMETER GroupName
        Display name of the target Entra group.
    .PARAMETER GroupId
        Object ID of the target Entra group.
    .PARAMETER Description
        New group description. Pass an empty string to clear the description.
    .PARAMETER PassThru
        Emit the updated group object.
    .EXAMPLE
        Set-EntraGroupDescription -GroupName "GitLab-Prod" -Description "Production GitLab access group"
    .EXAMPLE
        Set-EntraGroupDescription -GroupId "00000000-0000-0000-0000-000000000000" -Description "" -PassThru
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Alias('DisplayName', 'Name')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
        [string]$GroupId,

        [Parameter(Mandatory = $true, Position = 1)]
        [AllowEmptyString()]
        [string]$Description,

        [switch]$PassThru
    )

    $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
    if (-not $graphConnected) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return
    }

    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
        return
    }

    $resolvedGroup = if ($PSCmdlet.ParameterSetName -eq 'ById') {
        Resolve-NCEntraGroup -GroupName $GroupId -GroupId $GroupId
    }
    else {
        Resolve-NCEntraGroup -GroupName $GroupName
    }

    if (-not $resolvedGroup) {
        return
    }

    if (-not $PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, 'Update group description')) {
        return
    }

    try {
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($resolvedGroup.Id)" -Method PATCH -Body (@{ description = $Description } | ConvertTo-Json -Depth 10) -ContentType 'application/json' | Out-Null
        Write-NCMessage "Updated description for group '$($resolvedGroup.DisplayName)'." -Level SUCCESS

        if ($PassThru.IsPresent) {
            Get-MgGroup -GroupId $resolvedGroup.Id -Property @(
                'id',
                'displayName',
                'description',
                'groupTypes',
                'mailEnabled',
                'mailNickname',
                'securityEnabled',
                'visibility',
                'onPremisesSyncEnabled',
                'isAssignableToRole'
            )
        }
    }
    catch {
        Write-NCMessage "Failed to update description for group '$($resolvedGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
    }
}

function Set-EntraGroupDisplayName {
    <#
    .SYNOPSIS
        Updates the display name of an Entra group.
    .DESCRIPTION
        Resolves an Entra group by display name or object ID and updates only the displayName
        property in Microsoft Graph.
    .PARAMETER GroupName
        Current display name of the target Entra group.
    .PARAMETER GroupId
        Object ID of the target Entra group.
    .PARAMETER DisplayName
        New display name for the group.
    .PARAMETER PassThru
        Emit the updated group object.
    .EXAMPLE
        Set-EntraGroupDisplayName -GroupName "GitLab-Prod" -DisplayName "GitLab - Production"
    .EXAMPLE
        Set-EntraGroupDisplayName -GroupId "00000000-0000-0000-0000-000000000000" -DisplayName "GitLab - Production" -PassThru
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [Alias('CurrentName', 'CurrentDisplayName')]
        [string]$GroupName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
        [string]$GroupId,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$DisplayName,

        [switch]$PassThru
    )

    $graphConnected = Test-MgGraphConnection -Scopes @('Group.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
    if (-not $graphConnected) {
        Add-EmptyLine
        Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        return
    }

    if (-not (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-NCMessage "Invoke-MgGraphRequest is not available in the current Microsoft Graph session." -Level ERROR
        return
    }

    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        Write-NCMessage "DisplayName cannot be empty." -Level WARNING
        return
    }

    $resolvedGroup = if ($PSCmdlet.ParameterSetName -eq 'ById') {
        Resolve-NCEntraGroup -GroupName $GroupId -GroupId $GroupId
    }
    else {
        Resolve-NCEntraGroup -GroupName $GroupName
    }

    if (-not $resolvedGroup) {
        return
    }

    if ($resolvedGroup.OnPremisesSyncEnabled -eq $true) {
        Write-NCMessage "Group '$($resolvedGroup.DisplayName)' is synchronized from on-premises AD and cannot be renamed directly in Entra." -Level ERROR
        return
    }

    if ($resolvedGroup.DisplayName -eq $DisplayName) {
        Write-NCMessage "Group '$($resolvedGroup.DisplayName)' already has that display name." -Level WARNING
        if ($PassThru.IsPresent) {
            Get-MgGroup -GroupId $resolvedGroup.Id -Property @(
                'id',
                'displayName',
                'description',
                'groupTypes',
                'mailEnabled',
                'mailNickname',
                'securityEnabled',
                'visibility',
                'onPremisesSyncEnabled',
                'isAssignableToRole'
            )
        }
        return
    }

    if (-not $PSCmdlet.ShouldProcess($resolvedGroup.DisplayName, "Rename group to '$DisplayName'")) {
        return
    }

    try {
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($resolvedGroup.Id)" -Method PATCH -Body (@{ displayName = $DisplayName } | ConvertTo-Json -Depth 10) -ContentType 'application/json' | Out-Null
        Write-NCMessage "Updated display name for group '$($resolvedGroup.DisplayName)' to '$DisplayName'." -Level SUCCESS

        if ($PassThru.IsPresent) {
            Get-MgGroup -GroupId $resolvedGroup.Id -Property @(
                'id',
                'displayName',
                'description',
                'groupTypes',
                'mailEnabled',
                'mailNickname',
                'securityEnabled',
                'visibility',
                'onPremisesSyncEnabled',
                'isAssignableToRole'
            )
        }
    }
    catch {
        Write-NCMessage "Failed to update display name for group '$($resolvedGroup.DisplayName)': $($_.Exception.Message)" -Level ERROR
    }
}
