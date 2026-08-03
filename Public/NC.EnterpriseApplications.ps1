#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Enterprise Application lifecycle cmdlets =============================================================================================

function Export-EnterpriseApplication {
    <#
    .SYNOPSIS
        Exports an Enterprise Application (Application + Service Principal) to a JSON snapshot file.
    .DESCRIPTION
        Reads the source Application, its Service Principal, owners, and (optionally) App Role
        Assignments, then writes a normalized, versioned snapshot to -OutputPath. Client secret and
        certificate values are never read; only metadata (display name, key ID, expiry) is captured.
    .PARAMETER ApplicationName
        Display name of the source Enterprise Application.
    .PARAMETER ApplicationId
        Object ID of the source Application.
    .PARAMETER OutputPath
        Destination JSON file path.
    .PARAMETER IncludeAppRoleAssignments
        Also export App Role Assignments (users/groups assigned to the app).
    .PARAMETER Force
        Overwrite -OutputPath if it already exists.
    .EXAMPLE
        Export-EnterpriseApplication -ApplicationName "Contoso Test App" -OutputPath .\contoso-test-app.json
    .EXAMPLE
        Export-EnterpriseApplication -ApplicationName "Contoso Test App" -OutputPath .\contoso-test-app.json -IncludeAppRoleAssignments -Force
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [string]$ApplicationName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [string]$ApplicationId,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$OutputPath,

        [switch]$IncludeAppRoleAssignments,
        [switch]$Force
    )

    begin {
        $graphReady = Test-MgGraphConnection -Scopes @('Application.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphReady) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }
    }

    process {
        if (-not $graphReady) {
            return
        }

        if ((Test-Path -LiteralPath $OutputPath) -and -not $Force.IsPresent) {
            Write-NCMessage "Output file '$OutputPath' already exists. Use -Force to overwrite." -Level ERROR
            return
        }

        $snapshotParams = @{ IncludeAppRoleAssignments = $IncludeAppRoleAssignments.IsPresent }
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            $snapshotParams.ApplicationId = $ApplicationId
        }
        else {
            $snapshotParams.ApplicationName = $ApplicationName
        }

        $snapshot = Get-NCEnterpriseApplicationSnapshot @snapshotParams
        if (-not $snapshot) {
            return
        }

        try {
            $snapshot | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutputPath -Encoding UTF8 -ErrorAction Stop
            Write-NCMessage "Exported Enterprise Application '$($snapshot.Application.DisplayName)' to '$OutputPath'." -Level SUCCESS
        }
        catch {
            Write-NCMessage "Failed to write snapshot to '$OutputPath': $($_.Exception.Message)" -Level ERROR
        }
    }
}

function Import-EnterpriseApplication {
    <#
    .SYNOPSIS
        Creates or updates an Enterprise Application from a JSON snapshot file.
    .DESCRIPTION
        Reads a snapshot produced by Export-EnterpriseApplication and applies its properties to the
        Enterprise Application identified by -TargetDisplayName, creating it if it does not exist.
    .PARAMETER InputPath
        Path to the JSON snapshot file.
    .PARAMETER TargetDisplayName
        Display name of the destination Enterprise Application. Created if missing, updated if it exists.
    .PARAMETER IncludeAppRoleAssignments
        Also apply App Role Assignments captured in the snapshot.
    .PARAMETER PassThru
        Emit the apply-result summary object.
    .EXAMPLE
        Import-EnterpriseApplication -InputPath .\contoso-test-app.json -TargetDisplayName "Contoso Prod App"
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$InputPath,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$TargetDisplayName,

        [switch]$IncludeAppRoleAssignments,
        [switch]$PassThru
    )

    begin {
        $graphReady = Test-MgGraphConnection -Scopes @('Application.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphReady) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }
    }

    process {
        if (-not $graphReady) {
            return
        }

        if (-not (Test-Path -LiteralPath $InputPath)) {
            Write-NCMessage "Input file '$InputPath' not found." -Level ERROR
            return
        }

        try {
            $snapshot = Get-Content -LiteralPath $InputPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to read or parse snapshot '$InputPath': $($_.Exception.Message)" -Level ERROR
            return
        }

        if (-not $PSCmdlet.ShouldProcess($TargetDisplayName, "Import Enterprise Application from '$InputPath'")) {
            return
        }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName $TargetDisplayName -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent -Confirm:$false
        if ($PassThru.IsPresent) {
            $result
        }
    }
}

function Copy-EnterpriseApplication {
    <#
    .SYNOPSIS
        Clones an Enterprise Application directly into a new or existing destination, without an intermediate file.
    .DESCRIPTION
        Combines Get-NCEnterpriseApplicationSnapshot and Set-NCEnterpriseApplicationFromSnapshot to read the
        source Enterprise Application and apply it to -TargetDisplayName in one step.
    .PARAMETER SourceApplicationName
        Display name of the source Enterprise Application.
    .PARAMETER SourceApplicationId
        Object ID of the source Application.
    .PARAMETER TargetDisplayName
        Display name of the destination Enterprise Application. Created if missing, updated if it exists.
    .PARAMETER IncludeAppRoleAssignments
        Also copy App Role Assignments (users/groups assigned to the app).
    .PARAMETER PassThru
        Emit the apply-result summary object.
    .EXAMPLE
        Copy-EnterpriseApplication -SourceApplicationName "Contoso Test App" -TargetDisplayName "Contoso Prod App"
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName', Position = 0)]
        [string]$SourceApplicationName,

        [Parameter(Mandatory = $true, ParameterSetName = 'ById', Position = 0)]
        [string]$SourceApplicationId,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$TargetDisplayName,

        [switch]$IncludeAppRoleAssignments,
        [switch]$PassThru
    )

    begin {
        $graphReady = Test-MgGraphConnection -Scopes @('Application.ReadWrite.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
        if (-not $graphReady) {
            Add-EmptyLine
            Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
        }
    }

    process {
        if (-not $graphReady) {
            return
        }

        $snapshotParams = @{ IncludeAppRoleAssignments = $IncludeAppRoleAssignments.IsPresent }
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            $snapshotParams.ApplicationId = $SourceApplicationId
        }
        else {
            $snapshotParams.ApplicationName = $SourceApplicationName
        }

        $snapshot = Get-NCEnterpriseApplicationSnapshot @snapshotParams
        if (-not $snapshot) {
            return
        }

        if ($snapshot.Application.DisplayName -eq $TargetDisplayName) {
            Write-NCMessage "Source and destination Enterprise Applications are the same. Aborting." -Level ERROR
            return
        }

        if (-not $PSCmdlet.ShouldProcess($TargetDisplayName, "Clone Enterprise Application '$($snapshot.Application.DisplayName)' into '$TargetDisplayName'")) {
            return
        }

        $result = Set-NCEnterpriseApplicationFromSnapshot -Snapshot $snapshot -TargetDisplayName $TargetDisplayName -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent -Confirm:$false
        if ($PassThru.IsPresent) {
            $result
        }
    }
}

function Compare-EnterpriseApplication {
    <#
    .SYNOPSIS
        Diffs two Enterprise Applications, each given as a JSON snapshot file or a live application.
    .DESCRIPTION
        Loads the reference and difference sides (file or live Graph lookup, independently), diffs them,
        and returns the differing properties on the pipeline. Optionally writes a CSV or JSON report.
    .PARAMETER ReferencePath
        JSON snapshot file for the reference ("A") side.
    .PARAMETER ReferenceApplicationName
        Display name of a live application for the reference side.
    .PARAMETER ReferenceApplicationId
        Object ID of a live application for the reference side.
    .PARAMETER DifferencePath
        JSON snapshot file for the difference ("B") side.
    .PARAMETER DifferenceApplicationName
        Display name of a live application for the difference side.
    .PARAMETER DifferenceApplicationId
        Object ID of a live application for the difference side.
    .PARAMETER IncludeAppRoleAssignments
        Also compare App Role Assignments.
    .PARAMETER OutputReportPath
        Optional report file. Written as JSON if the path ends in .json, otherwise as CSV.
    .PARAMETER PassThru
        Also emit the diff rows on the pipeline even when -OutputReportPath is used (rows are emitted by default; this has no additional effect and exists for symmetry with the other cmdlets).
    .EXAMPLE
        Compare-EnterpriseApplication -ReferencePath .\contoso-test-app.json -DifferenceApplicationName "Contoso Prod App"
    .EXAMPLE
        Compare-EnterpriseApplication -ReferenceApplicationName "Contoso Test App" -DifferenceApplicationName "Contoso Prod App" -OutputReportPath .\diff.csv
    #>
    [CmdletBinding()]
    param(
        [string]$ReferencePath,
        [string]$ReferenceApplicationName,
        [string]$ReferenceApplicationId,

        [string]$DifferencePath,
        [string]$DifferenceApplicationName,
        [string]$DifferenceApplicationId,

        [switch]$IncludeAppRoleAssignments,
        [string]$OutputReportPath,
        [switch]$PassThru
    )

    begin {
        $referenceSourceCount = @($ReferencePath, $ReferenceApplicationName, $ReferenceApplicationId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
        $differenceSourceCount = @($DifferencePath, $DifferenceApplicationName, $DifferenceApplicationId | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count

        $inputValid = $true
        if ($referenceSourceCount -ne 1) {
            Write-NCMessage "Specify exactly one of -ReferencePath, -ReferenceApplicationName, or -ReferenceApplicationId." -Level ERROR
            $inputValid = $false
        }
        if ($differenceSourceCount -ne 1) {
            Write-NCMessage "Specify exactly one of -DifferencePath, -DifferenceApplicationName, or -DifferenceApplicationId." -Level ERROR
            $inputValid = $false
        }

        $graphReady = $true
        if ($inputValid -and ((-not $ReferencePath) -or (-not $DifferencePath))) {
            $graphReady = Test-MgGraphConnection -Scopes @('Application.Read.All', 'Directory.Read.All') -EnsureExchangeOnline:$false
            if (-not $graphReady) {
                Add-EmptyLine
                Write-NCMessage "Can't connect or use Microsoft Graph modules. Please check logs." -Level ERROR
            }
        }
    }

    process {
        if (-not $inputValid -or -not $graphReady) {
            return
        }

        $referenceSnapshot = if ($ReferencePath) {
            if (-not (Test-Path -LiteralPath $ReferencePath)) {
                Write-NCMessage "Reference file '$ReferencePath' not found." -Level ERROR
                return
            }
            Get-Content -LiteralPath $ReferencePath -Raw | ConvertFrom-Json
        }
        elseif ($ReferenceApplicationId) {
            Get-NCEnterpriseApplicationSnapshot -ApplicationId $ReferenceApplicationId -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent
        }
        else {
            Get-NCEnterpriseApplicationSnapshot -ApplicationName $ReferenceApplicationName -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent
        }

        if (-not $referenceSnapshot) {
            return
        }

        $differenceSnapshot = if ($DifferencePath) {
            if (-not (Test-Path -LiteralPath $DifferencePath)) {
                Write-NCMessage "Difference file '$DifferencePath' not found." -Level ERROR
                return
            }
            Get-Content -LiteralPath $DifferencePath -Raw | ConvertFrom-Json
        }
        elseif ($DifferenceApplicationId) {
            Get-NCEnterpriseApplicationSnapshot -ApplicationId $DifferenceApplicationId -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent
        }
        else {
            Get-NCEnterpriseApplicationSnapshot -ApplicationName $DifferenceApplicationName -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent
        }

        if (-not $differenceSnapshot) {
            return
        }

        $rows = Compare-NCEnterpriseApplicationSnapshot -ReferenceSnapshot $referenceSnapshot -DifferenceSnapshot $differenceSnapshot -IncludeAppRoleAssignments:$IncludeAppRoleAssignments.IsPresent

        if ($OutputReportPath) {
            try {
                if ($OutputReportPath -match '\.json$') {
                    $rows | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputReportPath -Encoding UTF8 -ErrorAction Stop
                }
                else {
                    $rows | Export-Csv -LiteralPath $OutputReportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                }
                Write-NCMessage "Wrote comparison report to '$OutputReportPath' ($($rows.Count) difference(s))." -Level SUCCESS
            }
            catch {
                Write-NCMessage "Failed to write comparison report to '$OutputReportPath': $($_.Exception.Message)" -Level ERROR
            }
        }

        if ($rows.Count -eq 0) {
            Write-NCMessage "No differences found." -Level SUCCESS
        }
        else {
            Write-NCMessage "Found $($rows.Count) difference(s)." -Level WARNING
        }

        $rows
    }
}
