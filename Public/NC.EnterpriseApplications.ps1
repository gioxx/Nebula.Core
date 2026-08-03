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
