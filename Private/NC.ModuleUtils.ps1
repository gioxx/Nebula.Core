#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Module's Utilities =========================================================================================================

function Invoke-NCRetry {
    <#
    .SYNOPSIS
        Executes a scriptblock with retry logic.
    .DESCRIPTION
        Runs the provided block up to MaxAttempts, invoking OnError between retries.
        Throws the last error once all attempts are exhausted.
    .PARAMETER Action
        Script block to execute.
    .PARAMETER MaxAttempts
        Maximum number of attempts before throwing (default 3).
    .PARAMETER DelaySeconds
        Pause between attempts (default 5 seconds).
    .PARAMETER OperationDescription
        Friendly description used in log messages.
    .PARAMETER OnError
        Optional callback invoked after each failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,
        [ValidateRange(1, [int]::MaxValue)]
        [int]$MaxAttempts = 3,
        [ValidateRange(0, [int]::MaxValue)]
        [int]$DelaySeconds = 5,
        [string]$OperationDescription = 'operation',
        [scriptblock]$OnError
    )

    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            return & $Action
        }
        catch {
            if ($OnError) {
                & $OnError -ArgumentList $attempt, $MaxAttempts, $_
            }
            else {
                Write-NCMessage "Operation '$OperationDescription' failed (attempt $attempt of $MaxAttempts). $($_.Exception.Message)" -Level ERROR
            }

            if ($attempt -ge $MaxAttempts) {
                throw
            }

            if ($DelaySeconds -gt 0) {
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
}

function New-File {
    <#
    .SYNOPSIS
        Generates a non-colliding file path.
    .DESCRIPTION
        Given a desired path, appends _N before the extension until an unused name is found.
    .PARAMETER Path
        Desired output file path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)
    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if (-not $directory) {
        $directory = (Get-Location).ProviderPath
    }

    $candidate = Join-Path -Path $directory -ChildPath ($baseName + $extension)
    $count = 1
    while (Test-Path -LiteralPath $candidate) {
        $fileName = "{0}_{1}{2}" -f $baseName, $count, $extension
        $candidate = Join-Path -Path $directory -ChildPath $fileName
        $count++
    }

    return $candidate
}

function Restore-ProgressAndInfoPreferences {
    <#
    .SYNOPSIS
        Restores Information/Progress preference variables.
    .DESCRIPTION
        Reverts preference variables previously captured by Set-ProgressAndInfoPreferences.
        No-ops if nothing was captured.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:PreferencesCaptured) {
        return
    }

    if ($null -ne $script:PreviousInformationPreference) {
        Set-Variable -Name InformationPreference -Value $script:PreviousInformationPreference -Scope Global
    }

    if ($null -ne $script:PreviousProgressPreference) {
        Set-Variable -Name ProgressPreference -Value $script:PreviousProgressPreference -Scope Global
    }

    $script:PreferencesCaptured = $false
    $script:PreviousInformationPreference = $null
    $script:PreviousProgressPreference = $null
}

function Set-ProgressAndInfoPreferences {
    <#
    .SYNOPSIS
        Forces Information/Progress preference variables to Continue.
    .DESCRIPTION
        Saves current preference values (once per session) and sets global
        InformationPreference and ProgressPreference to Continue for verbose output.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:PreferencesCaptured) {
        $script:PreviousInformationPreference = $InformationPreference
        $script:PreviousProgressPreference = $ProgressPreference
        $script:PreferencesCaptured = $true
    }

    Set-Variable -Name InformationPreference -Value Continue -Scope Global
    Set-Variable -Name ProgressPreference -Value Continue -Scope Global
}

function Show-Table {
    <#
    .SYNOPSIS
        Outputs a table of rows.
    .DESCRIPTION
        Outputs a table of rows with a title.
    .PARAMETER Rows
        Table rows.
    .PARAMETER AsTable
        Output as a table.
    #>
    [CmdletBinding()]
    param(
        [array]$Rows,
        [switch]$AsTable
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-NCMessage "(none)" -Level INFO
        return
    }

    if ($AsTable) {
        $Rows | Format-Table -AutoSize
    }
    else {
        $Rows | Format-List
    }
}

function Test-Folder {
    <#
    .SYNOPSIS
        Normalizes and validates a folder path.
    .DESCRIPTION
        Returns the current directory when input is blank, trims trailing separators,
        and throws if the path is invalid.
    .PARAMETER Path
        Folder path to validate (optional).
    #>
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return (Get-Location).ProviderPath
    }

    $normalized = $Path.TrimEnd('\')
    try {
        return [System.IO.Path]::GetFullPath($normalized)
    }
    catch {
        throw "Invalid folder path '$Path'. $($_.Exception.Message)"
    }
}