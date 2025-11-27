#Requires -Version 5.0

function Get-NebulaConfig {
    <#
    .SYNOPSIS
        Displays the effective Nebula.Core configuration.
    .DESCRIPTION
        Lists active configuration values, environment overrides, and (when present) differences introduced
        by machine or user custom configuration files.
    #>
    [CmdletBinding()]
    param()

    $info = $script:NebulaConfigInfo
    $summary = [pscustomobject]@{
        ModuleRoot         = $script:ModuleRoot
        UserConfigPath     = $info.UserConfigPath
        UserConfigExists   = $info.UserConfigExists
        UserConfigLoaded   = $info.UserConfigLoaded
        MachineConfigPath  = $info.MachineConfigPath
        MachineConfigExists= $info.MachineConfigExists
        MachineConfigLoaded= $info.MachineConfigLoaded
    }

    # $envRows = foreach ($key in $info.EnvironmentOverrideKeys) {
    #     $envVar = "NEBULA_{0}" -f ($key.ToUpperInvariant())
    #     [pscustomobject]@{
    #         Key    = $key
    #         EnvVar = $envVar
    #         Value  = [Environment]::GetEnvironmentVariable($envVar)
    #     }
    # }

    $configRows = foreach ($key in ($script:NC_Config.Keys | Sort-Object)) {
        [pscustomobject]@{
            Key   = $key
            Value = $script:NC_Config[$key]
        }
    }

    $licenseRows = foreach ($source in $script:NCLicenseSources.Keys) {
        $src = $script:NCLicenseSources[$source]
        [pscustomobject]@{
            Source    = $source
            CacheFile = $src.CacheFileName
            FileUrl   = $src.FileUrl
        }
    }

    $summary | Format-List
    Show-Table -Rows $configRows -AsTable
    Show-Table -Rows $licenseRows -AsTable

    $defaults = $script:NC_Defaults

    if ($info.MachineConfigLoaded -and $info.MachineOverrideKeys.Count -gt 0) {
        $machineRows = foreach ($key in ($info.MachineOverrideKeys | Sort-Object)) {
            [pscustomobject]@{
                Key          = $key
                DefaultValue = $defaults[$key]
                CurrentValue = $script:NC_Config[$key]
            }
        }
        Show-Table -Rows $machineRows
    }

    if ($info.UserConfigLoaded -and $info.UserOverrideKeys.Count -gt 0) {
        $userRows = foreach ($key in ($info.UserOverrideKeys | Sort-Object)) {
            [pscustomobject]@{
                Key          = $key
                DefaultValue = $defaults[$key]
                CurrentValue = $script:NC_Config[$key]
            }
        }
        Show-Table -Rows $userRows -AsTable
    }

    # return [pscustomobject]@{
    #     Summary              = $summary
    #     EnvironmentOverrides = $envRows
    #     ActiveConfiguration  = $configRows
    #     LicenseSources       = $licenseRows
    # }
}

function Sync-NebulaConfig {
    <#
    .SYNOPSIS
        Reloads Nebula.Core configuration without re-importing the module.
    .DESCRIPTION
        Re-runs the initialization logic so that changes to machine/user PSD1 files or environment variables
        take effect immediately in the current session.
    #>
    [CmdletBinding()]
    param()

    try {
        Initialize-NebulaConfig
        Write-NCMessage "Nebula configuration reloaded." -Level SUCCESS
    }
    catch {
        Write-NCMessage "Unable to reload configuration. $($_.Exception.Message)" -Level ERROR
        throw
    }
}
