#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Configuration bootstrap ==============================================================================================================

function Import-NCConfigurationFile {
    [CmdletBinding()]
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{}
    }

    try {
        $data = Import-PowerShellDataFile -LiteralPath $Path
        if ($data -isnot [hashtable]) {
            throw "Configuration file '$Path' does not contain a hashtable."
        }
        return $data
    }
    catch {
        Write-NCMessage "Failed to load Nebula.Core configuration from '$Path'. $($_.Exception.Message)" -Level ERROR
        return @{}
    }
}

function Merge-NCConfig {
    [CmdletBinding()]
    param(
        [hashtable]$Base,
        [hashtable]$Override
    )

    foreach ($key in $Override.Keys) {
        $Base[$key] = $Override[$key]
    }

    return $Base
}

if (-not $script:NCLicenseSources) {
    $script:NCLicenseSources = @{
        Primary = @{
            CacheFileName = 'M365_licenses.json'
            ApiUrl        = "https://api.github.com/repos/gioxx/Nebula.Core/commits?path=JSON/M365_licenses.json"
            FileUrl       = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/JSON/M365_licenses.json'
        }
        Custom  = @{
            CacheFileName = 'M365_licenses_custom.json'
            ApiUrl        = "https://api.github.com/repos/gioxx/Nebula.Core/commits?path=JSON/M365_licenses_custom.json"
            FileUrl       = 'https://raw.githubusercontent.com/gioxx/Nebula.Core/main/JSON/M365_licenses_custom.json'
        }
    }
}

if (-not $script:NC_Defaults) {
    $script:NC_Defaults = [ordered]@{
        CSV_DefaultLimiter      = ";"
        CSV_Encoding            = 'UTF-8'
        CheckUpdatesOnConnect   = $true
        CheckUpdatesIntervalHours = 24
        DateTimeString_CSV      = 'yyyyMMdd'
        DateTimeString_Full     = 'yyy-MM-dd HH:mm:ss'
        LicenseCacheDays        = 7
        LicenseCacheDirectory   = (Join-Path $env:USERPROFILE '.NebulaCore\Cache')
        MaxFieldLength          = 35
        UsageLocation           = 'US'
        UserConfigRoot          = (Join-Path $env:USERPROFILE '.NebulaCore')
    }
}

function Initialize-NebulaConfig {
    $config = [ordered]@{}
    foreach ($key in $script:NC_Defaults.Keys) {
        $config[$key] = $script:NC_Defaults[$key]
    }

    $userConfigPath = Join-Path -Path $script:NC_Defaults.UserConfigRoot -ChildPath 'settings.psd1'
    $machineConfigPath = Join-Path -Path $env:ProgramData -ChildPath 'Nebula.Core\settings.psd1'
    $userConfigExists = Test-Path -LiteralPath $userConfigPath
    $machineConfigExists = Test-Path -LiteralPath $machineConfigPath

    $userConfig = Import-NCConfigurationFile -Path $userConfigPath
    $machineConfig = Import-NCConfigurationFile -Path $machineConfigPath

    $machineOverrideKeys = [System.Collections.Generic.HashSet[string]]::new()
    if ($machineConfig.Count) {
        foreach ($key in $machineConfig.Keys) {
            $machineOverrideKeys.Add($key) | Out-Null
        }
        $config = Merge-NCConfig -Base $config -Override $machineConfig
    }

    $userOverrideKeys = [System.Collections.Generic.HashSet[string]]::new()
    if ($userConfig.Count) {
        foreach ($key in $userConfig.Keys) {
            $userOverrideKeys.Add($key) | Out-Null
        }
        $config = Merge-NCConfig -Base $config -Override $userConfig
    }

    $envOverrideKeys = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($key in $config.Keys) {
        $envVarName = "NEBULA_{0}" -f ($key.ToUpperInvariant())
        $envValue = [Environment]::GetEnvironmentVariable($envVarName)
        if ($envValue) {
            $config[$key] = $envValue
            $envOverrideKeys.Add($key) | Out-Null
        }
    }

    $script:NC_Config = $config
    New-Variable -Name NCVars -Value $config -Scope Script -Force

    $script:NebulaConfigInfo = [ordered]@{
        UserConfigPath          = $userConfigPath
        UserConfigExists        = $userConfigExists
        UserConfigLoaded        = $userConfig.Count -gt 0
        UserOverrideKeys        = @($userOverrideKeys)
        MachineConfigPath       = $machineConfigPath
        MachineConfigExists     = $machineConfigExists
        MachineConfigLoaded     = $machineConfig.Count -gt 0
        MachineOverrideKeys     = @($machineOverrideKeys)
        EnvironmentOverrideKeys = @($envOverrideKeys)
    }
}
