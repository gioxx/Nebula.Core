# Nebula.Core.psm1
$script:ModuleRoot = $PSScriptRoot
$ncUserRoot  = Join-Path $env:USERPROFILE '.NebulaCore'
$ncCacheRoot = Join-Path $ncUserRoot 'Cache'

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

# --- Helper functions to load user's custom configuration ---
function Import-NCConfigurationFile {
    <#
    .SYNOPSIS
        Loads an override configuration file if present.
    .DESCRIPTION
        Attempts to import a PowerShell data file (.psd1) and returns a hashtable
        containing user-specified overrides. On failure, warns and returns an empty hashtable.
    .PARAMETER Path
        Full path to the configuration file.
    #>
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
        Write-Warning "Failed to load Nebula.Core configuration from '$Path'. $($_.Exception.Message)"
        return @{}
    }
}

function Merge-NCConfig {
    <#
    .SYNOPSIS
        Applies override key/value pairs onto a base hashtable.
    .DESCRIPTION
        Iterates over the override hashtable and copies each entry into the base,
        returning the mutated base reference for chaining.
    .PARAMETER Base
        Hashtable to mutate.
    .PARAMETER Override
        Hashtable containing replacement values.
    .EXAMPLE
        Create a new hashtable with overrides and save it in a PSD1 file (e.g. %USERPROFILE%\.NebulaCore\settings.psd1 or C:\ProgramData\Nebula.Core\settings.psd1):
        @{
            CSV_Encoding        = 'UTF8'
            DateTimeString_Full = 'yyyy-MM-dd HH:mm:ss'
            MaxFieldLength      = 50
        }

    #>
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

# --- Load Private helpers first (NOT exported) ---
$privateDir = Join-Path $PSScriptRoot 'Private'
if (Test-Path $privateDir) {
    Get-ChildItem -Path $privateDir -Filter '*.ps1' -File -Recurse | Sort-Object FullName | ForEach-Object {
        try {
            . $_.FullName  # dot-source
        } catch {
            throw "Failed to load Private script '$($_.Name)': $($_.Exception.Message)"
        }
    }
}

# --- Load Public entry points (will be exported) ---
$publicDir = Join-Path $PSScriptRoot 'Public'
if (Test-Path $publicDir) {
    Get-ChildItem -Path $publicDir -Filter '*.ps1' -File | ForEach-Object {
        try {
            . $_.FullName  # dot-source
        } catch {
            throw "Failed to load Public script '$($_.Name)': $($_.Exception.Message)"
        }
    }
}

# --- Custom variables to be used in the module (user configuration can override them) ---
$NC_Defaults = [ordered]@{
    CSV_DefaultLimiter      = ";"
    CSV_Encoding            = 'ISO-8859-15'
    DateTimeString_CSV      = 'yyyyMMdd'
    DateTimeString_Full     = 'dd/MM/yyyy HH:mm:ss'
    LicenseCacheDays        = 7
    LicenseCacheDirectory   = $ncCacheRoot
    MaxFieldLength          = 35
    UserConfigRoot          = $ncUserRoot
}

$NC_Config = [ordered]@{}
foreach ($key in $NC_Defaults.Keys) {
    $NC_Config[$key] = $NC_Defaults[$key]
}

$userConfigPath = Join-Path -Path $NC_Defaults.UserConfigRoot -ChildPath 'settings.psd1'
$machineConfigPath = Join-Path -Path $env:ProgramData -ChildPath 'Nebula.Core\settings.psd1'

$userConfig = Import-NCConfigurationFile -Path $userConfigPath
$machineConfig = Import-NCConfigurationFile -Path $machineConfigPath

if ($machineConfig.Count) {
    $NC_Config = Merge-NCConfig -Base $NC_Config -Override $machineConfig
}

if ($userConfig.Count) {
    $NC_Config = Merge-NCConfig -Base $NC_Config -Override $userConfig
}

foreach ($key in $NC_Config.Keys) {
    $envVarName = "NEBULA_{0}" -f ($key.ToUpperInvariant())
    $envValue = [Environment]::GetEnvironmentVariable($envVarName)
    if ($envValue) {
        $NC_Config[$key] = $envValue
    }
}

New-Variable -Name NCVars -Value $NC_Config -Scope Script -Force
