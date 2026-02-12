#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Module update checks =================================================================================================================

function Test-NebulaModuleUpdates {
    <#
    .SYNOPSIS
        Checks PowerShell Gallery for Nebula module updates.
    .DESCRIPTION
        Compares locally installed Nebula.* modules with the latest versions on PSGallery.
        Returns $true when updates are found, otherwise $false. Logs results via Write-NCMessage.
    .PARAMETER Force
        Forces a new check even if one was already performed in this session.
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    Write-NCMessage "Checking PowerShell Gallery for module updates, please wait ..." -Level INFO

    if (-not $Force.IsPresent -and $script:NC_ModuleUpdateChecked) {
        return $false
    }

    $intervalHours = $NCVars.CheckUpdatesIntervalHours
    if ($intervalHours -is [string]) {
        $parsed = 0
        if ([int]::TryParse($intervalHours, [ref]$parsed)) {
            $intervalHours = $parsed
        }
    }

    if (-not $Force.IsPresent -and $intervalHours -is [int] -and $intervalHours -gt 0) {
        $lastCheckUtc = Get-NCModuleUpdateLastCheck
        if ($lastCheckUtc) {
            $elapsedHours = ((Get-Date).ToUniversalTime() - $lastCheckUtc).TotalHours
            if ($elapsedHours -lt $intervalHours) {
                $script:NC_ModuleUpdateChecked = $true
                return $false
            }
        }
    }

    $script:NC_ModuleUpdateChecked = $true

    if (-not (Get-Command -Name Get-InstalledModule -ErrorAction SilentlyContinue)) {
        Write-NCMessage "PowerShellGet is required to check for Nebula updates. Install it with Install-Module PowerShellGet." -Level WARNING
        return $false
    }

    $installedModules = @()
    try {
        $installedModules = Get-InstalledModule -Name 'Nebula.*' -ErrorAction Stop
    }
    catch {
        Write-NCMessage "Unable to read installed Nebula modules. $($_.Exception.Message)" -Level WARNING
        return $false
    }

    $extraModules = @()
    try {
        $extraModules = Get-InstalledModule -Name @('ExchangeOnlineManagement', 'Microsoft.Graph') -ErrorAction SilentlyContinue
    }
    catch {
        $extraModules = @()
    }

    if ($extraModules) {
        $installedModules = @($installedModules + $extraModules | Where-Object { $_ }) |
            Group-Object Name | ForEach-Object { $_.Group | Select-Object -First 1 }
    }

    if (-not $installedModules -or $installedModules.Count -eq 0) {
        return $false
    }


    $updates = @()
    foreach ($module in $installedModules) {
        $galleryModule = $null
        try {
            $galleryModule = Find-Module -Name $module.Name -Repository PSGallery -ErrorAction Stop
        }
        catch {
            Write-NCMessage "Unable to reach PowerShell Gallery to check updates. $($_.Exception.Message)" -Level WARNING
            return $false
        }

        if (-not $galleryModule) {
            continue
        }

        $installedVersion = [version]$module.Version
        $latestVersion = [version]$galleryModule.Version

        if ($latestVersion -gt $installedVersion) {
            $updates += [pscustomobject]@{
                Name             = $module.Name
                InstalledVersion = $installedVersion
                LatestVersion    = $latestVersion
            }
        }
    }

    if (-not $updates -or $updates.Count -eq 0) {
        Save-NCModuleUpdateLastCheck
        return $false
    }

    foreach ($update in $updates) {
        Write-NCMessage ("Update available for {0}: {1} -> {2}" -f $update.Name, $update.InstalledVersion, $update.LatestVersion) -Level WARNING
    }

    Save-NCModuleUpdateLastCheck
    return $true
}

function Get-NCModuleUpdateLastCheck {
    [CmdletBinding()]
    param()

    $path = Get-NCModuleUpdateCheckPath
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        $data = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($data -and $data.LastCheckUtc) {
            return [datetime]::Parse($data.LastCheckUtc, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
        }
    }
    catch {
        return $null
    }

    return $null
}

function Save-NCModuleUpdateLastCheck {
    [CmdletBinding()]
    param()

    $path = Get-NCModuleUpdateCheckPath
    $directory = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    $payload = @{
        LastCheckUtc = (Get-Date).ToUniversalTime().ToString('o')
    }

    $payload | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8
}

function Get-NCModuleUpdateCheckPath {
    [CmdletBinding()]
    param()

    return (Join-Path -Path $NCVars.UserConfigRoot -ChildPath 'update-check.json')
}
