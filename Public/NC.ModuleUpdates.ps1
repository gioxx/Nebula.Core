#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: Module update checks (public) ========================================================================================================

function Get-NebulaModuleUpdates {
    <#
    .SYNOPSIS
        Checks PowerShell Gallery for Nebula module updates on demand.
    .DESCRIPTION
        Forces a fresh check against PSGallery and reports any available updates.
        Returns $true when updates are found, otherwise $false.
    #>
    [CmdletBinding()]
    param()

    return (Test-NebulaModuleUpdates -Force)
}
