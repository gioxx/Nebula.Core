#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Security helpers ===========================================================================================================

function Get-NCContentFilterPolicyValues {
    <#
    .SYNOPSIS
        Normalizes content filter policy list values.
    .DESCRIPTION
        Extracts sender or domain values from hosted content filter policy properties and returns
        a sorted, unique string array. Used internally by the Security cmdlets.
    .PARAMETER PolicyObject
        Hosted content filter policy object.
    .PARAMETER PropertyName
        Property that contains the list to normalize.
    .PARAMETER PreferredValueProperty
        Preferred field name to read from each item (for example Sender or Domain).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$PolicyObject,
        [Parameter(Mandatory)]
        [string]$PropertyName,
        [string]$PreferredValueProperty
    )

    $result = [System.Collections.Generic.List[string]]::new()

    foreach ($item in @($PolicyObject.$PropertyName)) {
        if ($null -eq $item) {
            continue
        }

        $value = $null
        if ($PreferredValueProperty -and $item.PSObject.Properties.Match($PreferredValueProperty).Count -gt 0) {
            $value = [string]$item.$PreferredValueProperty
        }
        elseif ($item.PSObject.Properties.Match('Sender').Count -gt 0) {
            $value = [string]$item.Sender
        }
        elseif ($item.PSObject.Properties.Match('Domain').Count -gt 0) {
            $value = [string]$item.Domain
        }
        else {
            $value = [string]$item
        }

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $result.Add($value.Trim()) | Out-Null
        }
    }

    @($result | Sort-Object -Unique)
}
