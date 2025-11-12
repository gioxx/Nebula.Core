#Requires -Version 5.0
using namespace System.Management.Automation

# Nebula.Core: (Private) Custom Write Information ===================================================================================================

# Script-scoped color map so it can be overridden at runtime
# Foreground and Background are [ConsoleColor] values.
# Background can be $null to preserve the current background color.
Set-Variable -Name InfoColorMap -Scope Script -Force -Value ([ordered]@{
    INFO    = @{ Foreground = [ConsoleColor]::Gray; Background = $null }
    SUCCESS = @{ Foreground = [ConsoleColor]::Green; Background = $null }
    WARNING = @{ Foreground = [ConsoleColor]::Yellow; Background = $null }
    ERROR   = @{ Foreground = [ConsoleColor]::Red; Background = $null }
    DEBUG   = @{ Foreground = [ConsoleColor]::Cyan; Background = $null }
    VERBOSE = @{ Foreground = [ConsoleColor]::DarkGray; Background = $null }
})

Function Set-InfoColorMap {
    <#
    .SYNOPSIS
        Overrides one or more level color mappings.
    .DESCRIPTION
        Provide a hashtable where keys are levels (INFO, SUCCESS, WARNING, ERROR, DEBUG, VERBOSE)
        and values are hashtables with Foreground/Background [ConsoleColor] (Background can be $null).
    .EXAMPLE
        Set-InfoColorMap @{ WARNING = @{ Foreground = 'DarkYellow'; Background = $null } }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Map
    )

    foreach ($k in $Map.Keys) {
        if (-not $script:InfoColorMap.Contains($k)) {
            throw "Unknown level '$k'. Valid levels: $($script:InfoColorMap.Keys -join ', ')."
        }
        $entry = $Map[$k]
        if (-not ($entry -is [hashtable] -and $entry.ContainsKey('Foreground'))) {
            throw "Map for '$k' must contain at least a 'Foreground' key."
        }
        $fg = [ConsoleColor]::Parse([ConsoleColor], "$($entry.Foreground)")
        $bg = if ($entry.ContainsKey('Background') -and $null -ne $entry.Background) {
            [ConsoleColor]::Parse([ConsoleColor], "$($entry.Background)")
        }
        else { $null }
        $script:InfoColorMap[$k] = @{ Foreground = $fg; Background = $bg }
    }
}

Function Write-NCMessage {
    <#
    .SYNOPSIS
        Writes messages to the Information stream with color and level tagging.
    .DESCRIPTION
        Honors $InformationPreference. Adds level-based color mapping and tags.
        If host coloring is not available, falls back to plain Write-Information.
    .PARAMETER Message
        Object to write (string or any object). Strings are prefixed with [LEVEL].
    .PARAMETER Level
        One of: INFO, SUCCESS, WARNING, ERROR, DEBUG, VERBOSE.
    .PARAMETER NoNewline
        Do not append a newline when writing to host.
    .PARAMETER ForegroundColor
        Override level's foreground color.
    .PARAMETER BackgroundColor
        Override level's background color (use $null to keep current).
    .EXAMPLE
        Write-NCMessage "Processing complete" -Level SUCCESS
    .EXAMPLE
        Write-NCMessage "Low disk space" -Level WARNING -NoNewline
    .EXAMPLE
        Write-NCMessage "Custom color" -Level INFO -ForegroundColor Magenta
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [Alias('MessageData')]
        [object]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR', 'DEBUG', 'VERBOSE')]
        [string]$Level = 'INFO',
        [switch]$NoNewline,
        [Nullable[ConsoleColor]]$ForegroundColor,
        [Nullable[ConsoleColor]]$BackgroundColor
    )

    # Resolve colors: explicit overrides > map > host defaults
    $map = $script:InfoColorMap[$Level]
    $fg = if ($ForegroundColor.HasValue) { $ForegroundColor.Value }
    elseif ($map -and $map.Foreground) { $map.Foreground }
    else { $Host.UI.RawUI.ForegroundColor }

    $bg = if ($BackgroundColor.IsPresent) { $BackgroundColor.Value }
    elseif ($map -and $map.Background) { $map.Background }
    else { $Host.UI.RawUI.BackgroundColor }

    # Prepare message text: add [LEVEL] prefix for strings (keeps objects intact)
    $msgText =
    # if ($Message -is [string]) { "[{0}] {1}" -f $Level, $Message }
    if ($Message -is [string]) { "{0}" -f $Message }
    else { $Message }

    # Build HostInformationMessage. If the host doesn't support colors, these fields are ignored.
    $hostMsg = [HostInformationMessage]@{
        Message         = $msgText
        ForegroundColor = $fg
        BackgroundColor = $bg
        NoNewline       = $NoNewline.IsPresent
    }

    # Pass level as a tag for filtering (Receive-Job, transcript parsers, etc.)
    Write-Information -MessageData $hostMsg -Tags $Level -InformationAction Continue
}