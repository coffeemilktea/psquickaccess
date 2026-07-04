<#
.SYNOPSIS
    Pins one or more folders to Quick Access in Windows File Explorer.
.DESCRIPTION
    Accepts folder paths as command-line arguments, or prompts for them
    interactively when none are given. Each existing folder is pinned to
    Quick Access via the Shell.Application COM object.
.EXAMPLE
    .\pinqa.ps1 "C:\Projects" "D:\Media"
.EXAMPLE
    .\pinqa.ps1
    Prompts for paths one per line; type 'done' or leave blank to finish.
#>
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths = @()
)

# Prompt interactively when no paths were passed as arguments.
# Note: $entry, not $input - $input is a reserved automatic variable.
if ($Paths.Count -eq 0) {
    Write-Output "Enter the folder paths to pin to Quick Access ('done' or a blank line to end the list):"
    while ($entry = Read-Host) {
        if ($entry -eq 'done') { break }
        $Paths += $entry
    }
}

if ($Paths.Count -eq 0) {
    Write-Output 'No paths entered; nothing to pin.'
    return
}

$shell = New-Object -ComObject Shell.Application

foreach ($path in $Paths) {
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Warning "The path '$path' does not exist."
        continue
    }

    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        Write-Warning "'$path' is a file. Quick Access only pins folders - skipping."
        continue
    }

    $folder = (Resolve-Path -LiteralPath $path).ProviderPath

    # 'pintohome' is the canonical name of the "Pin to Quick access" verb,
    # so it works regardless of the Windows display language.
    $shell.Namespace($folder).Self.InvokeVerb('pintohome')
    Write-Output "Pinned '$folder' to Quick Access."
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
