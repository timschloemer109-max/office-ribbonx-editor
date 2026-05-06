[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [string] $OutputPath,

    [switch] $ForceInPlace
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$updateScript = Join-Path $scriptRoot 'Update-CustomUI14.ps1'
$definition = Join-Path $scriptRoot 'ribbon-def.txt'

$arguments = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', $updateScript,
    '-WorkbookPath', $WorkbookPath,
    '-DefinitionPath', $definition
)

if ($ForceInPlace) {
    $arguments += '-ForceInPlace'
} elseif (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $arguments += @('-OutputPath', $OutputPath)
}

& powershell.exe @arguments
exit $LASTEXITCODE
