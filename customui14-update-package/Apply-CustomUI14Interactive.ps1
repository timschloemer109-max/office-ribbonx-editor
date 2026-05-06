[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$setScript = Join-Path $scriptRoot 'Set-CustomUI14Xml.ps1'

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $setScript -Interactive
exit $LASTEXITCODE
