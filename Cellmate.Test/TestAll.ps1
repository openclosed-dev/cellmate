param([string]$Configuration = 'Debug')

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

New-Item "$here\result" -ItemType Directory -Force > $null
Remove-Item "$here\result\*"

Import-Module $root\Cellmate\bin\$Configuration\net47\Cellmate.dll
Invoke-Pester
Remove-Module Cellmate
