param([string]$Configuration = 'Debug')

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

Import-Module $root\Cellmate\bin\$Configuration\net47\Cellmate.dll

New-Item "$here\result" -ItemType Directory -Force > $null
Remove-Item "$here\result\*"

Invoke-Pester
