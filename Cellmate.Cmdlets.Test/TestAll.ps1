param([string]$Configuration = 'Debug')

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

Import-Module $root\Cellmate.Cmdlets\bin\$Configuration\net47\Cellmate.Cmdlets.dll
Invoke-Pester
