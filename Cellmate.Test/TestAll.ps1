param([string]$configuration = 'Product')

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

New-Item "$here\result" -ItemType Directory -Force > $null
Get-ChildItem $here -Filter "result" | Get-ChildItem | Remove-Item

if ($configuration -eq 'Product') {
    $sut = 'Cellmate'
} else {
    $sut = "$root\Cellmate\bin\$configuration\net47\Cellmate.dll"
}

Write-Host "Loading a module: $sut"
Import-Module $sut
Invoke-Pester
Remove-Module Cellmate
