param([string]$TargetModule = 'Cellmate')
Import-Module $TargetModule
Invoke-Pester
