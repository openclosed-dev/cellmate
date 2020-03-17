Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll
$VerbosePreference = "continue"

Get-Item spec*.xlsx |
    Import-Excel |
    Test-DateCell |
    Out-Null
