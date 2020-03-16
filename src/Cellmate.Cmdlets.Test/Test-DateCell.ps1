Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll

$VerbosePreference = 'Continue'

Get-Item sample*.xlsx |
    Import-Excel |
    Test-DateCell -Range 3:3 |
    Out-Null
