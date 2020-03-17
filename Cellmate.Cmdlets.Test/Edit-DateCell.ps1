Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll
$VerbosePreference = "continue"

Get-Item spec*.xlsx |
    Import-Excel -Visible |
    Edit-DateCell -Before 2020/6/1 -Value 2020/12/25 | 
    Merge-Excel -Destination "replaced.pdf" |
    Out-Null
