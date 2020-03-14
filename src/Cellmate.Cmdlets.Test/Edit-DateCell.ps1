Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll

Get-Item *.xlsx |
    Import-Excel |
    Edit-DateCell -Value 2020/3/1 | 
    Export-Excel -Format pdf |
    Out-Null
