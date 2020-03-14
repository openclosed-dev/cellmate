Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll

Get-Item *.xlsx |
    Import-Excel -Visible |
    Edit-DateCell -Verbose -Range 1:2 -Value 2020/3/1 | 
    ConvertFrom-Excel -Format pdf |
    Out-Null
