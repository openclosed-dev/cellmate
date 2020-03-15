Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll
$books = @(
    "sample1.xlsx",
    "sample2.xlsx"
)

Get-Item $books |
    Import-Excel -Visible |
    Merge-Excel -Verbose -PageNumber -Path merged.pdf
