Import-Module .\bin\Debug\net47\Cellmate.Cmdlets.dll
$VerbosePreference = "continue"

$books = @(
    "spec1.xlsx",
    "spec2.xlsx"
)

$pdf = "merged.pdf"

Get-Item $books |
    Import-Excel -Visible |
    Merge-Excel -PageNumber -Destination $pdf |
    Out-Null
