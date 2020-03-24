$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
$VerbosePreference = 'Continue'
Describe "Merge-Cell" {

    $books = 'MergeCell-1.xlsx'

    It "merges cells" -SKip {
        Get-Item $books | 
            Import-Workbook |
            Merge-Cell -Range 1:2 -Bordered -SkipColumns 1 |
            Export-Workbook $result
    }
}
