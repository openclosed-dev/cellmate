$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
$VerbosePreference = 'Continue'
Describe "Merge-Workbook" {

    $books = @('Months.csv', 'SolarSystem.csv')

    It "merges workbooks as PDF" {
        Get-Item $books | 
            Import-Workbook |
            Merge-Workbook -As Pdf -Destination "$result\merged-1.pdf" > $null

        "$result\merged-1.pdf" | Should -Exist
    }

    It "merges workbooks as PDF using positional parameter" {
        Get-Item $books | 
            Import-Workbook |
            Merge-Workbook -As Pdf "$result\merged-2.pdf" > $null

        "$result\merged-2.pdf" | Should -Exist
    }

    It "merges workbooks as PDF with page number right" {
        Get-Item $books | 
            Import-Workbook |
            Merge-Workbook -As Pdf -PageNumber Right -Destination "$result\merged-3.pdf" > $null

        "$result\merged-3.pdf" | Should -Exist
    }

    It "merges workbooks as PDF with page number center" {
        Get-Item $books | 
            Import-Workbook |
            Merge-Workbook -As Pdf -PageNumber Center -Destination "$result\merged-4.pdf" > $null

        "$result\merged-4.pdf" | Should -Exist
    }
}
