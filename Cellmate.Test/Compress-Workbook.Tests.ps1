$VerbosePreference = 'Continue'

BeforeAll {
    $here = Split-Path -Parent $PSCommandPath
    $result = "$here\result"
    $books = @("Months.csv", "SolarSystem.csv")
}

Describe "Compress-Workbook" {

    It "compresses workbooks as ZIP" {
        Get-Item $books |
            Import-Workbook |
            Compress-Workbook -Destination "$result\output-1.zip" > $null

        "$result\output-1.zip" | Should -Exist
    }

    It "compresses workbooks as ZIP using positional parameter" {
        Get-Item $books |
            Import-Workbook |
            Compress-Workbook "$result\output-2.zip" > $null

        "$result\output-2.zip" | Should -Exist
    }
}
