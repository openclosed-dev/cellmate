$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
$VerbosePreference = 'Continue'
Describe "Compress-Workbook" {

    $books = @("Months.csv", "SolarSystem.csv")

    It "compresses workbooks as ZIP" {
        Get-Item $books | 
            Import-Workbook |
            Compress-Workbook -Destination "$result\output-1.zip" > $null

        "$result\output-1.zip" | Should -Exist
    }
}
