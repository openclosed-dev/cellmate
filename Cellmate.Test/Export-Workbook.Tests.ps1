$VerbosePreference = 'Continue'

BeforeAll {
    $here = Split-Path -Parent $PSCommandPath
    $result = "$here\result"
}

Describe "Export-Workbook" {

    It "exports as default" {
        Get-Item "Export-Workbook-1.csv" |
            Import-Workbook |
            Export-Workbook -Destination $result > $null

        "$result\Export-Workbook-1.csv" | Should -Exist
    }

    It "exports as explicit default" {
        Get-Item "Export-Workbook-2.csv" |
            Import-Workbook |
            Export-Workbook -As Default -Destination $result > $null

        "$result\Export-Workbook-2.xlsx" | Should -Exist
    }

    It "exports as CSV" {
        Get-Item "Export-Workbook-1.csv" |
            Import-Workbook |
            Export-Workbook -As Csv -Destination $result > $null

        "$result\Export-Workbook-1.csv" | Should -Exist
        Get-Content "$result\Export-Workbook-1.csv" |
            Should -Be (Get-Content "$here\Export-Workbook-1.csv")
    }

    It "exports as PDF" {
        Get-Item "Export-Workbook-1.csv" |
            Import-Workbook |
            Export-Workbook -As Pdf -Destination $result > $null

        "$result\Export-Workbook-1.pdf" | Should -Exist
    }

    It "throws if the format is invalid" {
        {
            Get-Item "Export-Workbook-1.csv" |
                Import-Workbook |
                Export-Workbook -As unknown
        } | Should -Throw
    }
}
