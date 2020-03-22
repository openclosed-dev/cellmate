$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
$VerbosePreference = 'Continue'

Describe "Edit-DateCell" {

    It "replaces all dates" {
        Get-Item "Dates.csv" | 
            Import-Workbook |
            Edit-DateCell -Value 2020/12/25 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates.csv" |
            Should -Be @('2020/12/25', '2020/12/25', '2020/12/25')
    }

    It "replaces dates before the specified" {
        Get-Item "Dates.csv" | 
            Import-Workbook |
            Edit-DateCell -Before 2020/5/1 -Value 2020/12/25 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates.csv" |
            Should -Be @('2020/12/25', '2020/12/25', '2020/5/5')
    }

    It "replaces dates after the specified" {
        Get-Item "Dates.csv" | 
            Import-Workbook |
            Edit-DateCell -After 2020/4/1 -Value 2020/12/25 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates.csv" |
            Should -Be @('2020/3/3', '2020/12/25', '2020/12/25') 
    }

    It "replaces dates in the period" {
        Get-Item "Dates.csv" | 
            Import-Workbook |
            Edit-DateCell -After 2020/4/1 -Before 2020/5/1 -Value 2020/12/25 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates.csv" |
            Should -Be @('2020/3/3', '2020/12/25', '2020/5/5') 
    }

    It "replaces dates in the range" {
        Get-Item "Dates.csv" | 
            Import-Workbook |
            Edit-DateCell -Range 1:2 -Value 2020/12/25 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates.csv" |
            Should -Be @('2020/12/25', '2020/12/25', '2020/5/5') 
    }

    It "replaces dates in the range and period" {
        Get-Item "Dates.csv" | 
            Import-Workbook |
            Edit-DateCell -Range 1:2 -After 2020/4/1 -Value 2020/12/25 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates.csv" |
            Should -Be @('2020/3/3', '2020/12/25', '2020/5/5') 
    }

    It "throws an exception if the value is not a date" {
        {
            Get-Item "Dates.csv" | 
                Import-Workbook |
                Edit-DateCell -Value hello
        } | Should -Throw
    }
}
