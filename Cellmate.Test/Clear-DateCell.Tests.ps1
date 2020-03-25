$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
$VerbosePreference = 'Continue'

Describe "Clear-DateCell" {

    It "clears all dates" {
        Get-Item "Dates-1.csv" | 
            Import-Workbook |
            Clear-DateCell |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates-1.csv" |
            Should -Be @('', '', '')
    }

    It "clears dates before the specified" {
        Get-Item "Dates-1.csv" | 
            Import-Workbook |
            Clear-DateCell -Before 2020/5/1 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates-1.csv" |
            Should -Be @('', '', '2020/5/5')
    }

    It "clears dates after the specified" {
        Get-Item "Dates-1.csv" | 
            Import-Workbook |
            Clear-DateCell -After 2020/4/1 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates-1.csv" |
            Should -Be @('2020/3/3', '', '') 
    }

    It "clears dates in the period" {
        Get-Item "Dates-1.csv" | 
            Import-Workbook |
            Clear-DateCell -After 2020/4/1 -Before 2020/5/1 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates-1.csv" |
            Should -Be @('2020/3/3', '', '2020/5/5') 
    }

    It "clears dates in the range" {
        Get-Item "Dates-1.csv" | 
            Import-Workbook |
            Clear-DateCell -Range 1:2 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates-1.csv" |
            Should -Be @('', '', '2020/5/5') 
    }

    It "clears dates in the range and period" {
        Get-Item "Dates-1.csv" | 
            Import-Workbook |
            Clear-DateCell -Range 1:2 -After 2020/4/1 |
            Export-Workbook -As Csv -Destination $result > $null

        Get-Content "$result\Dates-1.csv" |
            Should -Be @('2020/3/3', '', '2020/5/5') 
    }
}
