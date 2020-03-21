$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
Describe "Test-DateCell" {

    It "outputs all dates" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell > $null 3> "$result\output-1.txt"

        Get-Content "$result\output-1.txt" |
            Should -Be @(
                'Test-DateCell-1.csv:Test-DateCell-1:A1 2020/03/03 0:00:00',
                'Test-DateCell-1.csv:Test-DateCell-1:A2 2020/04/04 0:00:00',
                'Test-DateCell-1.csv:Test-DateCell-1:A3 2030/05/05 0:00:00'
            )
    }

    It "outputs dates before the specified" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Before 2020/5/1 > $null 3> "$result\output-2.txt"

        Get-Content "$result\output-2.txt" |
            Should -Be @(
                'Test-DateCell-1.csv:Test-DateCell-1:A1 2020/03/03 0:00:00',
                'Test-DateCell-1.csv:Test-DateCell-1:A2 2020/04/04 0:00:00'
            )
    }

    It "outputs dates after the specified" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -After 2020/4/1 > $null 3> "$result\output-3.txt"

        Get-Content "$result\output-3.txt" |
            Should -Be @(
                'Test-DateCell-1.csv:Test-DateCell-1:A2 2020/04/04 0:00:00',
                'Test-DateCell-1.csv:Test-DateCell-1:A3 2030/05/05 0:00:00'
            )
    }

    It "outputs dates in the period" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -After 2020/4/1 -Before 2020/5/1 > $null 3> "$result\output-4.txt"

        Get-Content "$result\output-4.txt" |
            Should -Be @(
                'Test-DateCell-1.csv:Test-DateCell-1:A2 2020/04/04 0:00:00'
            )
    }

    It "outputs dates in the range" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Range 1:2 > $null 3> "$result\output-5.txt"

        Get-Content "$result\output-5.txt" |
            Should -Be @(
                'Test-DateCell-1.csv:Test-DateCell-1:A1 2020/03/03 0:00:00',
                'Test-DateCell-1.csv:Test-DateCell-1:A2 2020/04/04 0:00:00'
            )
    }

    It "outputs dates in the range and period" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Range 1:2 -After 2020/4/1 > $null 3> "$result\output-6.txt"

        Get-Content "$result\output-6.txt" |
            Should -Be @(
                'Test-DateCell-1.csv:Test-DateCell-1:A2 2020/04/04 0:00:00'
            )
    }

    It "throws an exception if the range is invalid" {
        {
            Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Range 1
        } | Should -Throw
    }
}
