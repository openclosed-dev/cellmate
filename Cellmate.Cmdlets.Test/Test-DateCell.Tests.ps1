$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Describe "Test-DateCell" {

    It "outputs all dates" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell > $null 3> 'TestDrive:\output-1.txt'

        Get-Content 'TestDrive:\output-1.txt' |
            Should -Be (Get-Content "$here\Test-DateCell-1.expected.txt") 
    }

    It "outputs dates before the specified" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Before 2020/4/1 > $null 3> 'TestDrive:\output-2.txt'

        Get-Content 'TestDrive:\output-2.txt' |
            Should -Be (Get-Content "$here\Test-DateCell-2.expected.txt") 
    }

    It "outputs dates after the specified" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -After 2020/3/1 > $null 3> 'TestDrive:\output-3.txt'

        Get-Content 'TestDrive:\output-3.txt' |
            Should -Be (Get-Content "$here\Test-DateCell-3.expected.txt") 
    }

    It "outputs dates in the period" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -After 2020/3/1 -Before 2020/4/1 > $null 3> 'TestDrive:\output-4.txt'

        Get-Content 'TestDrive:\output-4.txt' |
            Should -Be (Get-Content "$here\Test-DateCell-4.expected.txt") 
    }

    It "outputs dates in the range" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Range 1:2 > $null 3> 'TestDrive:\output-5.txt'

        Get-Content 'TestDrive:\output-5.txt' |
            Should -Be (Get-Content "$here\Test-DateCell-5.expected.txt") 
    }

    It "outputs dates in the range and period" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Range 1:2 -After 2020/3/1 > $null 3> 'TestDrive:\output-6.txt'

        Get-Content 'TestDrive:\output-6.txt' |
            Should -Be (Get-Content "$here\Test-DateCell-6.expected.txt") 
    }

    It "throws an exception if the range is invalid" {
        {
            Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Range 1
        } | Should -Throw
    }
}
