Describe "Test-DateCell" {

    It "outputs all dates" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell > $null 3> 'TestDrive:\output-1.txt'

        Get-Content 'TestDrive:\output-1.txt' |
            Should -Be (Get-Content '.\Test-DateCell-1.expected.txt') 
    }

    It "outputs dates before the specified" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -Before 2020/4/1 > $null 3> 'TestDrive:\output-2.txt'

        Get-Content 'TestDrive:\output-2.txt' |
            Should -Be (Get-Content '.\Test-DateCell-2.expected.txt') 
    }

    It "outputs dates after the specified" {
        Get-Item "Test-DateCell-1.csv" | 
            Import-Workbook |
            Test-DateCell -After 2020/3/1 > $null 3> 'TestDrive:\output-3.txt'

        Get-Content 'TestDrive:\output-3.txt' |
            Should -Be (Get-Content '.\Test-DateCell-3.expected.txt') 
    }
}
