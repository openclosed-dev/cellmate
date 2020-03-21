$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$result = "$here\result"
$VerbosePreference = 'Continue'
Describe "Export-Workbook" {

    It "exports as default" {
        Get-Item "Export-Workbook-1.csv" | 
            Import-Workbook |
            Export-Workbook -Destination $result > $null
            
        "$result\Export-Workbook-1.xlsx" | Should -Exist 
    }

    It "exports as explicit default" {
        Get-Item "Export-Workbook-2.csv" | 
            Import-Workbook |
            Export-Workbook -As Default -Destination $result > $null
            
        "$result\Export-Workbook-2.xlsx" | Should -Exist 
    }

    It "exports as csv" {
        Get-Item "Export-Workbook-3.csv" | 
            Import-Workbook |
            Export-Workbook -As Csv -Destination $result > $null
            
        "$result\Export-Workbook-3.csv" | Should -Exist
        Get-Content "$result\Export-Workbook-3.csv" |
            Should -Be (Get-Content "$here\Export-Workbook-3.csv") 
    }

    It "throws if the format is invalid" {
        {
            Get-Item "Export-Workbook-1.csv" | 
                Import-Workbook |
                Export-Workbook -As unknown
        } | Should -Throw
    }
}
