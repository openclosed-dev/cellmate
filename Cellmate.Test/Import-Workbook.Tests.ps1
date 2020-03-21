Describe "Import-Workbook" {

    It "imports a single workbook" {
        $actual = Get-Item "Import-Workbook-1.csv" | 
            Import-Workbook |
            ForEach-Object { $_.Worksheets.Count }
            
        $actual | Should -Be 1 
    }

    It "imports multiple workbooks" {
        $books = "Import-Workbook-1.csv", "Import-Workbook-2.csv"
        $actual = Get-Item $books | 
            Import-Workbook |
            ForEach-Object { $_.Worksheets.Count }

            $actual | Should -Be 1 ,1
    }
}
