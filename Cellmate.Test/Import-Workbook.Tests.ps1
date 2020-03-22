Describe "Import-Workbook" {

    It "imports a single workbook" {
        $actual = Get-Item "Months.csv" | 
            Import-Workbook |
            ForEach-Object { $_.Worksheets.Count }
            
        $actual | Should -Be 1 
    }

    It "imports multiple workbooks" {
        $books = "Months.csv", "SolarSystem.csv"
        $actual = Get-Item $books | 
            Import-Workbook |
            ForEach-Object { $_.Worksheets.Count }

            $actual | Should -Be 1 ,1
    }
}
