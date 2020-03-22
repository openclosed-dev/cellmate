# Cellmate

Cellmate is a collection of PowerShell cmdlets for processing Excel workbooks.

## Prerequisite
* PowerShell 5.1 or higher
* .NET Framework 4.7 or higher
* Microsoft Excel

## How to Install

1. Unpack the zip file `Cellmate-<version>.zip`.
2. Move the unpacked `Cellmate` directory into `\Users\<user name>\Documents\WindowsPowerShell\Modules` for your account.

## Code Samples

This section shows PowerShell scripts as examples.

#### Merging workbooks into a PDF file
_merge-books.ps1_
```powershell
Import-Module Cellmate

$VerbosePreference = 'continue'
$books = 'book1.xlsx', 'book2.xlsx', 'book3.xlsx'

Get-Item $books |
    Import-Workbook |
    Merge-Workbook -As Pdf -PageNumber Right -Destination 'target.pdf' |
    Out-Null
```

#### Compressing workbooks into a ZIP file
_compress-books.ps1_
```powershell
Import-Module Cellmate

$VerbosePreference = 'continue'
$books = 'book1.xlsx', 'book2.xlsx', 'book3.xlsx'

Get-Item $books |
    Import-Workbook |
    Compress-Workbook -Destination 'target.pdf' |
    Out-Null
```

## Copyright Notice
Copyright 2020 the original author or authors. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this product except in compliance with the License.
You may obtain a copy of the License at
<http://www.apache.org/licenses/LICENSE-2.0>
