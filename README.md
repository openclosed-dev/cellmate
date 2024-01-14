# Cellmate

Cellmate is a [PowerShell module] for processing Microsoft Excel workbooks.

[PowerShell module]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_modules?view=powershell-5.1

## Prerequisite
* Windows PowerShell 5.1
* .NET Framework 4.8
* Microsoft Excel

Note that PowerShell 6 and .NET Core are not supported by this PowerShell module.

## How to Install

The latest stable version of zip file for distribution can be downloaded from [Release](https://github.com/openclosed-dev/cellmate/releases) page of this repository.

1. Close active PowerShell sessions if any exists.
2. Unpack the downloaded zip file `Cellmate-<version>.zip`.
3. Copy the unpacked `Cellmate` directory into `\Users\<user name>\Documents\WindowsPowerShell\Modules` for your account. The resulting directory will be `\Users\<user name>\Documents\WindowsPowerShell\Modules\Cellmate\<version>`.

## Code Samples

This section shows PowerShell scripts as examples.

### Merging workbooks into a PDF file
_merge-books.ps1_
```powershell
Import-Module Cellmate

$VerbosePreference = "continue"
$books = "book1.xlsx", "book2.xlsx", "book3.xlsx"

Get-Item $books |
    Import-Workbook |
    Merge-Workbook -As Pdf -PageNumber Right "target.pdf" |
    Out-Null
```

### Archiving workbooks into a ZIP file
_archive-books.ps1_
```powershell
Import-Module Cellmate

$VerbosePreference = "continue"
$books = "book1.xlsx", "book2.xlsx", "book3.xlsx"

Get-Item $books |
    Import-Workbook |
    Compress-Workbook "target.zip" |
    Out-Null
```

## List of Offered Cmdlets

| Name | Description |
| --- | --- |
| Compress-Workbook | Creates a ZIP archive containing one or more workbooks. |
| Export-Workbook | Saves a workbook into a file.  |
| Import-Workbook | Loads a workbook from the specified path. |
| Merge-Workbook | Merges one or more workbooks into a PDF file. |

## Legal Notice
Copyright 2020-2024 the original author or authors. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this product except in compliance with the License.
You may obtain a copy of the License at
<http://www.apache.org/licenses/LICENSE-2.0>
