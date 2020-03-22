# Cellmate

Cellmate is a collection of PowerShell cmdlets for processing Excel workbooks.

## How to Intall

Unpack the zip file `Cellmate-<version>.zip` into the folder `\Users\<your name>\Documents\WindowsPowerShell\Modules` for your account.

## Examples

This section shows PowerShell scripts as examples.

#### Merging multiple workbooks into a PDF file
```powershell
Import-Module Cellmate
$VerbosePreference = 'continue'
$books = 'book1.xlsx', 'book2.xlsx'
Get-Item $books |
    Import-Workbook |
    Merge-Workbook -As Pdf -PageNumber Right -Destination 'merged.pdf' |
    Out-Null
```

## Copyright Notice
Copyright 2020 the original author or authors. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this product except in compliance with the License.
You may obtain a copy of the License at
http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
