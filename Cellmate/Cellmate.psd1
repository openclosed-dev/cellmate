#
# Module manifest for module 'Cellmate'
#

@{

# Script module or binary module file associated with this manifest.
RootModule = '.\Cellmate.dll'

# Version number of this module.
ModuleVersion = '0.10.0'

# ID used to uniquely identify this module
GUID = 'c04e8722-3530-4700-ba8c-03123835738f'

Author = 'Cellmate Author'
CompanyName = ''
Copyright = '(c) 2020 Cellmate Author. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Cellmate is a collection of PowerShell cmdlets for processing Excel workbooks.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.1'

# Cmdlets to export from this module
CmdletsToExport = @(
    'Clear-DateCell',
    'Compress-Workbook',
    'Edit-DateCell',
    'Export-Workbook',
    'Import-Workbook',
    'Merge-BorderedWorksheetRange',
    'Merge-Workbook',
    'Remove-HiddenWorksheet'
    'Remove-Worksheet'
    'Test-DateCell'
)
}
