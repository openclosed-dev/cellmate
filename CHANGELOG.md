# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased
### Added
- `Compress-Workbook` cmdlet now receives the following parameters.
    - `-FileMode` to specify the file mode of the archive file.
      See https://learn.microsoft.com/en-us/dotnet/api/system.io.filemode?view=netframework-4.8.1
    - `-LastWriteTime` to specify the date time of the archive entries.

### Changed
- `Compress-Workbook` cmdlet now creates a new archive file by default.
  This behavior can be changed using `-FileMode` parameter.

## 0.9.0 - 2024-01-21

### Added
- Added a new cmdlet `Remove-HiddenWorksheet`.

## 0.8.0 - 2024-01-08
### Changed
- Updated .NET Framework to 4.8.

### Added
- Added the following new parameters to `Merge-Workbook` cmdlet.
    - `-RestartPageNumber`
    - `-PageNumberFormat`

## 0.7.0 - 2020-03-29
### Changed
- `Merge-Cell` cmdlet was renamed to `Merge-BorderedWorksheetRange`.

### Added
- Added `SkipBooks` parameter to the cmdlets processing workbooks.

### Fixed
- Fix resolution of path passed to `Destination` parameter when the path is a UNC.

## 0.6.0 - 2020-03-24
### Added
- Merge-Cell cmdlet.

## 0.5.0 - 2020-03-22
### Added
- `Compress-Workbook` cmdlet for compressing input workbooks.
- `Clear-DateCell` cmdlet.

### Changed
- Rename cmdlets `Import-Excel`/`Export-Excel`/`Merge-Excel` to
`Import-Workbook`/`Export-Workbook`/`Merge-Workbook` respectively.
- `PageNumber` parameter of `Merge-Workbook` now takes `left`/`center`/`right` as its value.
- DLL name was shortened to `Cellmate.dll`.

## 0.4.0 - 2020-03-17
### Added
- Make cell testing much faster.
- Improve page counting.

## 0.3.0 - 2020-03-17
### Changed
- `Path` parameter for `Merge-Excel` was renamed to `Destination`.

### Fixed
- A bug in cell range calculation.

## 0.2.0 - 2020-03-15
### Changed
- Now provided as a PowerShell Module composed of cmdlets for processing Excel.

## 0.1.0 - 2020-02-24
- First release.
