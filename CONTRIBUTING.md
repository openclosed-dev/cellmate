# Contribution Guide

## How to Buid

### Prerequites

* PowerShell 5.x
* .NET 8.0 SDK (v8.0.100)
* .NET Framework 4.8 Developer Pack
* Pester 5.5

Pester can be install using following command.

```
Install-Module -Name Pester -Scope CurrentUser 
```

Confirm the installation result.

```
Import-Module Pester -Passthru

ModuleType Version    Name                                ExportedCommands
---------- -------    ----                                ----------------
Script     5.5.0      Pester                              {Add-ShouldOperator, AfterAll, AfterEach, Assert-MockCall...
```

### Build for release

Run the build script in the repository root.

```
.\build.ps1
```
