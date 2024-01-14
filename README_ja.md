# Cellmate

Cellmate は、Microsoft Excel のワークブックを処理するための [PowerShell モジュール] です。

[PowerShell モジュール]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_modules?view=powershell-5.1

## 動作要件
* Windows PowerShell 5.1
* .NET Framework 4.8
* Microsoft Excel

PowerShell 6 および .NET Core は、この PowerShell モジュールではサポートされていないことに注意してください。

## インストール方法

配布用の zip ファイルの最新の安定バージョンは、このリポジトリの [Release](https://github.com/openclosed-dev/cellmate/releases) ページからダウンロードできます。

1. アクティブな PowerShell セッションが存在する場合は閉じます。
2. ダウンロードした zip ファイル `Cellmate-<バージョン>.zip` を解凍します。
3. 解凍した `Cellmate` ディレクトリを、アカウントの `\Users\<ユーザー名>\Documents\WindowsPowerShell\Modules` にコピーします。 結果のディレクトリは `\Users\<ユーザー名>\Documents\WindowsPowerShell\Modules\Cellmate\<バージョン>` になります。

## コードサンプル

このセクションでは、例として PowerShell スクリプトを示します。

### ワークブックを PDF ファイルに結合する
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

### ワークブックを ZIP ファイルにアーカイブする
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

## 提供されているコマンドレットのリスト

| 名前 | 説明 |
| --- | --- |
| Compress-Workbook | 1つ以上のワークブックを含む ZIP アーカイブを作成します。 |
| Export-Workbook | ワークブックをファイルに保存します。  |
| Import-Workbook | 指定されたパスからワークブックを読み込みます。 |
| Merge-Workbook | 1つ以上のワークブックを PDF ファイルに結合します。 |

## 法的通知
Copyright 2020-2024 the original author or authors. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this product except in compliance with the License.
You may obtain a copy of the License at
<http://www.apache.org/licenses/LICENSE-2.0>

