#
# Copyright 2020 the original author or authors.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

if ($PSVersionTable.PSVersion.Major -ne '5') {
    Write-Error "PowerShell 5.x is required to execute."
    exit 1
}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$name = "Cellmate"
$projectDir = "$here\Cellmate"
$project = "$projectDir\Cellmate.csproj"
$testProjectDir = "$here\Cellmate.Test"

$xml = [xml](Get-Content $project)
$version = $xml.Project.PropertyGroup.version[0]

$binDir = "$projectDir\bin"
$destDir = "$binDir\$name\$version"
$archive = "$here\$name-$version.zip"

if (Test-Path $archive) {
    Remove-Item $archive
}

if (Test-Path $binDir) {
    Remove-Item -Force -Recurse $binDir
}

dotnet build $project -c Release

# Runs unit tests
PUsh-Location $testProjectDir
powershell -File .\TestAll.ps1 -Configuration Release
Pop-Location

dotnet publish $project -c Release -o $destDir

Copy-Item -Path "$here\README.md" -Destination $destDir
Copy-Item -Path "$here\README_ja.md" -Destination $destDir
Copy-Item -Path "$here\LICENSE" -Destination $destDir
Copy-Item -Path "$here\CHANGELOG.md" -Destination $destDir
Copy-Item -Path "$here\NOTICE.md" -Destination $destDir
Copy-Item -Path "$projectDir\Cellmate.psd1" -Destination $destDir

Compress-Archive -Path "$projectDir\bin\$name" -DestinationPath $archive
