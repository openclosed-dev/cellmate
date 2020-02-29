$xml = [xml](Get-Content Cellmate.csproj)
$name = "cellmate"
$version = $xml.Project.PropertyGroup.version
$outdir = "bin\Release\net47"

dotnet clean -c Release
dotnet publish -c Release -o $outdir\$name

Compress-Archive -Path $outdir\$name -DestinationPath $outdir\$name-$version.zip
