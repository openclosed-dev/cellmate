Remove-Item cellmate.zip -ErrorAction Ignore
Remove-Item src\Cellmate\bin\Release -Recurse -ErrorAction Ignore

cd src\Cellmate
dotnet clean -c Release
dotnet publish -c Release
cd ..\..

Move-Item src\Cellmate\bin\Release\net47\win-x64\publish src\Cellmate\bin\Release\net47\win-x64\cellmate
Compress-Archive -Path src\Cellmate\bin\Release\net47\win-x64\cellmate -DestinationPath cellmate.zip
