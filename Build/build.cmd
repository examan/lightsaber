for /f "tokens=3*" %%x in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\devenv.exe"') do set DEVENV="%%x %%y"
mkdir Release
%DEVENV% /Rebuild "Release|x86" ..\Source\Lightsaber.sln
copy ..\Source\Setup86\Release\Lightsaber86.msi Release
%DEVENV% /Rebuild "Release|x64" ..\Source\Lightsaber.sln
copy ..\Source\Setup64\Release\Lightsaber64.msi Release
