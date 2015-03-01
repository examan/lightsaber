for /f "tokens=3*" %%x in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\devenv.exe"') do set DEVENV="%%x %%y"

powershell -ExecutionPolicy ByPass -File VersionUpdater.ps1 Version.txt
powershell -ExecutionPolicy ByPass -File InstallerProductVersionUpdater.ps1 Version.txt ..\Source\Setup86\Setup86.vdproj
powershell -ExecutionPolicy ByPass -File InstallerProductVersionUpdater.ps1 Version.txt ..\Source\Setup64\Setup64.vdproj

mkdir Release
%DEVENV% /Out x86_build.log /Rebuild "Release|x86" ..\Source\Lightsaber.sln
copy ..\Source\Setup86\Release\Lightsaber86.msi Release
%DEVENV% /Out x64_build.log /Rebuild "Release|x64" ..\Source\Lightsaber.sln
copy ..\Source\Setup64\Release\Lightsaber64.msi Release
