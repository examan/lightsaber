for /f "tokens=3*" %%x in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\devenv.exe"') do set DEVENV="%%x %%y"
cd ../Source
%DEVENV% /Rebuild "Release|x86" Lightsaber.sln
%DEVENV% /Rebuild "Release|x64" Lightsaber.sln
cd ../Build
dotNetInstaller\installerLinker /Output:Release\Lightsaber.exe /Template:dotNetInstaller/dotNetInstaller.exe /Configuration:dotNetInstaller.configuration /Verbose+
pause
