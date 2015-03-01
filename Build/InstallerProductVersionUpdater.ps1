param(
	[string]$versionFilePath = $(throw "Mandatory parameter -versionFilePath not supplied."),
	[string]$installerProjectFilePath = $(throw "Mandatory parameter -installerProjectFilePath not supplied.")
)

$originalContent = [IO.File]::ReadAllText($installerProjectFilePath)
$originalVersion = [IO.File]::ReadAllText($versionFilePath).trim()

$newVersion = $originalVersion -replace "(\d+\.\d+\.\d+)\.\d+", "`$1"
$newContent = $originalContent -replace "`"ProductVersion`" = `"8:.*?`"", "`"ProductVersion`" = `"8:$newVersion`""

[IO.File]::WriteAllText($installerProjectFilePath, $newContent)
