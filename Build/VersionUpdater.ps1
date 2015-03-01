param(
	[string]$versionFilePath = $(throw "Mandatory parameter -versionFilePath not supplied.")
)

$originalVersion = [IO.File]::ReadAllText($versionFilePath).trim()
$originalComponents = $originalVersion.Split(".")

$newBuild = [int32](((get-date).Year-2000)*366)+(Get-Date).DayOfYear
$newRevision = [int32](((get-date)-(Get-Date).Date).TotalSeconds / 2)
$newComponents = $originalComponents[0], $originalComponents[1], $newBuild, $newRevision

$newVersion = "$($newComponents[0]).$($newComponents[1]).$($newComponents[2]).$($newComponents[3])"
set-Content $versionFilePath $newVersion -encoding "UTF8"
