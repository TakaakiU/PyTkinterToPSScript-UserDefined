$DebugPreference = 'Continue'

$rootFolder = "C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript-UserDefined\classes\script\PSRefresh"
$scriptFile1 = "Refresh-CDrivePSScript.ps1"
$scriptFile2 = "Refresh-ExecutableFile.ps1"
$scriptFile3 = "Refresh-SettingsXML.ps1"

$scriptFilePath1 = "$rootFolder\$scriptFile1"
$scriptFilePath2 = "$rootFolder\$scriptFile2"
$scriptFilePath3 = "$rootFolder\$scriptFile3"

$scriptLists = @($scriptFilePath1, $scriptFilePath2, $scriptFilePath3)

foreach ($script in $scriptLists) {
    Invoke-Expression "& '$script'"
}

Write-Host "すべてのスクリプトの更新が完了しました。"

$DebugPreference = 'SilentlyContinue'
