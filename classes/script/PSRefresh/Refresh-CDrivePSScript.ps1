$DebugPreference = 'Continue'

$TargetFolder = "C:\PyTkinterToPSScript\script"
$PSClassesFolder = "C:\PyTkinterToPSScript\script\PSClasses"
$PSModulesFolder = "C:\PyTkinterToPSScript\script\PSModules"
$SourceFolder = "C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript-UserDefined\classes\script"
$ExcludeFolders = @("$SourceFolder\.old", "$SourceFolder\PSRefresh")
$TargetFileExtensions = @("*.ps1", "*.psm1")

$confirmMessages = Read-Host @"

■ Refresh-CDrivePSScript 確認メッセージ
　 テスト環境のスクリプト格納フォルダー「$TargetFolder」にあるPowerShellスクリプト（*.ps1）を削除し、
　 最新ソース「$SourceFolder」のPowerShellスクリプトをコピーします。
　
　 続行しますか？［Y / N］
"@

if ($confirmMessages -match "^[Yy]$") {
    if (-not (Test-Path $TargetFolder)) {
        Write-Host "スクリプト格納フォルダー「$TargetFolder」が存在しないため作成します..."
        New-Item -Path $TargetFolder -ItemType Directory | Out-Null
    }
    if (-not (Test-Path $PSClassesFolder)) {
        Write-Host "スクリプト格納フォルダー「$PSClassesFolder」が存在しないため作成します..."
        New-Item -Path $PSClassesFolder -ItemType Directory | Out-Null
    }
    if (-not (Test-Path $PSModulesFolder)) {
        Write-Host "スクリプト格納フォルダー「$PSModulesFolder」が存在しないため作成します..."
        New-Item -Path $PSModulesFolder -ItemType Directory | Out-Null
    }

    Write-Host "スクリプト格納フォルダー「$TargetFolder」にあるPowerShellスクリプトを削除します。"
    Get-ChildItem -Path $TargetFolder -Include $TargetFileExtensions -Recurse -File | ForEach-Object {
        Write-Debug "Delete Item:$($_.FullName)"
        Remove-Item $_.FullName -Force
    }

    Write-Host "最新ソース格納フォルダー「$SourceFolder」をスクリプト格納先にコピーします。"
    Get-ChildItem -Path $SourceFolder  -Include $TargetFileExtensions -Recurse -File | Where-Object {
        $FileParentFolder = $_.DirectoryName
        $ExcludeFolders -notcontains $FileParentFolder
    } | ForEach-Object {
        $targetPath = $_.FullName -replace [regex]::Escape($SourceFolder), $TargetFolder
        if ($targetPath -ne $_.FullName) {
            Write-Debug "Copy From:$($_.FullName)"
            Write-Debug "Copy To  :$targetPath"
            Copy-Item -Path $_.FullName -Destination $targetPath -Force
        }
    }

    Write-Host "スクリプトの更新が完了しました。"
}
else {
    Write-Host "処理をキャンセルしました。"
}

$DebugPreference = 'SilentlyContinue'
