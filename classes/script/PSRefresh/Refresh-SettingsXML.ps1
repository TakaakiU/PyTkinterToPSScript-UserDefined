$DebugPreference = 'Continue'

$TargetFolder = "C:\PyTkinterToPSScript\config"
$SourceFolder = "C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript-UserDefined\classes\config"
$ExcludeFolders = @("$SourceFolder\.old")
$TargetFileExtensions = @("*.xml")

$confirmMessages = Read-Host @"

■ Refresh-SettingsXML 確認メッセージ
　 テスト環境のスクリプト格納フォルダー「$TargetFolder」にある設定ファイル（settings.xml）を削除し、
　 最新ファイル「$SourceFolder」をコピーします。
　
　 続行しますか？［Y / N］
"@

if ($confirmMessages -match "^[Yy]$") {
    if (-not (Test-Path $TargetFolder)) {
        Write-Host "設定ファイルの格納フォルダー「$TargetFolder」が存在しないため作成します..."
        New-Item -Path $TargetFolder -ItemType Directory | Out-Null
    }

    Write-Host "設定ファイルの格納フォルダー「$TargetFolder」にあるXMLファイルを削除します。"
    Get-ChildItem -Path $TargetFolder -Include $TargetFileExtensions -Recurse -File | ForEach-Object {
        Write-Debug "Delete Item:$($_.FullName)"
        Remove-Item $_.FullName -Force
    }

    Write-Host "最新ファイル「$SourceFolder」を設定ファイルの格納先にコピーします。"
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
