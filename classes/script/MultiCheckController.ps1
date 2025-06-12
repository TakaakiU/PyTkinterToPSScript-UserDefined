param (
    [System.String]$InputPath,
    [System.String]$OutputPath
)

$statusCode = 0

# 使用する関数を読み込む
$adpackController = "$(($PSScriptRoot).Replace('\', '/'))/AdpackController.ps1"
$commonModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/CommonModules.psm1"
$adpackModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/AdpackModules.psm1"
if (-not (Test-Path $adpackController) -or -not (Test-Path $commonModules) -or -not (Test-Path $adpackModules)) {
    Write-Error "必要な外部モジュールファイル（*.ps1, *.psm1）が存在しません。"
    $statusCode = -7201
}
else {
    try {
        # 共通で使用する関数を読み込み
        Import-Module $commonModules
        # Adpackで使用する関数を読み込み
        Import-Module $adpackModules
    }
    catch {
        Write-Error "外部モジュールファイル（*.ps1, *.psm1）の読み込み処理でエラーが発生しました。: $($_.Exception.Message)"
        $statusCode = -7202
    }
}

# 引数ごとの論理チェック
if ($statusCode -eq 0) {
    # 入力データのパスはフォルダーか判断
    if (-not (Test-PathType -Path $InputPath -PathRole Source -PathType Container)) {
        Write-Error "入力データは「フォルダー」を指定してください。"
        $statusCode = -7203
    }
}

# # PowerShellバージョン確認
# if ($statusCode -eq 0) {
#     $statusCode = Test-PowerShellVersion7OrLater
# }

# # Adpack実行ファイルの存在チェック
# if ($statusCode -eq 0) {
#     $exePath = "$(($PSScriptRoot).Replace('\', '/'))/../exe/adpack.exe"
#     if (-not (Test-Path $exePath)) {
#         Write-Error "Adpackの実行ファイルが存在しません。"
#         $statusCode = -7204
#     }
# }

# 引数に応じて実行
if ($statusCode -eq 0) {
    # ZIPファイルの存在確認 と ZIPファイルリストの作成
    $zipFiles = Get-ChildItem -Path $InputPath -Filter "*.zip" | Select-Object -ExpandProperty FullName
    if (-not $zipFiles) {
        Write-Error "指定されたフォルダーにZIPファイルが存在しません。"
        $statusCode = -7205
    }
}

# ZIPファイルリストに従い繰り返しチェック
if ($statusCode -eq 0) {
    # AdpackController.ps1のUnPackを繰り返し行う。
    $results = @()
    $number = 1
    foreach ($zipFile in $zipFiles) {
        Write-Host "チェックファイル: $zipFile"
        # $process = Start-Process -FilePath "powershell" -ArgumentList "-File $adpackController -Unpack -InputPath `"$zipFile`"" -NoNewWindow -Wait -PassThru
        $process = Start-Process -FilePath "pwsh" -ArgumentList "-File $adpackController -Unpack -InputPath `"$zipFile`"" -NoNewWindow -Wait -PassThru
        $statusCode = $($process.ExitCode)

        # $results += [PSCustomObject]@{
        #     FilePath = $zipFile
        #     StatusCode = $statusCode
        # }
        $results += [PSCustomObject]@{
            "No." = $number
            "Index" = $zipFile.Replace("$($InputPath.Replace('/', '\'))\", "")
            "HashValue" = $statusCode
        }
        $number++

        # 異常が発生した場合は中断
        if ($statusCode -ne 0) {
            break
        }
    }

    # リストを外部出力
    $results | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation
}

if ($statusCode -eq 0) {
    Write-Host "正常終了しました。"
}
else {
    Write-Host "異常終了しました。[リザルトコード:$statusCode]"
}

exit $statusCode
