param (
    [Switch]$Pack,
    [Switch]$Unpack,
    # [System.String]$Hash = "s256",      # adpack.exeを使用する場合
    [System.String]$Hash = "SHA256",  # 自作関数を使用する場合
    [System.String]$InputPath,
    [System.String]$OutputPath = "",
    [Switch]$Check,
    [Switch]$NoCheck
)

$statusCode = 0

# 使用する関数を読み込む
$commonModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/CommonModules.psm1"
$adpackModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/AdpackModules.psm1"
if (-not (Test-Path $commonModules) -or -not (Test-Path $adpackModules)) {
    Write-Error "必要な外部モジュールファイル（*.psm1）が存在しません。"
    $statusCode = -7001
}
else {
    try {
        # 共通で使用する関数を読み込み
        Import-Module $commonModules
        # Adpackで使用する関数を読み込み
        Import-Module $adpackModules
    }
    catch {
        Write-Error "外部モジュールファイル（*.psm1）の読み込み処理でエラーが発生しました。: $($_.Exception.Message)"
        $statusCode = -7002
    }
}

# 引数の論理チェック
if ($statusCode -eq 0) {
    if ($Pack -and $Unpack) {
        Write-Error "引数を見直してください。[理由：Pack + UnPack 両方を設定]"
        $statusCode = -7003
    }
    elseif (-not $Pack -and -not $Unpack) {
        Write-Error "引数を見直してください。[理由：Pack と UnPack どちらも未設定]"
        $statusCode = -7004
    }
    elseif ($Pack -and ($Check -or $NoCheck)) {
        Write-Error "引数を見直してください。[理由：Pack + Check、もしくは Pack + UnCheck で設定]"
        $statusCode = -7005
    }
    elseif ($Unpack -and ($Check -and $NoCheck)) {
        Write-Error "引数を見直してください。[理由：UnPack + Check + UnCheck 3つを設定]"
        $statusCode = -7006
    }
}

# 引数ごとの論理チェック
if ($statusCode -eq 0) {
    if ($Pack) {
        # 入力データのパスはフォルダーか判断
        if (-not (Test-PathType -Path $InputPath -PathRole Source -PathType Container)) {
            Write-Error "パックでの入力データは「フォルダー」を指定してください。"
            $statusCode = -7007
        }
        else {
            # 出力データ未指定の場合は、自動で入力データから出力先を指定
            if ($OutputPath -eq "") {
                $OutputPath = "$($InputPath).zip"
            }
            # 出力データのパスはZIPファイルか判断
            else {
                if (-not (Test-PathType -Path $OutputPath -PathRole Target -PathType Leaf -Extension zip)) {
                    Write-Error "パックでの出力データは「ファイル（*.zip）」を指定してください。"
                    $statusCode = -7008
                }
            }
        }
    }
    elseif ($Unpack) {
        # 入力データのパスはZIPファイルか判断
        if (-not (Test-PathType -Path $InputPath -PathRole Source -PathType Leaf -Extension zip)) {
            Write-Error "アンパックでの入力データは「ファイル（*.zip）」を指定してください。"
            $statusCode = -7009
        }
        if ($statusCode -eq 0) {
            # 出力データ未指定の場合は、自動で入力データから出力先を指定
            if ($OutputPath -eq "") {
                # # 入力データの格納フォルダーを取得
                # $parentPath = (Split-Path -Path $InputPath)
                # # 入力ファイルの拡張子を除いたファイル名を取得
                # $baseName = (Get-Item $InputPath).BaseName
                # # 出力フォルダーを設定
                # $OutputPath = Join-Path -Path $parentPath -ChildPath $baseName
                # 入力データの格納フォルダーを出力フォルダーとして設定
                $OutputPath = (Split-Path -Path $InputPath)
            }
            # 出力データのパスはフォルダーか判断
            elseif (-not (Test-PathType -Path $OutputPath -PathRole Target -PathType Container)) {
                Write-Error "アンパックでの出力データは「フォルダー」を指定してください。"
                $statusCode = -7010
            }
        }
    }
}

# # PowerShellバージョン確認
# if ($statusCode -eq 0) {
#     $statusCode = Test-PowerShellVersion7OrLater
# }

# Adpack実行ファイルの存在チェック
if ($statusCode -eq 0) {
    $exePath = "$(($PSScriptRoot).Replace('\', '/'))/../exe/adpack.exe"
    if (-not (Test-Path $exePath)) {
        Write-Error "Adpackの実行ファイルが存在しません。"
        $statusCode = -7011
    }
}

# 引数に応じて実行
if ($statusCode -eq 0) {
    # 問題ない引数は実行
    if ($Pack) {
        # $statusCode = (Compress-Package_Adpack -ExePath $exePath -HashAlgorithm $Hash -FolderPath $InputPath -ZipFilePath $OutputPath)
        $statusCode = (Compress-Package_UserDefined -HashAlgorithm $Hash -FolderPath $InputPath -ZipFilePath $OutputPath)
    }
    elseif ($Unpack) {
        # Unpack (解凍 と チェック)
        if (!$Check -And !$NoCheck) {
            # $statusCode = (Expand-Package_Adpack -ExePath $exePath -HashAlgorithm $Hash -ZipFilePath $InputPath -FolderPath $OutputPath)
            $statusCode = (Expand-Package_UserDefined -HashAlgorithm $Hash -ZipFilePath $InputPath -FolderPath $OutputPath)
        }
        # Unpack + Check (解凍 → チェック → 削除。チェック結果のみ)
        elseif ($Check) {
            # $statusCode = (Expand-Package_Adpack -Check -ExePath $exePath  -HashAlgorithm $Hash -ZipFilePath $InputPath -FolderPath $OutputPath)
            $statusCode = (Expand-Package_UserDefined -Check -HashAlgorithm $Hash -ZipFilePath $InputPath -FolderPath $OutputPath)
        }
        # Unpack + NoCheck (解凍)
        elseif ($NoCheck) {
            # $statusCode = (Expand-Package_Adpack -NoCheck -ExePath $exePath -HashAlgorithm $Hash -ZipFilePath $InputPath -FolderPath $OutputPath)
            $statusCode = (Expand-Package_UserDefined -NoCheck -HashAlgorithm $Hash -ZipFilePath $InputPath -FolderPath $OutputPath)
        }
    }
}

if ($statusCode -eq 0) {
    Write-Host "正常終了しました。"
}
else {
    Write-Host "異常終了しました。[リザルトコード:$statusCode]"
}

exit $statusCode
