param (
    [Switch]$PackForm,
    [Switch]$CheckForm,
    [Switch]$ErrorForm,
    [System.String]$OutputPath,
    [System.String]$RootPath = 'C:/PyTkinterToPSScript',
    [System.String]$DataMapping_Header = 'DataMapping_Header.csv',
    [System.String]$DataMapping_Body = 'DataMapping_Body.csv',
    [System.String]$PackForm_Template = 'Template_PackFormLists.xlsx',
    [System.String]$PackForm_HeaderValues = 'Packagelist_HeaderValues.csv',
    [System.String]$PackForm_BodyValues = 'Packagelist_BodyValues.csv',
    [System.String]$CheckForm_Template = 'Template_CheckFormLists.xlsx',
    [System.String]$CheckForm_HeaderValues = 'Checklist_HeaderValues.csv',
    [System.String]$CheckForm_BodyValues = 'Checklist_BodyValues.csv',
    [System.String]$ErrorForm_Template = 'Template_ErrorFormLists.xlsx',
    [System.String]$ErrorForm_HeaderValues = 'Checklist_HeaderValues.csv',
    [System.String]$ErrorForm_BodyValues = 'Checklist_ZipFileList.csv'
)

# メイン処理
$statusCode = 0

# 使用するクラスを読み込み
$packageListConfig = "$(($PSScriptRoot).Replace('\', '/'))/PSClasses/PrintListConfig.ps1"
# 使用するモジュールを読み込む
$commonModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/CommonModules.psm1"
$printModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/PrintModules.psm1"
if (-not (Test-Path $packageListConfig) -or
    -not (Test-Path $commonModules) -or
    -not (Test-Path $printModules)) {
    $statusCode = -8001
    Write-Error "必要な外部モジュールファイル（*.ps1, *.psm1）が存在しません。: $($_.Exception.Message)"
}
else {
    try {
        # カスタムクラスを読み込み
        # Import-Module $packageListConfig -Force
        . $packageListConfig
        [System.String]$TEMPLATESHEET1 = [PrintListConfig]::TEMPLATESHEET1
        [System.String]$TEMPLATESHEET2 = [PrintListConfig]::TEMPLATESHEET2

        # 共通で使用する関数を読み込み
        Import-Module $commonModules -Force
        
        # Adpackで使用する関数を読み込み
        Import-Module $printModules -Force
    }
    catch {
        Write-Error "外部モジュールファイル（*.psm1）の読み込み処理でエラーが発生しました。: $($_.Exception.Message)"
        $statusCode = -8002
    }
}

# 引数の論理チェック
if ($statusCode -eq 0) {
    $switchCount = @($PackForm, $CheckForm, $ErrorForm | Where-Object { $_ }).Count
    if ($switchCount -gt 1) {
        Write-Error "引数を見直してください。[理由：複数のフォームを同時に設定]"
        $statusCode = -8003
    }
    elseif ($switchCount -eq 0) {
        Write-Error "引数を見直してください。[理由：いずれのフォームも未設定]"
        $statusCode = -8004
    }
}

# 各種フォルダーやファイルのパスを宣言・引数チェック
if ($statusCode -eq 0) {
    # パスを宣言
    #   テンプレートファイル
    $TemplatePath = "$($RootPath)/template"
    $packFormPath = "$($TemplatePath)/$($PackForm_Template)"
    $checkFormPath = "$($TemplatePath)/$($CheckForm_Template)"
    $errorFormPath = "$($TemplatePath)/$($ErrorForm_Template)"
    $headerMappingPath = "$($TemplatePath)/$($DataMapping_Header)"
    $bodyMappingPath = "$($TemplatePath)/$($DataMapping_Body)"
    #   入力ファイル
    $InputPath = "$($RootPath)/input"
    $packform_HeaderValuesPath = "$($InputPath)/$($PackForm_HeaderValues)"
    $packform_BodyValuesPath = "$($InputPath)/$($PackForm_BodyValues)"
    $checkform_HeaderValuesPath = "$($InputPath)/$($CheckForm_HeaderValues)"
    $checkform_BodyValuesPath = "$($InputPath)/$($CheckForm_BodyValues)"
    $errorform_HeaderValuesPath = "$($InputPath)/$($ErrorForm_HeaderValues)"
    $errorform_BodyValuesPath = "$($InputPath)/$($ErrorForm_BodyValues)"

    # 引数のチェック
    if ($PackForm) {
        $formPath = $packFormPath
        $headerValuesPath = $packform_HeaderValuesPath
        $bodyValuesPath = $packform_BodyValuesPath
    }
    elseif ($CheckForm) {
        $formPath = $checkFormPath
        $headerValuesPath = $checkform_HeaderValuesPath
        $bodyValuesPath = $checkform_BodyValuesPath
    }
    elseif ($ErrorForm) {
        $formPath = $errorFormPath
        $headerValuesPath = $errorform_HeaderValuesPath
        $bodyValuesPath = $errorform_BodyValuesPath
    }
}

# 各種フォルダーやファイルの存在チェック
if ($statusCode -eq 0) {
    if (-not (Test-PathType $RootPath -PathRole Source -PathType Container)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$RootPath]"
        $statusCode = -8005
    }
    elseif (-not (Test-PathType $TemplatePath -PathRole Source -PathType Container)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$TemplatePath]"
        $statusCode = -8006
    }
    elseif (-not (Test-Path $formPath)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$formPath]"
        $statusCode = -8007
    }
    elseif (-not (Test-Path $headerMappingPath)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$headerMappingPath]"
        $statusCode = -8008
    }
    elseif (-not (Test-Path $bodyMappingPath)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$bodyMappingPath]"
        $statusCode = -8009
    }
    elseif (-not (Test-Path $headerValuesPath)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$headerValuesPath]"
        $statusCode = -8010
    }
    elseif (-not (Test-Path $bodyValuesPath)) {
        Write-Error "引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：$bodyValuesPath]"
        $statusCode = -8011
    }
}

# Excelテンプレートファイル内のシート存在チェック
if ($statusCode -eq 0) {
    if (-not (Test-ExcelSheetExists $formPath $TEMPLATESHEET1) -and
            -not (Test-ExcelSheetExists $formPath $TEMPLATESHEET2)) {
            Write-Error "帳票テンプレートファイル（Excel）内に既定のシートがありませんでした。[対象ファイル: $($formPath), シート1: $([PrintListConfig]::TEMPLATESHEET1), シート2: $($TEMPLATESHEET2)]"
            $statusCode = -8012
    }
}

# Excelテンプレートファイルを一時フォルダーにコピー
if ($statusCode -eq 0) {
    # C:/Users/XXX/AppData/Local/Temp/PyTkInterToPSScript_ExcelTemplaete.xlsx
    $excelWorkfilePath_format = "$(($Env:TEMP).Replace('\', '/'))/PyTkInterToPSScript_ExcelTemplaete.xlsx"

    # すでに一時フォルダーにある場合は削除
    try {
        Remove-File $excelWorkfilePath_format
        # 削除できた場合は、そのままのファイル名で続行
        $excelWorkfilePath = $excelWorkfilePath_format
    }
    catch {
        # 何らかの理由で削除できなかった場合は、ファイル名を一意に変更
        $excelWorkfilePath = Get-UniqueFilePath $excelWorkfilePath_format
    }

    $copyFrom = $formPath
    $copyTo = $excelWorkfilePath
    try {
        Copy-Item $copyFrom $copyTo -Force
    }
    catch {
        Write-Error "帳票テンプレートファイル（Excel）のコピー処理でエラーが発生しました。[コピー元: $($copyFrom), コピー先: $($copyTo)]"
        $statusCode = -8013
    }
}

# ヘッダー情報とメイン情報の読み込み
if ($statusCode -eq 0) {
    # ヘッダー情報の読み込みとチェック
    try {
        # ヘッダー情報の入力位置データ
        $headerMapping = @(Import-Csv -Path $headerMappingPath)
        # ヘッダー情報の入力データ
        $headerValue = @(Import-Csv -Path $headerValuesPath)
    }
    catch {
        Write-Error "ヘッダー情報の位置データと入力データの読み込みでエラーが発生しました。"
        $statusCode = -8014
    }

    if ($statusCode -eq 0) {
        if (-not (Test-PSCustomObjectEquality $headerMapping $headerValue)) {
            Write-Error 'ヘッダー情報における 位置データ と 入力データ の項目名が一致しません。'
            $statusCode = -8015
        }
    }
}

# データ本体の読み込みとチェック
if ($statusCode -eq 0) {
    try {
        # ヘッダー情報の入力位置データ
        $bodyMapping = @(Import-Csv -Path $bodyMappingPath)
        # ヘッダー情報の入力データ
        $bodyValue = @(Import-Csv -Path $bodyValuesPath)
    }
    catch {
        Write-Error "メイン情報の登録内、データを読み込みでエラーが発生しました。"
        $statusCode = -8016
    }

    if ($statusCode -eq 0) {
        if (-not (Test-PSCustomObjectEquality $bodyMapping $bodyValue)) {
            Write-Error 'データ本体の書き込み位置データ と 入力データ の項目名が一致しません。'
            $statusCode = -8017
        }
    }
}

#   Excelファイルに値を設定
if ($statusCode -eq 0) {
    $config = [PrintListConfig]::new($excelWorkfilePath, $headerMapping, $headerValue, $bodyMapping, $bodyValue)
    try {
        Set-PackagelistValues -Config $config
    }
    catch {
        Write-Error "帳票ファイル（Excel）に入力データを設定する処理でエラーが発生しました。[対象ファイル: $($excelWorkfilePath)]"
        Write-Error "エラー情報 [詳細: $($_.Exception.Message), 場所: $($_.InvocationInfo.MyCommand.Name)]"
        $statusCode = -8018
    }
}

# 印刷処理
if ($statusCode -eq 0) {
    #   Excelファイルのテンプレートシートを削除
    try {
        $removeSheet = @($TEMPLATESHEET1, $TEMPLATESHEET2)
        Remove-ExcelSheets $excelWorkfilePath $removeSheet
    }
    catch {
        Write-Error "Excelファイル テンプレートシートの削除処理でエラーが発生しました。[対象ファイル: $($excelWorkfilePath), 対象シート: $($removeSheet -join ', ')]"
        $statusCode = -8019
    }
}

#   PDFファイルを出力処理
if ($statusCode -eq 0) {
    try {
        Remove-Item $OutputPath -Force
        Export-ExcelDocumentAsPDF $excelWorkfilePath $OutputPath
    }
    catch {
        Write-Error "PDFファイルの出力処理でエラーが発生しました。[対象ファイル: $($excelWorkfilePath), 出力先: $($OutputPath)]"
        $statusCode = -8020
    }
}

exit $statusCode
