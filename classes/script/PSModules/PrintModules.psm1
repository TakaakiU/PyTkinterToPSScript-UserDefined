enum XlLookAt {
    xlWhole = 1
    xlPart = 2
}
enum XlFixedFormatType {
    xlTypePDF = 0
    xlTypeXPS = 1
}

# 使用するクラスを読み込み
$packageListConfig = "$(($PSScriptRoot).Replace('\', '/'))/../PSClasses/PrintListConfig.ps1"
if (-not (Test-Path $packageListConfig)) {
    Write-Error "クラス・モジュールを読み込む際、いずれかのpsmファイルが存在しません。: $($_.Exception.Message)"
}
. $packageListConfig
[System.String]$TEMPLATESHEET1 = [PrintListConfig]::TEMPLATESHEET1
[System.Int32]$TEMPLATEROWS1 = [PrintListConfig]::TEMPLATEROWS1
[System.String]$TEMPLATESHEET2 = [PrintListConfig]::TEMPLATESHEET2
[System.Int32]$TEMPLATEROWS2 = [PrintListConfig]::TEMPLATEROWS2
[System.String]$HEADERRANGE = [PrintListConfig]::HEADERRANGE
[System.String]$MAINRANGE = [PrintListConfig]::MAINRANGE

# 個別のFunciton
Function Test-ExcelSheetExists {
    param(
        [System.String]$Path,
        [System.String]$CheckSheet
    )

    $excelApp = $null
    $workBooks = $null
    $workBook = $null
    $workSheets = $null
    $workSheet = $null

    $sheetExists = $false
    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)
        $workSheets = $workBook.Sheets

        foreach ($workSheet in $workSheets) {
            if ($workSheet.Name -eq $CheckSheet) {
                $sheetExists = $true
                break
            }
        }
    }
    catch {
        Write-Error 'Excel操作中にエラーが発生しました。'
    }
    finally {
        # ワークブックまで解放
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # ワークブックの保存しないで終了
            $workBook.Close($false)

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Excelアプリ終了
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }

    return $sheetExists
}

# ファイルのロック状態をチェック
function Test-FileLocked {
    param (
        [Parameter(Mandatory=$true)][System.String]$Path
    )

    if (-Not(Test-Path $Path)) {
        Write-Error '対象ファイルが存在しません。' -ErrorAction Stop
    }

    # 相対パスだとOpenメソッドが正常動作しない為、絶対パスに変換
    $fullPath = (Resolve-Path -Path $Path -ErrorAction SilentlyContinue).Path

    $fileLocked = $false
    try {
        # 読み取り専用でファイルを開く処理を実行
        $fileStream = [System.IO.File]::Open($fullPath, 'Open', 'ReadWrite', 'None')
    }
    catch {
        # ファイルが開けない場合、ロック状態と判断
        $fileLocked = $true
    }
    finally {
        if ($null -ne $fileStream) {
            $fileStream.Close()
        }
    }

    return $fileLocked
}

# 正しいシート名かチェック
function Test-ExcelSheetname {
    param(
        [Parameter(Mandatory=$true)][System.String]$SheetName
    )
    
    # 名前が空白でないかチェック
    if ([string]::IsNullOrWhiteSpace($SheetName.Trim())) {
        Write-Warning '空白である（Nullや空文字を含む）'
        return $false
    }

    # 文字数が31文字以内かチェック
    if ($SheetName.Length -gt 31) {
        Write-Warning '文字数が31文字以内ではない'
        return $false
    }

    # 禁止された文字を含むかチェック
    if ($SheetName -match "[:\\\/\?\*\[\]]") {
        Write-Warning 'コロン(:)または円記号(\)、スラッシュ(/)、疑問符(?)、アスタリスク(*)、左右の角かっこ([])が含まれている'
        return $false
    }

    return $true
}

# 文字列の拡張子をチェック
function Test-FileExtension {
    param (
        [Parameter(Mandatory=$true)][System.String]$FullFilename,
        [Parameter(Mandatory=$true)][System.String[]]$Extensions
    )

    # 文字列の存在チェック
    #   空文字・空白・Nullチェック
    if ([System.String]::IsNullOrWhiteSpace($FullFilename.Trim())) {
        Write-Error 'チェック対象の文字列に値が設定されていません。'
        return $false
    }
    #   ピリオドを含んでいるかチェック
    if ($FullFilename -notmatch '\.') {
        Write-Error 'チェック対象の文字列にピリオドが含まれていません。'
        return $false
    }
    #   ピリオドの位置が先頭、または末尾でないことをチェック
    $dotIndex = $FullFilename.LastIndexOf('.')
    if (($dotIndex -eq 0) -or
        ($dotIndex -eq $FullFilename.Length - 1)) {
        Write-Error 'チェック対象の文字列が正しいファイル名の表記ではありません。'
        return $false
    }

    # 配列内のチェック
    foreach ($item in $Extensions) {
        # Nullまたは空文字、空白のチェック
        if ([System.String]::IsNullOrWhiteSpace($item.Trim())) {
            Write-Warning '拡張子の配列内で値が設定されていないデータがあります。'
            return $false
        }
        # 先頭文字がピリオドから始まるかチェック
        if ($item -notmatch '^\.') {
            Write-Warning '拡張子の配列内に先頭文字がピリオドで始まっていないデータがあります。'
            return $false
        }
    }

    # 拡張子のチェック
    #   拡張子を取得
    [System.String]$fileExtension = $FullFilename -replace '.*(\..*)', '$1'

    #   拡張子の比較
    $isHit = $false
    foreach ($item in $Extensions) {
        # チェック対象の拡張子と比較する拡張子が合致した場合
        if ($fileExtension -eq $item) {
            $isHit = $true
            break
        }
    }

    # 判定した結果
    return $isHit
}
# シートの存在チェック
Function Test-ExcelSheetExists {
    param(
        [System.String]$Path,
        [System.String]$CheckSheet
    )

    $sheetExists = $false

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)
        $workSheets = $workBook.Sheets

        foreach ($workSheet in $workSheets) {
            if ($workSheet.Name -eq $CheckSheet) {
                $sheetExists = $true
                break
            }
        }
    }
    catch {
        Write-Error "シートの存在チェックで予期しないエラーが発生しました。[詳細: $($_.Exception.Messsage)]"
    }
    finally {
        # ワークブックまで解放
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # ワークブックの保存しないで終了
            $workBook.Close($false)
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Excelアプリ終了
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }

    return $sheetExists
}

# 指定したシート名を削除するFunction
function Remove-ExcelSheets {
    param (
        [System.String]$Path,
        [System.String[]]$RemoveSheets
    )

    # 入力チェック
    if (-not (Test-Path $Path)) {
        Write-Warning "対象パスが有効ではありません。[対象パス: $($Path)]"
        return
    }
    # 拡張子のチェック
    elseif (-not (Test-FileExtension $Path @('.xls', '.xlsx'))) {
        Write-Warning "対象パスのファイルがExcelファイルではありません。[対象パス: $($Path)]"
        return
    }
    # ファイルのロック状態をチェック
    elseif (Test-FileLocked($Path)) {
        Write-Warning "対象ファイルが開かれています。ファイルを閉じてから再試行してください。[対象ファイル: $($Path)]"
        return
    }

    # シートの削除処理
    $excelApp = $null
    $workBooks = $null
    $workBook = $null
    $workSheets = $null
    $workSheet = $null

    try {
        # COMオブジェクトを参照
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false

        # 対象ファイルを開く処理
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)
                
        # シートを参照
        $workSheets = $workBook.Worksheets

        # 引数で指定されたシート分を繰り返す
        foreach ($removeSheet in $RemoveSheets) {
            # シート数が1つの場合は、削除処理を中断
            if ($workSheets.Count -eq 1){
                Write-Warning "現在、Excel内のシート数は1つです。Excelでは最低1つのシートが必要となるため、削除処理を中断します。"
                break
            }
            # シート名の値チェック（空文字・シート名として使用可能な文字列）
            elseif (-not (Test-ExcelSheetname $removeSheet)) {
                Write-Warning "削除対象のシート名が適正な値ではありません。[削除対象シート名: $($removeSheet)]]"
                continue
            }
            # 削除シートの存在チェック
            elseif (-not (Test-ExcelSheetExists $Path $removeSheet)) {
                Write-Warning "削除対象のシート名が存在しません。[対象パス: $($Path), 削除対象のシート名: $($removeSheet)]]"
                continue
            }

            # シートの削除
            $workSheet = $workSheets.Item($removeSheet)
            $workSheet.Delete()
        }
    }

    catch {
        Write-Error "シートの削除処理で予期しないエラーが発生しました。[詳細: $($_.Exception.Message), 場所: $($_.InvocationInfo.MyCommand.Name)]"
    }

    finally {
        # ワークブックまで解放
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # 保存して終了
            $workBook.Close($true)

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Excelアプリ終了
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }
}

# 配列の種類を判定
function Get-ArrayType {
    param(
        $InputObject
    )
    
    [System.Collections.Hashtable]$arrayTypes = @{
        "OtherTypes" = -1
        "SingleArray" = 0
        "MultiLevel" = 1
        "MultiDimensional" = 2
    }

    # データがない場合
    if ($null -eq $InputObject) {
        return $arrayTypes["OtherTypes"]
    }

    # 一番外枠が配列ではない場合
    if ($InputObject -isnot [System.Array]) {
        return $arrayTypes["OtherTypes"]
    }

    # ジャグ配列（多段階配列）か判定
    $isMultiLevel = $false
    foreach ($element in $InputObject) {
        if ($element -is [System.Array]) {
            # 配列の中も配列で多段配列
            $isMultiLevel = $true
            break
        }
    }
    if ($isMultiLevel) {
        return $arrayTypes["MultiLevel"]
    }    
    
    # 多次元配列か判定
    if ($InputObject.Rank -ge 2) {
        # 2次元以上の場合
        return $arrayTypes["MultiDimensional"]
    }
    else {
        # 1次元の場合
        # 前提：冒頭の「-isnot [System.Array]」により配列であることは確認済みとなる。
        return $arrayTypes["SingleArray"]
    }
}
# 多次元配列を比較
function Test-ArrayEquality {
    param (
        [Parameter(Mandatory=$true)]$Array1,
        [Parameter(Mandatory=$true)]$Array2
    )

    # 多次元配列か判定
    $resultArrayType = (Get-ArrayType $Array1)
    if ($resultArrayType -ne 2) {
        Write-Warning "引数の「Array1」が多次元配列ではありません。[配列の判定結果: $($resultArrayType)]"
        return
    }
    $resultArrayType = (Get-ArrayType $Array2)
    if ($resultArrayType -ne 2) {
        Write-Warning "引数の「Array2」が多次元配列ではありません。[配列の判定結果: $($resultArrayType)]"
        return
    }

    # 配列の次元数を比較
    $dimensionArray1 = $Array1.Rank
    $dimensionArray2 = $Array2.Rank

    if ($dimensionArray1 -ne $dimensionArray2) {
        return $false
    }

    # 各次元毎の要素数をチェック
    for ($i = 0; $i -lt $dimensionArray1; $i++) {
        if ($Array1.GetLength($i) -ne $Array2.GetLength($i)) {
            return $false
        }
    }

    # 要素数が一致
    return $true
}

# 最大ページ数を計算
function Get-MaxPage {
    param(
        [System.Int32]$FirstPageCount = $TEMPLATEROWS1,
        [System.Int32]$OtherPageCount = $TEMPLATEROWS2,
        [System.Int32]$DataCount
    )

    if ($DataCount -le $FirstPageCount) {
        $maxPage = 1
    } else {
        $maxPage = [System.Math]::Ceiling(($dataCount - $firstPageCount) / $otherPageCount)
        $maxPage += 1
    }
    return $maxPage
}

# シート内の定数をキーに値を設定する
function Set-PackagelistValues {
    param(
        [PrintListConfig]$Config
    )

    # 項目名を取得
    $headerColumns = @($Config.HeaderConstants[0].psobject.properties | ForEach-Object { $_.Name })
    $mainColumns = @($Config.MainConstants[0].psobject.properties | ForEach-Object { $_.Name })


    # 置換処理
    $excelApp = $null
    $workBooks = $null
    $workBook = $null
    $workSheets = $null
    $workSheet = $null

    # 現在のページ
    $currentPage = 1
    $maxPage = (Get-MaxPage -DataCount $Config.MainValues.Count)
    $currentRow = 1

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Config.Path)
        $workSheets = $workBook.Sheets

        # 1ページ目
        #   シート準備
        $currentPage = 1
        $currentSheet = "$($currentPage) Page"
        $workSheet = $workSheets.Item($TEMPLATESHEET1)
        # シートのコピー
        $workSheet.Copy($workSheet)
        # コピーしたシート名を変更（コピー後にアクティブシートが対象シートになることが前提）
        $excelApp.ActiveSheet.Name = $currentSheet

        #   ヘッダー情報を反映
        $workSheet = $workSheets.Item($currentSheet)
        $range = $workSheet.Range($HEADERRANGE)
                
        for ($i = 0; $i -lt $Config.HeaderConstants.Count; $i++) {
            for ($j = 0; $j -lt $headerColumns.Count; $j++) {
                $range.Replace(
                    $($Config.HeaderConstants[$i]."$($headerColumns[$j])"),
                    $($Config.HeaderValues[$i]."$($headerColumns[$j])"),
                    # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                    [XlLookAt]::xlWhole
                ) | Out-Null
            }
        }

        #   メイン情報を反映
        $startRow = $currentRow - 1
        [System.Int32]$rangeRows = $TEMPLATEROWS1
        $rangeMainValues = @($Config.MainValues | Select-Object -Skip $startRow -First $rangeRows)

        $workSheet = $workSheets.Item($currentSheet)
        $range = $workSheet.Range($MAINRANGE)

        for ($i = 0; $i -lt $rangeMainValues.Count; $i++) {
            # 定数の連番用
            $num = "{0:D2}" -f $($i + 1)
            for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                # 定数に連番を追加
                $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                $range.Replace(
                    $mainConstantsWithNum,
                    $($rangeMainValues[$i]."$($mainColumns[$j])"),
                    # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                    [XlLookAt]::xlWhole
                ) | Out-Null
            }
        }

        #   空欄を作成
        for ($i = 0; $i -lt $TEMPLATEROWS1; $i++) {
            # 定数の連番用
            $num = "{0:D2}" -f $($i + 1)
            for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                # 定数に連番を追加
                $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                $range.Replace(
                    $mainConstantsWithNum,
                    '',
                    # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                    [XlLookAt]::xlWhole
                ) | Out-Null
            }
        }

        # ページ数を進める
        $currentRow = $TEMPLATEROWS1 + 1

        # 2ページ目以降
        for ($currentPage = 2; $currentPage -le $maxPage; $currentPage++) {
            #   シート準備
            $currentSheet = "$($currentPage) Page"
            $workSheet = $workSheets.Item($TEMPLATESHEET2)
            # シートのコピー
            $workSheet.Copy($workSheet)
            # コピーしたシート名を変更（コピー後にアクティブシートが対象シートになることが前提）
            $excelApp.ActiveSheet.Name = $currentSheet

            #   ヘッダー情報を反映
            $workSheet = $workSheets.Item($currentSheet)
            $range = $workSheet.Range($HEADERRANGE)

            for ($i = 0; $i -lt $Config.HeaderConstants.Count; $i++) {
                for ($j = 0; $j -lt $headerColumns.Count; $j++) {
                    $range.Replace(
                        $($Config.HeaderConstants[$i]."$($headerColumns[$j])"),
                        $($Config.HeaderValues[$i]."$($headerColumns[$j])"),
                        # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                        [XlLookAt]::xlWhole
                    ) | Out-Null
                }
            }

            #   メイン情報を反映
            $startRow = $currentRow - 1
            [System.Int32]$rangeRows = $TEMPLATEROWS2
            $rangeMainValues = @($Config.MainValues | Select-Object -Skip $startRow -First $rangeRows)

            $workSheet = $workSheets.Item($currentSheet)
            $range = $workSheet.Range($MAINRANGE)

            for ($i = 0; $i -lt $rangeMainValues.Count; $i++) {
                # 定数の連番用
                $num = "{0:D2}" -f $($i + 1)
                for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                    # 定数に連番を追加
                    $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                    $range.Replace(
                        $mainConstantsWithNum,
                        $($rangeMainValues[$i]."$($mainColumns[$j])"),
                        # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                        [XlLookAt]::xlWhole
                    ) | Out-Null
                }
            }

            #   空欄を作成
            for ($i = 0; $i -lt $TEMPLATEROWS2; $i++) {
                # 定数の連番用
                $num = "{0:D2}" -f $($i + 1)
                for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                    # 定数に連番を追加
                    $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                    $range.Replace(
                        $mainConstantsWithNum,
                        '',
                        # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                        [XlLookAt]::xlWhole
                    ) | Out-Null
                }
            }

            # ページ数を進める
            $currentRow += $TEMPLATEROWS2
        }
    }
    catch {
        # エラー時の処理
        Write-Error "予期しないエラーが発生しました。[詳細: $($_.Exception.ToString())]"
    }
    finally {
        # ワークブックまで解放
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # ワークブックを保存して終了
            $workBook.Close($true)
                    
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Excelアプリ終了
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }
}
function Export-ExcelDocumentAsPDF {
    param(
        [parameter(Mandatory=$true)][string]$Path,
        [parameter(Mandatory=$true)][string]$OutputPath
    )

    $excelApp = $null
    $workBooks = $null
    $workBook = $null

    try {
        # Excelファイルを開く処理
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)

        # PDFファイルで出力する処理
        $workBook.ExportAsFixedFormat(
            # [Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF,
            [XlFixedFormatType]::xlTypePDF,
            $OutputPath
        )
    }
    catch {
        Write-Error "ExcelファイルをPDFファイルとして出力する処理でエラーが発生しました。"
        Write-Error "エラー情報 [詳細: $($_.Exception.Message), 場所: $($_.InvocationInfo.MyCommand.Name)]"
    }
    finally {
        if ($null -ne $workBook) {
            # ワークブックを保存しないで終了
            $workBook.Close($false)
                    
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Excelアプリ終了
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }
}

# PSCustomObjectの判定
function Test-IsPSCustomObject {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$Argument
    )

    foreach ($arg in $Argument) {
        if (-not ($arg -is [System.Management.Automation.PSCustomObject])) {
            return $false
        }
    }
    return $true
}

# PSCustomObjectの比較
Function Test-PSCustomObjectEquality {
    param (
        [Parameter(Mandatory=$true)][System.Object[]]$Object1,
        [Parameter(Mandatory=$true)][System.Object[]]$Object2
    )

    # データ存在チェック
    if (($Object1.Count -eq 0) -or ($Object2.Count -eq 0)) {
        return $false
    }

    # オブジェクト内がPSCustomObjectであるか判定
    if (-not (Test-IsPSCustomObject $Object1)) {
        return $false
    }
    elseif (-not (Test-IsPSCustomObject $Object2)) {
        return $false
    }

    # 項目名を比較
    $object1ColumnData = $Object1[0].psobject.properties | ForEach-Object { $_.Name }
    $object2ColumnData = $Object2[0].psobject.properties | ForEach-Object { $_.Name }
    $compareResult = (Compare-Object $object1ColumnData $object2ColumnData -SyncWindow 0)
    if (($null -ne $compareResult) -and ($compareResult.Count -ne 0)) {
        return $false
    }

    # 比較した結果2つのオブジェクトが一致
    return $true
}
