# PowerShell 7以上を判定（大容量ファイルをZIP圧縮するために必須）
#   PowerShell 5 では.NET Framework 、7 では .NET Core で動作。
#   .NET Coreでパフォーマンスやメモリ効率が強化している事により PowerShell 7 が必須となる。
function Test-PowerShellVersion7OrLater {
    $statusCode = 0
    # PowerShellのバージョンを取得
    $currentVersion = $PSVersionTable.PSVersion
    $requiredMajorVersion = 7

    # バージョン確認
    if ($currentVersion.Major -lt $requiredMajorVersion) {
        Write-Host "このスクリプトはPowerShell 7以降で実行する必要があります。" -ForegroundColor Red
        Write-Host "現在のバージョン: $currentVersion" -ForegroundColor Yellow
        # 処理を中断
        $statusCode = -6001
    }

    return $statusCode
}

# ファイルもしくはフォルダーをチェック
Function Test-PathType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Source", "Target")]
        [string]$PathRole, # Pathの役割: 入力(Source) or 出力(Target)

        [Parameter(Mandatory = $true)]
        [ValidateSet("Container", "Leaf")]
        [string]$PathType, # 期待されるPathタイプ: フォルダー(Container) or ファイル(Leaf)
        
        [Parameter(Mandatory = $false)]
        [string[]]$Extensions = $null # デフォルト: 拡張子チェックなし
    )

    $isMatch = $false
    
    # 引数チェック
    if (($PathType -eq "Container") -And $Extensions) {
        Write-Error "実行不可な引数の組み合わせ。引数を見直してください。"
        return $isMatch
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $localPathtype = $PathType
    # "Target"の場合は親ディレクトリを取得
    if ($PathRole -eq "Target") {
        $parentDir = Split-Path -Path $fullPath
        if (-Not (Test-Path -Path $parentDir)) {
            Write-Error "指定されたターゲットパスの親ディレクトリが存在しません: $parentDir"
            return $isMatch
        }
        # 以降、親ディレクトリーをチェック対象に変更
        $fullPath = $parentDir
        $localPathtype = "Container"
    }

    # パスの存在チェック
    if (-Not (Test-Path -Path $fullPath)) {
        Write-Error "指定されたパスは存在しません: $fullPath"
        return $isMatch
    }

    # パスタイプのチェック
    switch ($localPathtype) {
        "Leaf" {
            if (-Not (Test-Path -Path $fullPath -PathType Leaf)) {
                Write-Error "ファイルを期待しましたが、フォルダーが指定されました: $fullPath"
                return $isMatch
            }
        }
        "Container" {
            if (-Not (Test-Path -Path $fullPath -PathType Container)) {
                Write-Error "フォルダーを期待しましたが、ファイルが指定されました: $fullPath"
                return $isMatch
            }
        }
    }

    # ファイル拡張子のチェック（ファイルの場合）
    if ($Extensions -and $PathType -eq "Leaf") {
        $extension = [System.IO.Path]::GetExtension($Path).TrimStart(".").ToLower()
        if (-Not ($Extensions -contains $extension)) {
            Write-Error "指定されたファイル拡張子 '$extension' は許可されていません: $Extensions"
            return $isMatch
        }
    }

    # すべてのバリデーションを通過
    $isMatch = $true
    return $isMatch
}

# 一意のパスを取得
function Get-UniqueFilePath {
    param (
        [System.String]$Path,
        [System.Int32]$MaxAttempts = 0
    )

    # ファイル名と拡張子を分離（フォルダーの場合は拡張子なし）
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $Extension = [System.IO.Path]::GetExtension($Path)
    $ParentFolder = [System.IO.Path]::GetDirectoryName($Path)

    # 存在しない場合はそのまま返す。
    if (-not (Test-Path $Path)) {
            return $Path
    }

    # 一意のパスを設定
    if ($MaxAttempts -eq 0) {
        $counter = 1
        do {
            # 新しいパス名を生成（連番を付加）
            $newPath = "$ParentFolder\$BaseName-$counter$Extension"
            $counter++
        } while (Test-Path $newPath)  # 存在しないパスを見つけるまで繰り返す
    }
    else {
        for ($counter = 1; $counter -le $MaxAttempts; $counter++) {
            $newPath = "$ParentFolder\$BaseName-$counter$Extension"

            if (-not (Test-Path $newPath)) {
                break
            }
        }

        if ($counter -gt $MaxAttempts) {
            throw "一意のパスを取得しましたが、試行回数「$MaxAttempts」を超過したため、処理を中断します。"
        }
    }

    return $newPath
}

# フォルダーの削除
function Remove-Folder {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FolderPath
    )

    [System.Int32]$statusCode = 0

    if (Test-Path -Path $FolderPath) {
        try {
            Remove-Item -Path $FolderPath -Recurse -Force | Out-Null
        }
        catch {
            Write-Error "エラーが発生しました: $($_.Exception.Message)"
            $statusCode = -6101
        }
    }
    else {
        Write-Host "削除対象のフォルダーが存在しなかったので処理をスキップします：[$FolderPath]"
    }

    return $statusCode
}

# ファイルの削除
function Remove-File {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FilePath
    )

    [System.Int32]$statusCode = 0

    if (Test-Path -Path $FilePath) {
        try {
            Remove-Item -Path $FilePath -Force | Out-Null
        }
        catch {
            Write-Error "エラーが発生しました: $($_.Exception.Message)"
            $statusCode = -6102
        }
    }
    else {
        Write-Host "削除対象のファイルが存在しなかったので処理をスキップします：[$FilePath]"
    }

    return $statusCode
}

# フォルダーをZIPファイルに圧縮する関数
function Compress-FolderToZip {
    param (
        [System.String]$FolderPath,
        [System.String]$Destination
    )
    [System.Int32]$statusCode = 0

    # すでにZIPファイルがある場合は削除
    if (Test-Path $Destination) {
        Remove-Item -Path $Destination -Force
    }    
    # 圧縮
    try {
        ([System.IO.Compression.ZipFile]::CreateFromDirectory($FolderPath, $Destination)) | Out-Null
        # Compress-Archive -Path $FolderPath -DestinationPath $Destination | Out-Null

        # # 7Zip4Powershell
        # Compress-7Zip -Path $FolderPath -OutputPath $Destination
    }
    catch {
        Write-Error "エラーが発生しました: $($_.Exception.Message)"
        $statusCode = -6201
    }

    return $statusCode
}

# ZIPファイルを解凍する関数
function Expand-ZipToTempFolder {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$ZipFilePath,
        [Parameter(Mandatory = $true)]
        [System.String]$TempFolderPath
    )
    [System.Int32]$statusCode = 0

    try {
        # ZIPファイルを解凍
        # Expand-Archive -Path $ZipFilePath -DestinationPath $TempFolderPath | Out-Null
        $encodingSjis = [Text.Encoding]::GetEncoding("shift_jis")
        ([System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFilePath, $TempFolderPath, $encodingSjis)) | Out-Null

        # # 7Zip4Powershell
        # Expand-7Zip -ArchiveFileName $ZipFilePath -TargetPath $TempFolderPath
    }
    catch {
        Write-Error "エラーが発生しました: $($_.Exception.Message)"
        $statusCode = -6301
    }

    return $statusCode
}

# フォルダーを作成（Switchパラメーター「-ForceRecreate」の指定で削除と作成を行う）
function New-DirectoryIfNotExists {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FolderPath,
        [Switch]$ForceRecreate
    )
    [System.Int32]$statusCode = 0

    if (-Not (Test-Path -Path $FolderPath)) {
        try {
            New-Item -ItemType Directory -Path $FolderPath | Out-Null
            Write-Debug "フォルダーを作成しました: $FolderPath"
        }
        catch {
            Write-Error "エラーが発生しました: $($_.Exception.Message)"
            $statusCode = -6401
        }
    } else {
        if ($ForceRecreate) {
            try {
                Remove-Folder -FolderPath $FolderPath | Out-Null
                New-Item -ItemType Directory -Path $FolderPath | Out-Null
                Write-Debug "フォルダーを削除して再作成しました: $FolderPath"
            }
            catch {
                Write-Error "エラーが発生しました: $($_.Exception.Message)"
                $statusCode = -6402
            }
        } else {
            Write-Debug "フォルダーは既に存在します: $FolderPath"
        }
    }

    return $statusCode
}

# ZIPファイルと対象フォルダーを比較
function Compare-ZipAndFolderContent {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FolderPath,
    
        [Parameter(Mandatory = $true)]
        [System.String]$ZipFilePath
    )

    [System.Int32]$statusCode = 0

    # 解凍用の一時保管場所としてフォルダー作成
    $zipDirectory = (Split-Path -Path $ZipFilePath) -replace('\\', '/')
    $TempExtractFolder = $zipDirectory + "/.TempExtract_" + [System.Guid]::NewGuid()
    # 存在する場合は作成する前に削除
    $statusCode = (New-DirectoryIfNotExists $TempExtractFolder -ForceRecreate)

    # 一時保管用の解凍用フォルダーを作成
    $statusCode = (Expand-ZipToTempFolder $ZipFilePath $TempExtractFolder)
    
    # 比較
    if ($statusCode -eq 0) {
        # 入力データのディレクトリ構造を取得（システムファイルを含める）
        $SourceItems = Get-ChildItem -Path $FolderPath -Recurse -Force | ForEach-Object {$_.FullName -replace "\\", "/"}
        $SourceItems_DelString = $FolderPath + "/"
        $SourceItems = $SourceItems -replace $SourceItems_DelString, ""

        # 出力データのディレクトリ構造を取得
        $ExtractedItems = Get-ChildItem -Path $TempExtractFolder -Recurse -Force | ForEach-Object {$_.FullName -replace "\\", "/"}
        $ExtractedItems_DelString = $TempExtractFolder + "/"
        $ExtractedItems = $ExtractedItems -replace $ExtractedItems_DelString, ""

        Write-Debug ($SourceItems | Out-String)
        Write-Debug ($ExtractedItems | Out-String)

        # 基点を変えて差分を取得
        $OnlyInSource = $SourceItems | Where-Object { $_ -notin $ExtractedItems }
        $OnlyInZip = $ExtractedItems | Where-Object { $_ -notin $SourceItems }

        # 比較結果を表示
        if ($OnlyInSource.Count -eq 0 -and $OnlyInZip.Count -eq 0) {
            Write-Host "フォルダーとZIPファイルの内容は一致しています。"
        }
        else {
            if ($OnlyInSource.Count -gt 0) {
                Write-Host "以下のアイテムはフォルダー[$FolderPath]にのみ存在します："
                Write-Host ($OnlyInSource | Out-String)
            }
            if ($OnlyInZip.Count -gt 0) {
                Write-Host "以下のアイテムはZIPファイル[$ZipFilePath]にのみ存在します："
                Write-Host ($OnlyInZip | Out-String)
            }
            $statusCode = -6501
        }
    }

    # 一時フォルダーを削除
    if ($statusCode -eq 0) {
        $statusCode = (Remove-Folder -FolderPath $TempExtractFolder)
    }
    
    return $statusCode
}
