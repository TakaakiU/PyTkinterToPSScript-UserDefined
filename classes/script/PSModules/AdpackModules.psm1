# 必要なアセンブリを読み込む
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.IO.Compression.FileSystem

# 指定されたXMLファイルの格納先フォルダーを作成
function New-XmlFolder {
    param (
        [System.String]$FilePath
    )

    [System.Int32]$statusCode = 0

    # XMLファイルの格納先フォルダーを取得
    $xmlDirectory = Split-Path -Path $FilePath
    # 格納先フォルダーがない場合にフォルダー作成
    if (-Not (Test-Path -Path $xmlDirectory)) {
        try {
            New-Item -ItemType Directory -Path $xmlDirectory | Out-Null
        }
        catch {
            Write-Error "エラーが発生しました: $($_.Exception.Message)"
            $statusCode = -7101
        }
    }
}

# ファイルのSHA256ハッシュを計算する関数
function Get-FileHashEX {
    param (
        [System.String]$FilePath,
        [ValidateSet("BASE64", "HEX")]
        [System.String]$HashFormat = "BASE64", # "BASE64" または "HEX" を指定可能
        [ValidateSet("SHA256", "SHA384", "SHA512", "SHA1", "MD5")]
        [System.String]$HashAlgorithm = "SHA256" # 使用するハッシュアルゴリズム
    )

    # Get-FileHashを使用してハッシュ値を取得
    $hashObject = Get-FileHash -Path $FilePath -Algorithm $HashAlgorithm

    # フォーマットに応じてハッシュ値を変換して返す
    switch ($HashFormat) {
        "BASE64" {
            # HEX形式をBASE64形式に変換
            $hashBytes = [System.Convert]::FromHexString($hashObject.Hash)
            return [Convert]::ToBase64String($hashBytes)
        }
        "HEX" {
            # そのままHEX形式を返す
            return $hashObject.Hash
        }
        default {
            throw "Unsupported format: $HashFormat. Use 'BASE64' or 'HEX'."
        }
    }
}

# XML文書を生成する関数
# Index.xmlを生成する関数
function New-IndexXml {
    param (
        [System.String]$FolderPath,  # ハッシュ値を取得する対象フォルダー
        [System.String]$Destination  # Index.xml の保存先
    )

    [System.Int32]$statusCode = 0

    # 実行時間を取得
    $currentDate = (Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz")

    # ユーザー名とホスト名を取得
    $envUser = $env:USERNAME
    $envHost = $env:COMPUTERNAME

    try {
        # 新しいXMLドキュメントを作成
        $xmlDoc = New-Object System.Xml.XmlDocument
        $root = $xmlDoc.CreateElement("Index")
        $xmlDoc.AppendChild($root)

        # 各要素を作成してXMLに追加
        $title = $xmlDoc.CreateElement("Title")
        $title.InnerText = $FolderPath
        $root.AppendChild($title)

        $date = $xmlDoc.CreateElement("Date")
        $date.InnerText = $currentDate
        $root.AppendChild($date)

        $user = $xmlDoc.CreateElement("User")
        $user.InnerText = $envUser
        $root.AppendChild($user)

        $hostname = $xmlDoc.CreateElement("Host")
        $hostname.InnerText = $envHost
        $root.AppendChild($hostname)

        # XMLを保存
        $xmlDoc.Save($Destination)
    }
    catch {
        Write-Error "META-INF/Index.xmlを作成時にエラーが発生しました。[$($_.Exception.Message)]"
        $statusCode = -7102
    }

    return $statusCode
}

function New-ManifestXml {
    param (
        [System.String]$FolderPath,
        [System.String]$Destination,
        [ValidateSet("BASE64", "HEX")]
        [System.String]$HashFormat = "BASE64", # "BASE64" または "HEX" を指定可能
        [ValidateSet("SHA256", "SHA384", "SHA512", "SHA1")]
        [System.String]$HashAlgorithm = "SHA256" # 使用するハッシュアルゴリズム
    )

    [System.Int32]$statusCode = 0

    try {
        # 新しいXMLドキュメントを作成
        $xmlDoc = New-Object System.Xml.XmlDocument
        
        # 名前空間を含むルート要素を作成
        $namespace = "http://www.w3.org/2000/09/xmldsig#"
        $root = $xmlDoc.CreateElement("Manifest", $namespace)
        $root.SetAttribute("Id", "tupk-Manifest-0")
        $xmlDoc.AppendChild($root)
        
        # 各ファイルのハッシュ値を計算してXMLに追加
        # $files = (Get-ChildItem -Recurse -Path $FolderPath | Where-Object { -Not $_.PSIsContainer })
        $files = Get-ChildItem -Recurse -Path $FolderPath | Where-Object {
            -Not $_.PSIsContainer -and $_.FullName -ne $Destination
        }
        foreach ($file in $files) {
            $relativePath = $file.FullName.Substring($FolderPath.Length + 1)
            $hash = Get-FileHashEX -FilePath $file.FullName

            $reference = $xmlDoc.CreateElement("Reference", $namespace)
            $reference.SetAttribute("URI", $relativePath)

            $digestMethod = $xmlDoc.CreateElement("DigestMethod", $namespace)
            # ハッシュアルゴリズムごとに変更
            switch ($HashAlgorithm) {
                "SHA256" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256")
                }
                "SHA384" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2001/04/xmldsig-more#sha384")
                }
                "SHA512" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha512")
                }
                "SHA1" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2000/09/xmldsig#sha1")
                }
            }
            $reference.AppendChild($digestMethod)

            $digestValue = $xmlDoc.CreateElement("DigestValue", $namespace)
            $digestValue.InnerText = $hash
            $reference.AppendChild($digestValue)

            $root.AppendChild($reference)
        }

        # XMLを保存
        $xmlDoc.Save($Destination)
    }
    catch {
        Write-Error "META-INF/Manifest.xmlを作成時にエラーが発生しました。[$($_.Exception.Message)]"
        $statusCode = -7103
    }

    return $statusCode
}

function Test-DirectoryStructureMatch {
    param (
        # 実データを含むルートフォルダーのパス
        [Parameter(Mandatory = $true)]
        [string]$ExtractedFolder,

        # Manifest.xml の格納場所（ルートからの相対パス）
        [string]$ManifestRelativePath = "META-INF/Manifest.xml"
    )

    # Manifest ファイルのフルパスを取得
    $manifestFileFullPath = Join-Path -Path $ExtractedFolder -ChildPath $ManifestRelativePath
    
    # XML ファイルの読み込み
    try {
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($manifestFileFullPath)
    }
    catch {
        Write-Debug "XMLの読み込みに失敗しました: $_"
        return $false
    }

    # 名前空間マネージャの生成と設定
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $namespaceManager.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")

    # XML 内の <ds:Reference> 要素から、ファイルパス（URI）を抽出
    $references = $xmlDoc.SelectNodes("//ds:Reference", $namespaceManager)
    $manifestPaths = @()
    foreach ($ref in $references) {
        $uri = $ref.GetAttribute("URI")
        # パス区切りの形式を統一（Windows の "\" を "/" に変換）
        $manifestPaths += $uri.Replace("\", "/")
    }
    # Manifest 内のすべての META-INF 配下の項目を除外
    $manifestPaths = $manifestPaths | Where-Object { $_ -notmatch "^META-INF" }

    # 実データ側のファイル一覧を取得（※ META-INF配下は除外）
    $actualFiles = Get-ChildItem -Path $ExtractedFolder -Recurse -File | 
                   Where-Object { $_.FullName -notmatch [regex]::Escape("META-INF") }
    # ルート部分 ($ExtractedFolder) を除去した相対パスへ変換（"/" 区切り）
    $actualPaths = $actualFiles | ForEach-Object {
        ($_.FullName.Substring($ExtractedFolder.Length)).TrimStart('\','/').Replace("\", "/")
    }

    # XML に「ある」が、実データ側には「ない」ファイルを取得
    $missingFiles = $manifestPaths | Where-Object { $_ -notin $actualPaths }

    # 実データ側に「ある」が、XML に「ない」ファイルを取得
    $unexpectedFiles = $actualPaths | Where-Object { $_ -notin $manifestPaths }

    # 差分詳細の出力（Write-Debug を使用）
    if ($missingFiles.Count -eq 0 -and $unexpectedFiles.Count -eq 0) {
        Write-Debug "ディレクトリ構成は一致しています。"
        return $true
    }
    else {
        if ($missingFiles.Count -gt 0) {
            Write-Debug "Manifest.xml に記載がありますが、実データ側に存在しないファイル:"
            foreach ($file in $missingFiles) {
                Write-Debug "    $file"
            }
        }
        if ($unexpectedFiles.Count -gt 0) {
            Write-Debug "実データ側に存在しますが、Manifest.xml に記載されていないファイル:"
            foreach ($file in $unexpectedFiles) {
                Write-Debug "    $file"
            }
        }
        return $false
    }
}

function Test-HashValues {
    param (
        [System.String]$HashAlgorithm,
        [System.String]$ExtractedFolder
    )

    try {
        # XMLファイル読み込み
        $manifestFilePath = "$ExtractedFolder/META-INF/Manifest.xml"

        # XML署名ファイルを読み込む
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($manifestFilePath)

        # 名前空間マネージャの作成と設定
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager $xmlDoc.NameTable
        $namespaceManager.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")

        # 署名内のハッシュ情報を収集
        $references = $xmlDoc.SelectNodes("//ds:Reference", $namespaceManager)

        $isVerified = $true

        foreach ($ref in $references) {
            $uri = $ref.GetAttribute("URI")
            $expectedHash = $ref.SelectSingleNode("ds:DigestValue", $namespaceManager).InnerText
            $digestAlgorithm = $ref.SelectSingleNode("ds:DigestMethod", $namespaceManager).GetAttribute("Algorithm")

            # XMLファイルで明記されているハッシュアルゴリズム
            switch ($digestAlgorithm) {
                "http://www.w3.org/2001/04/xmlenc#sha256" {
                    Write-Debug "SHA256 アルゴリズムが使用されています。"
                }
                "http://www.w3.org/2001/04/xmlenc#sha512" {
                    Write-Debug "SHA512 アルゴリズムが使用されています。"
                }
                "http://www.w3.org/2001/04/xmldsig-more#sha384" {
                    Write-Debug "SHA384 アルゴリズムが使用されています。"
                }
                "http://www.w3.org/2000/09/xmldsig#sha1" {
                    Write-Debug "SHA1 アルゴリズムが使用されています。"
                }
            }

            # ファイルパスを構築
            $filePath = (Join-Path -Path $ExtractedFolder -ChildPath $uri)

            if (-Not (Test-Path -Path $filePath)) {
                Write-Host "ファイルが見つかりません: $uri" -ForegroundColor Red
                $isVerified = $false
                continue
            }

            # ハッシュを計算して比較
            $computedHash = (Get-FileHashEX -FilePath $filePath -Algorithm $HashAlgorithm)

            if ($computedHash -ne $expectedHash) {
                Write-Host "ハッシュが一致しません: $uri" -ForegroundColor Red
                Write-Host "期待されるハッシュ: $expectedHash" -ForegroundColor Yellow
                Write-Host "計算されたハッシュ: $computedHash" -ForegroundColor Yellow
                $isVerified = $false
            }
        }

        if ($isVerified) {
            Write-Host "すべてのファイルが署名と一致しています。" -ForegroundColor Green
        } else {
            Write-Host "一致しないファイルが存在します。" -ForegroundColor Red
            $statusCode = -7104
        }
    }
    catch {
        Write-Error "エラーが発生しました: $_"
    }

    return $statusCode
}

function Compress-Package_Adpack {
    param (
        [System.String]$ExePath,
        [System.String]$HashAlgorithm,
        [System.String]$FolderPath,
        [System.String]$ZipFilePath
    )

    [System.Int32]$statusCode = 0

    # # PowerShellバージョンチェック
    # $statusCode = Test-PowerShellVersion7OrLater

    # パッケージ生成実行
    if ($statusCode -eq 0) {
        try {
	        (& $ExePath -pack -hash $HashAlgorithm -in $FolderPath -out $ZipFilePath -force | Out-Null)
            # if (-not ($?)) {
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7105
            }
        }
        catch {
	        Write-Error "パッケージ生成中にエラーが発生しました。[対象：$($FolderPath)]"
	        $statusCode = -7106
        }
    }

    # 圧縮後のZIPファイルと対象フォルダーを比較
    if ($statusCode -eq 0) {
        Write-Host "ステップ3 開始：圧縮前後を比較"
        Write-Host ""
        $statusCode = (Compare-ZipAndFolderContent -FolderPath $FolderPath -ZipFilePath $ZipFilePath)
    }

    return $statusCode
}

function Compress-Package_UserDefined {
    param (
        [System.String]$HashAlgorithm,
        [System.String]$FolderPath,
        [System.String]$ZipFilePath
    )

    # META-INFフォルダーとXMLファイルのパスを指定
    $xmlDirectory = "$FolderPath/META-INF"
    $indexFilePath = "$xmlDirectory/Index.xml"
    $manifestFilePath = "$xmlDirectory/Manifest.xml"

    [System.Int32]$statusCode = 0

    # # PowerShellバージョンチェック
    # $statusCode = Test-PowerShellVersion7OrLater

    Write-Host "ステップ1 開始：XMLファイルの作成"
    Write-Host ""
    # META/INFフォルダーを作成
    if ($statusCode -eq 0) {
        $statusCode = (New-DirectoryIfNotExists $xmlDirectory -ForceRecreate)
    }
    if ($statusCode -eq 0) {
        # XMLファイルを生成
        $statusCode = (New-IndexXml -FolderPath $FolderPath -Destination $indexFilePath | Out-Null)
    }
    if ($statusCode -eq 0) {
        $statusCode = (New-ManifestXml -FolderPath $FolderPath -Destination $manifestFilePath -HashAlgorithm $HashAlgorithm | Out-Null)
    }

    # フォルダーをZIPファイルに圧縮
    if ($statusCode -eq 0) {
        Write-Host "ステップ2 開始：ZIPファイルに圧縮"
        Write-Host ""
        $statusCode = (Compress-FolderToZip -FolderPath $FolderPath -Destination $ZipFilePath)
    }

    # 圧縮後のZIPファイルと対象フォルダーを比較
    if ($statusCode -eq 0) {
        Write-Host "ステップ3 開始：圧縮前後を比較"
        Write-Host ""
        $statusCode = (Compare-ZipAndFolderContent -FolderPath $FolderPath -ZipFilePath $ZipFilePath)
        $resultData | ForEach-Object { Write-Host $_ }
    }

    return $statusCode
}

# ZIPファイルを解凍する関数
function Expand-Package_Adpack {
    param (
        [Switch]$Check,
        [Switch]$NoCheck,
        [System.String]$ExePath,
        [System.String]$HashAlgorithm,
        [System.String]$ZipFilePath,
        [System.String]$FolderPath
    )

    [System.Int32]$statusCode = 0

    # チェック処理
    if ($statusCode -eq 0) {
        if (-Not $NoCheck) {
            (& $ExePath -unpack -in $ZipFilePath -check | Out-Null)
            # if (-not ($?)) {
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7107
            }
        }
        else {
            Write-Host "オプション指定によりXML署名チェックをスキップしました。"
        }
    }

    if ($statusCode -eq 0) {
        Write-Host "Unpack Command[$ExePath -unpack -in $ZipFilePath -out $FolderPath -force -nocheck | Out-Null]"
        Write-Host ""
        # オプションがチェック以外の時は出力データを準備
        if (-Not ($Check)) {
            (& $ExePath -unpack -in $ZipFilePath -out $FolderPath -force -nocheck | Out-Null)
            # if (-not ($?)) {
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7108
            }
        }
        # チェックのみの場合は一時フォルダーを削除
        else {
            (& $ExePath -unpack -in $ZipFilePath -check | Out-Null)
            # if (-not ($?)) {
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7109
            }
        }
    }
    
    return $statusCode
}

function Expand-Package_UserDefined {
    param (
        [Switch]$Check,
        [Switch]$NoCheck,
        [System.String]$HashAlgorithm,
        [System.String]$ZipFilePath,
        [System.String]$FolderPath
    )

    [System.Int32]$statusCode = 0

    # # PowerShellバージョンチェック
    # $statusCode = Test-PowerShellVersion7OrLater

    # 一時解凍フォルダーを作成
    if ($statusCode -eq 0) {
        $zipDirectory = (Split-Path -Path $ZipFilePath)
        $TempExtractFolder = Join-Path -Path ($zipDirectory) -ChildPath (".TempExtract_" + [System.Guid]::NewGuid())
        $statusCode = (New-DirectoryIfNotExists -FolderPath $TempExtractFolder)
    }

    # 解凍
    if ($statusCode -eq 0) {
        Write-Host "ステップ1 開始：パッケージデータの解凍"
        Write-Host ""
        $statusCode = (Expand-ZipToTempFolder -ZipFilePath $ZipFilePath -TempFolderPath $TempExtractFolder)
    }

    # META-INF および manifest.xml の存在確認
    if ($statusCode -eq 0) {
        if (-not (Test-Path "$TempExtractFolder/META-INF" -PathType Container) -or
            -not (Test-Path "$TempExtractFolder/META-INF/Manifest.xml")) {
                Write-Error "META-INFフォルダー もしくは META-INF/Manifest.xml がありません。"
                $statusCode = -7110
        }
    }

    # META-INF/manifest.xml と 実データ のディレクトリ構造を比較
    if ($statusCode -eq 0) {
        Write-Host "ステップ2-1 開始：manifest.xml と ディレクトリ構造を比較"
        Write-Host ""
        if (-Not $NoCheck) {
            if (-not (Test-DirectoryStructureMatch $TempExtractFolder)) {
                $statusCode = -7111
            }
        }
        else {
            Write-Host "オプション指定によりXML署名チェックをスキップしました。"
        }
    }

    # META-INF/manifest.xml と 実データ の比較
    if ($statusCode -eq 0) {
        Write-Host "ステップ2-2 開始：manifest.xml と 実データ のハッシュ値を比較"
        Write-Host ""
        if (-Not $NoCheck) {
            $statusCode = (Test-HashValues -HashAlgorithm $HashAlgorithm -ExtractedFolder $TempExtractFolder -ZipFilePath $ZipFilePath)
        }
        else {
            Write-Host "オプション指定によりXML署名チェックをスキップしました。"
        }
    }

    if ($statusCode -eq 0) {
        Write-Host "ステップ3 開始：事後処理"
        Write-Host ""
        # オプションがチェック以外の時は出力データを準備
        if (-Not ($Check)) {
            # $movePath = [System.IO.Path]::ChangeExtension($ZipFilePath, $null)
            $rootPath = [System.IO.Path]::GetDirectoryName($ZipFilePath)
            $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($ZipFilePath)
            $movePath = [System.IO.Path]::Combine($rootPath, $fileNameWithoutExt)

            $statusCode = (Remove-Folder -FolderPath $movePath)

            if ($statusCode -eq 0) {
                try {
                    #Move-Item -Path $TempExtractFolder -Destination $FolderPath
                    Move-Item -Path $TempExtractFolder -Destination $movePath

                }
                catch {
                    Write-Error "一時フォルダーの名前変更時にエラーが発生しました[$($_.Exception.Message)]"
                    $statusCode = -7112
                }
            }
        }
        # チェックのみの場合は一時フォルダーを削除
        else {
            $statusCode = (Remove-Folder -FolderPath $TempExtractFolder)
        }
    }
    
    return $statusCode
}
