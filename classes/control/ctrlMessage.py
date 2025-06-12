import tkinter as tk
from tkinter import messagebox


class ctrlMessage():
    def view_top_warning(result):
        root = tk.Tk()
        root.attributes('-topmost', True)
        root.withdraw()
        root.lift()
        root.focus_force()
        messages = ctrlMessage.get_message(result)
        messagebox.showwarning(messages[0], messages[1])
        return

    def view_top_okcancel(result):
        root = tk.Tk()
        root.attributes('-topmost', True)
        root.withdraw()
        root.lift()
        root.focus_force()
        messages = ctrlMessage.get_message(result)
        return messagebox.askokcancel(messages[0], messages[1])

    def print_error(err):
        print(' --- Error message --- ')
        print(f" type   :[{str(type(err))}]")
        print(f" args   :[{str(err.args)}]")
        print(f" message:[{err.message}]")
        print(f" error  :[{err}]")

    def get_message(result):
        list_message = []
        messages = []
        list_message.append('[コード：')
        list_message.append(str(result))
        list_message.append(']\r\n')
        # formPackageMain
        if result == 0:
            list_message.append('　正常終了しました。')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == 9999:
            list_message.append('　 ---- ')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == 9000:
            list_message.append('　実行中です。しらばくお待ちください……')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == 9001:
            list_message.append('　実行しますか？')
            messages.append('実行有無')
            messages.append(''.join(list_message))
        elif result == 9002:
            list_message.append('　処理をキャンセルしました。')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == -1001:
            list_message.append('　チェックモード切り替え時にエラーが発生しました。')
            messages.append('モード切り替え｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1002:
            list_message.append('　パッケージ化モードに戻す際にエラーが発生しました。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1101:
            list_message.append('　必須項目「対象フォルダー」に値がありません。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1102:
            list_message.append('　項目「対象フォルダー」の指定先にフォルダーがありません。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1111:
            list_message.append('　必須項目「作業対象」に値がありません。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1121:
            list_message.append('　必須項目「作業日付」に値がないか正しい値ではありません。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1131:
            list_message.append('　必須項目「部署名・作業者名」に値がありません。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1141:
            list_message.append('　必須項目「作業端末名」に値がありません。')
            messages.append('必須項目エラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1201:
            list_message.append('　対象フォルダー全体のファイルサイズが閾値を超えています。')
            messages.append('ファイル数とファイルサイズのチェックエラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1202:
            list_message.append('　対象フォルダー内のファイル数が閾値を超えています。')
            messages.append('ファイル数とファイルサイズのチェックエラー｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1301:
            list_message.append('　インストールフォルダーから帳票ファイルをコピーする処理でエラーが発生しました。')
            messages.append('帳票ファイルコピーのエラー｜formPackageMain')
            messages.append(''.join(list_message))
        # ctrlCsv
        elif result == -3001:
            list_message.append('　CSVファイルを出力中にエラーが発生しました。')
            messages.append('CSV出力処理エラー｜ctrlCsv')
            messages.append(''.join(list_message))
        elif result == -3101:
            list_message.append('　Manifest.xmlをCSVの中間ファイルに変換中にエラーが発生しました。')
            messages.append('XMLからCSVへの変換処理エラー｜ctrlCsv')
            messages.append(''.join(list_message))
        elif result == -3201:
            list_message.append('　複数のManifest.xmlをCSVの中間ファイルに変換中にエラーが発生しました。')
            messages.append('複数XMLからCSVへの変換処理エラー｜ctrlCsv')
            messages.append(''.join(list_message))
        # ctrlBatch
        elif result == -4101:
            list_message.append('　PowerShellのバージョンの取得処理で例外的なエラーが発生しました。（バージョン情報が数値以外）')
            messages.append('バッチ処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4102:
            list_message.append('　PowerShellのバージョンが7未満です。')
            messages.append('バッチ処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4111:
            list_message.append('　pwshコマンドが見つかりません。PowerShell 7以降が未導入かインストール後に再起動されていない。もしくは環境変数の定義内容を見直してください。')
            messages.append('バッチ処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4112:
            list_message.append('　pwshのバージョンチェック中にエラーが発生しました。')
            messages.append('バッチ処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4201:
            list_message.append('　Windows環境で実行してください。')
            messages.append('バッチ処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4202:
            list_message.append('　PowerShellの実行時にエラーが発生しました。')
            messages.append('バッチ処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4301:
            list_message.append('　実行後に画面の入力内容を変更されているので格納先を開けません。')
            messages.append('格納先を開く処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4302:
            list_message.append('　格納先を開く処理でエラーが発生しました。')
            messages.append('格納先を開く処理エラー｜ctrlBatch')
            messages.append(''.join(list_message))
        # ctrlExcel
        elif result == -5001:
            list_message.append('　Office 2019以降のExcelがインストールされていません。')
            messages.append('Microsoft Excel エラー｜ctrlExcel')
            messages.append(''.join(list_message))
        elif result == -5002:
            list_message.append('　Excelアプリが参照できません。Office 2019以降をインストールしてください。')
            messages.append('Microsoft Excel エラー｜ctrlExcel')
            messages.append(''.join(list_message))
        # CommonModules.psm1
        elif result == -6001:
            list_message.append('　このPowerShellはバージョン7.0以降で実行する必要があります。')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6101:
            list_message.append('　フォルダーの削除処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6102:
            list_message.append('　ファイルの削除処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6201:
            list_message.append('　ZIPファイルへの圧縮処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6301:
            list_message.append('　ZIPファイルの解凍処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6401:
            list_message.append('　フォルダー再作成処理で新規フォルダーを作成した時にエラーが発生しました。場所: New-DirectoryIfNotExists')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6402:
            list_message.append('　フォルダー再作成処理でフォルダーを再作成した時にエラーが発生しました。場所: New-DirectoryIfNotExists')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6501:
            list_message.append('　パッケージ化前後のデータを比較し、差異が発生しました。')
            messages.append('PowerShellスクリプトエラー｜CommonModules.psm1')
            messages.append(''.join(list_message))
        # AdpackController.ps1
        elif result == -7001:
            list_message.append('　必要な外部モジュールファイル（*.ps1, *.psm1）が存在しません。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7002:
            list_message.append('　外部モジュールファイル（*.ps1, *.psm1）の読み込み処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7003:
            list_message.append('　引数を見直してください。[理由：Pack + UnPack 両方を設定]')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7004:
            list_message.append('　引数を見直してください。[理由：Pack と UnPack どちらも未設定]')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7005:
            list_message.append('　引数を見直してください。[理由：Pack + Check、もしくは Pack + UnCheck で設定]')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7006:
            list_message.append('　引数を見直してください。[理由：UnPack + Check + UnCheck 3つを設定]')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7007:
            list_message.append('　入力データは「フォルダー」を指定してください。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7008:
            list_message.append('　パックでの出力データは「ファイル（*.zip）」を指定してください。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7009:
            list_message.append('　アンパックでの入力データは「ファイル（*.zip）」を指定してください。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7010:
            list_message.append('　アンパックでの出力データは「フォルダー」を指定してください。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7011:
            list_message.append('　Adpackの実行ファイルが存在しません。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        # AdpackModules.psm1
        elif result == -7101:
            list_message.append('　XMLファイルを格納するフォルダーの作成でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7102:
            list_message.append('　Index.xmlを作成する際にエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7103:
            list_message.append('　Manifest.xmlを作成する際にエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7104:
            list_message.append('　ハッシュチェックで一致しないファイルがあります。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7105:
            list_message.append('　AdPackを使ったパッケージ化処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7106:
            list_message.append('　AdPackを使ったパッケージ化処理で例外的なエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7107:
            list_message.append('　AdPackを使ったチェック処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7108:
            list_message.append('　AdPackを使ったアンパック処理（NoCheck引数あり）でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7109:
            list_message.append('　AdPackを使ったチェック処理後にある解凍データの削除でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7110:
            list_message.append('　自作モジュールのアンパック処理対象のデータにメタデータ（META-INF）が含まれていません。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7111:
            list_message.append('　自作モジュールのアンパック処理対象のデータとメタデータ（META-INF）の情報が一致していません。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7112:
            list_message.append('　自作モジュールのアンパック処理で一時フォルダーの名前変更でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜AdpackController.ps1')
            messages.append(''.join(list_message))
        # MultiCheckController.ps1
        elif result == -7201:
            list_message.append('　必要な外部モジュールファイル（*.ps1, *.psm1）が存在しません。')
            messages.append('PowerShellスクリプトエラー｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7202:
            list_message.append('　外部モジュールファイル（*.ps1, *.psm1）の読み込み処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7203:
            list_message.append('　入力データは「フォルダー」を指定してください。')
            messages.append('PowerShellスクリプトエラー｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7204:
            list_message.append('　Adpackの実行ファイルが存在しません。')
            messages.append('PowerShellスクリプトエラー｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7205:
            list_message.append('　指定されたフォルダーにZIPファイルが存在しません。')
            messages.append('PowerShellスクリプトエラー｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        # PrintController.ps1
        elif result == -8001:
            list_message.append('　必要な外部モジュールファイル（*.ps1, *.psm1）が存在しません。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8002:
            list_message.append('　外部モジュールファイル（*.psm1）の読み込み処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8003:
            list_message.append('　引数を見直してください。[理由：PackForm + CheckForm 両方を設定]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8004:
            list_message.append('　引数を見直してください。[理由：PackForm と CheckForm どちらも未設定]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8005:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：RootPath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8006:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：TemplatePath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8007:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：formPath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8008:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：headerMappingPath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8009:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：bodyMappingPath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8010:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：headerValuesPath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8011:
            list_message.append('　引数で指定されたフォルダー、もしくはファイルが存在しません。[対象：bodyValuesPath]')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8012:
            list_message.append('　帳票テンプレートファイル（Excel）内に既定のシートがありませんでした。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8013:
            list_message.append('　帳票テンプレートファイル（Excel）のコピー処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8014:
            list_message.append('　ヘッダー情報の位置データと入力データの読み込みでエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8015:
            list_message.append('　ヘッダー情報における 位置データ と 入力データ の項目名が一致しません。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8016:
            list_message.append('　メイン情報の登録内、データを読み込みでエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8017:
            list_message.append('　データ本体の書き込み位置データ と 入力データ の項目名が一致しません。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8018:
            list_message.append('　帳票ファイル（Excel）に入力データを設定する処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8019:
            list_message.append('　Excelファイル テンプレートシートの削除処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8020:
            list_message.append('　PDFファイルの出力処理でエラーが発生しました。')
            messages.append('PowerShellスクリプトエラー｜PrintController.ps1')
            messages.append(''.join(list_message))
        # 画面系のエラーメッセージ
        elif result == -9001:
            list_message.append('　多重起動は禁止されています。')
            messages.append('警告：多重起動禁止')
            messages.append(''.join(list_message))
        else:
            list_message.append('　例外が発生し異常終了しました。')
            messages.append('')
            messages.append(''.join(list_message))
        return messages
