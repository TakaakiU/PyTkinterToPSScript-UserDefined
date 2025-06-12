import sys
import os
import subprocess
import platform
from pathlib import Path

from classes.control import ctrlCommon
from classes.control import ctrlMessage


class ctrlBatch():
    # ファイルのパスを取得
    def get_path(relative_path):
        if hasattr(sys, '_MEIPASS'):
            # EXE実行時のパス
            return os.path.join(sys._MEIPASS, relative_path)
        else:
            # デバッグ中のパス
            return os.path.join(
                os.path.abspath("."),
                'classes/script',
                relative_path)
    
    def _test_pwsh():
        # 事前チェック：pwshが正常に実行できるか、バージョンが7以上か確認します。
        try:
            # $PSVersionTable.PSVersion.Major でメジャーバージョンを取得する
            version_proc = subprocess.run(
                ["pwsh", "-NoProfile", "-Command", "$PSVersionTable.PSVersion.Major"],
                capture_output=True, text=True
            )
            
            # コマンド実行に失敗した場合
            if version_proc.returncode != 0:
                ctrlMessage.print_error("PowerShellのバージョン確認に失敗しました。pwshが正常に動作しているか確認してください。")
                return -4001
            
            major_version_str = version_proc.stdout.strip()

            _result = 0
            try:
                major_version = int(major_version_str)
            except ValueError:
                ctrlMessage.print_error(f"PowerShellのバージョンの取得処理で例外的なエラーが発生しました。（バージョン情報が数値以外）: {major_version_str}")
                _result = -4101
            
            if major_version < 7:
                ctrlMessage.print_error(f"PowerShellのバージョンが7未満です (実際のバージョン: {major_version})。PowerShell 7以降が必要です。")
                _result = -4102
        
        except FileNotFoundError:
            ctrlMessage.print_error("pwshコマンドが見つかりません。PowerShell 7以降が未導入かインストール後に再起動されていない。もしくは環境変数の定義内容を見直してください。")
            _result = -4111
        except Exception as err:
            ctrlMessage.print_error(f"pwshのバージョンチェック中にエラーが発生しました: {err}")
            _result = -4112
            
        return _result

    def exe_powershell(script_path, *args):
        _result = 0

        # pwshの実行テスト
        _result = ctrlBatch._test_pwsh()

        # PowerShellスクリプト実行時の引数を準備
        if _result == 0:
            pscommand = [
                'pwsh',
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-file',
                script_path
            ]
            # pscommand = [
            #     'powershell',
            #     '-NoProfile',
            #     '-ExecutionPolicy',
            #     'Bypass',
            #     '-file',
            #     script_path
            # ]

            # PowerShellスクリプトの引数がある場合は追加
            if args:
                pscommand.extend(args)

            # Windows OS以外は処理を中断
            osname = platform.system()
            if osname != 'Windows':
                _result = -4201
            
            if _result == 0:
                try:
                    CREATE_NO_WINDOW = 0x08000000
                    _compps = subprocess.run(pscommand, creationflags=subprocess.CREATE_NO_WINDOW)
                    _result = _compps.returncode
                    # 戻り値が負の値の場合は符号なし整数として受け取るため、負の値に変換
                    if _result > 2**31 - 1:
                        _result -= 2**32
                except Exception as err:
                    _result = -4202
                    ctrlMessage.print_error(err)
        
        return _result
    
    def open_folder(target_path):
        _result = 0

        _open_path = Path(target_path)
        if not _open_path.exists():
            _result = -4301
        elif sys.platform.startswith("win"):
            try:
                subprocess.Popen(["explorer", str(_open_path)])
            except Exception as err:
                _result = -4302
                ctrlMessage.print_error(err)
        
        return _result