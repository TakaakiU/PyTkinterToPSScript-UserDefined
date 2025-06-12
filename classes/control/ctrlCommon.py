import sys
import os


class ctrlCommon():
    # ファイルのパスを取得
    def get_path(relative_path, filename):
        if hasattr(sys, '_MEIPASS'):
            # EXE実行時のパス
            return os.path.join(sys._MEIPASS, relative_path, filename)
        else:
            # デバッグ中のパス
            return os.path.join(
                os.path.abspath("."),
                relative_path,
                filename)
