import tkinter as tk
from tkinter import simpledialog

class formAuth(simpledialog.Dialog):
    def __init__(self, parent, title):
        super().__init__(parent, title)

    def body(self, master):
        """ ダイアログのメイン部分（パスワード入力） """
        self.geometry("300x120")  # サイズ固定
        self.resizable(False, False)  # サイズ変更不可
        self.overrideredirect(True)  # 最大化・最小化を無効化
        self.attributes("-topmost", True) # 最上位で表示
        
        self.frame_border = tk.Frame(self, highlightcolor='#ffffff', highlightthickness=3)
        self.frame_border.pack(padx=3, pady=3, fill=tk.BOTH, expand=True)

        tk.Label(self.frame_border, text="パスワードを入力してください:").pack(pady=10)
        self.entry = tk.Entry(self.frame_border, show="*")
        self.entry.pack(pady=5)
        self.entry.focus_set()
        self.entry.bind("<Return>", lambda event: self.apply())

        return self.entry
    
    # ボタンの領域を定義
    def buttonbox(self):
        _frame_buttunbox = tk.Frame(self.frame_border)

        _btn_ok = tk.Button(_frame_buttunbox, text="OK", width=10, command=self.apply)
        _btn_ok.pack(side=tk.LEFT, padx=5, pady=5)

        _btn_cancel = tk.Button(_frame_buttunbox, text="キャンセル", width=10, command=self.cancel)
        _btn_cancel.pack(side=tk.LEFT, padx=5, pady=5)

        _frame_buttunbox.pack()
        
    # OKボタンの動作
    def apply(self):
        self.result = self.entry.get()
        self.master.focus_set()
        self.destroy()
    
    # キャンセルボタンの動作
    def cancel(self, event=None):
        if self.result is None:
            self.result = None
        super().cancel()
