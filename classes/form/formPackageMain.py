import sys
import os
import shutil
from  pathlib import Path
import tkinter as tk
import tkinter.ttk as ttk
import tkinterdnd2 as tkdnd2
import tkcalendar as tkcal
from tkinter import PhotoImage
from tkinter import filedialog
from datetime import datetime

# 自作モジュールの読み込み
from classes.structure import structureEntrydata
from classes.control import ctrlBatch
from classes.control import ctrlConfig
from classes.control import ctrlMessage
from classes.control import ctrlString
from classes.control import ctrlCsv
from classes.control import ctrlExcel
from classes.form import formBase
from classes.form.formAuth import formAuth

class formPackageMain(formBase):
    def __init__(self, master):
        # 継承元を呼び出し
        super().__init__(master)

        # 初期設定
        # _xmldata = ctrlConfig.get_xmldata('settings.xml')
        self.settings["installdir"] = 'C:/PyTkinterToPSScript'
        self.settings_filepath = self.settings['installdir'] + '/config/settings.xml'
        _xmldata = ctrlConfig.get_xmldata(self.settings_filepath)
        self.settings = self._add_settings(_xmldata)

        # オブジェクト作成
        self.master.title(self.settings.get('title'))
        self.master.geometry(self.settings.get('window_size'))
        self._create_header()
        self._create_body()
        self._create_footer()
        master.bind('<KeyPress>', self._key_handler)
        
        # WM_DELETE_WINDOW イベントに独自の終了処理をバインド
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.master.destory()
        sys.exit(0)
        
    def __del__(self):
        print('close_process')

    def _add_settings(self, xmldata):
        # 継承元の設定
        settings = self.settings

        # xmlファイルの読み込み
        settings = ctrlConfig.read_xmlfile(settings, xmldata)

        #   コンボボックス設定
        for combosettings in xmldata.findall('combosettings'):
            _array = []
            for _combodata in combosettings.findall('targetrange'):
                for _value in _combodata.findall('value'):
                    _array.append(_value.text)
                settings['targetrange'] = _array
            _array = []
            for _combodata in combosettings.findall('workername'):
                for _value in _combodata.findall('value'):
                    _array.append(_value.text)
                settings['workername'] = _array
            _array = []
            for _combodata in combosettings.findall('terminalname'):
                for _value in _combodata.findall('value'):
                    _array.append(_value.text)
                settings['terminalname'] = _array

        # 追加の設定
        settings['title'] = 'PyTkinterToPSScript（自作関数版）'
        settings['label_logo'] = 'PyTkinterToPSScript'
        settings["label_result_init"] = ' ---- '
        settings['label_result'] = '        '
        settings['text_messages'] = 'メッセージはありません。'
        settings['label_mode_package'] = '現在｜パッケージ化 モード'
        settings['label_mode_check'] =   '現在｜チェック モード        '
        settings['button_modechange'] = 'モード変更'
        settings['label_basicsettings'] = '基本設定'
        settings['label_targetfolder'] = '対象フォルダー'
        settings['textbox_folderpath'] = ''
        settings['button_databrose'] = '参照'
        settings['label_reportsettings'] = '帳票設定'
        settings['label_targetrange'] = '作業対象'
        settings['label_workdate'] = '作業日'
        settings['label_workername'] = '部署名・作業者名'
        settings['label_terminalname'] = '作業端末名'
        #   文字色
        settings['fg_color'] = '#696969'
        settings['fg_color_package'] = '#ffffff'
        settings['fg_color_check'] = '#ffffff'
        #   背景色
        settings['bg_color'] = '#f2f2f2'
        settings['bg_color_package'] = '#8b0000'
        settings['bg_color_check'] = '#006400'
        #   ラベル設定
        settings['fb_color_ok'] = '#00008b'
        settings['fb_color_ng'] = '#ff0000'
        settings['fb_color_normal'] = '#000000'

        return settings

    # key_hander
    def _key_handler(self, event):
        if (event.keysym == 'F1' and
                self._button_f1['state'] == 'normal'):
            self._click_f1()
        elif (event.keysym == 'F2' and
                self._button_f2['state'] == 'normal'):
            self._click_f2()
        elif (event.keysym == 'F3' and
                self._button_f3['state'] == 'normal'):
            self._click_f3()
        elif (event.keysym == 'F4' and
                self._button_f4['state'] == 'normal'):
            self._click_f4()
        elif (event.keysym == 'F5' and
                self._button_f5['state'] == 'normal'):
            self._click_f5()
        elif (event.keysym == 'F6' and
                self._button_f6['state'] == 'normal'):
            self._click_f6()
        elif (event.keysym == 'F7' and
                self._button_f7['state'] == 'normal'):
            self._click_f7()
        elif (event.keysym == 'F8' and
                self._button_f8['state'] == 'normal'):
            self._click_f8()
        elif (event.keysym == 'F9' and
                self._button_f9['state'] == 'normal'):
            self._click_f9()
        elif (event.keysym == 'F10' and
                self._button_f10['state'] == 'normal'):
            self._click_f10()
        elif (event.keysym == 'F11' and
                self._button_f11['state'] == 'normal'):
            self._click_f11()
        elif (event.keysym == 'F12' and
                self._button_f12['state'] == 'normal'):
            self._click_f12()

    # コントロール定義 --->

    # ヘッダーエリア
    def _create_header(self):
        # ロゴ
        self._current_dir = os.path.dirname(__file__)
        self._image_path = os.path.join(
            self._current_dir,
            '..',
            'image', 'logo.png'
        )
        self._logoImage = PhotoImage(file=self._image_path)
        self._label_logo = tk.Label(
            self.frame_header,
            width=340,
            height=60,
            image=self._logoImage
        )
        self._label_logo.grid(row=0, column=0, sticky="w")
        # 処理結果
        self._label_result = tk.Label(
            self.frame_header,
            text=self.settings['label_result'],
            font=self.settings['font_label_h2'],
            bg=self.settings['bg_color']
        )
        self._label_result.grid(row=0, column=1, sticky="e")
        # メッセージ内容
        _frame_messages = tk.Frame(
            self.frame_header,
            bg=self.settings['bg_color'],
            padx=0,
            pady=0
        )
        _frame_messages.grid(row=0, column=2, columnspan=3, sticky="e")

        self._text_messages = tk.Text(
            _frame_messages,
            width=55,
            height=3,
            wrap=tk.NONE,
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._text_messages.grid(row=0, column=0, sticky="nsew")

        # スクロールを付与
        scrollbar_y = tk.Scrollbar(
            _frame_messages,
            command=self._text_messages.yview
        )
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x = tk.Scrollbar(
            _frame_messages,
            orient='horizontal',
            command=self._text_messages.xview
        )
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        # スクロールを適用
        self._text_messages.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # 初期値を反映
        self._text_messages.configure(state='normal')
        self._text_messages.delete('1.0', 'end')
        self._text_messages.insert('1.0', self.settings['text_messages'])
        self._text_messages.configure(state='disabled')
        # モード
        self._label_mode = tk.Label(
            self.frame_header,
            text=self.settings['label_mode_package'],
            font=self.settings['font_label_h2'],
            fg=self.settings['fg_color_package'],
            bg=self.settings['bg_color_package']
        )
        self._label_mode.grid(row=1, column=3, sticky="e")
        # モード変更ボタン
        self._button_modechange = tk.Button(
            self.frame_header,
            text=self.settings['button_modechange'],
            font=self.settings['font_button'],
            command=self._click_modechange)
        self._button_modechange.grid(row=1, column=4, sticky="e")
        # レイアウト
        self.frame_header.rowconfigure(0, weight=2)
        self.frame_header.rowconfigure(1, weight=1)
        self.frame_header.columnconfigure(0, weight=2)  # ロゴ
        self.frame_header.columnconfigure(1, weight=1)  # 処理結果
        self.frame_header.columnconfigure(2, weight=20)  # メッセージ(1/3)
        self.frame_header.columnconfigure(3, weight=1)  # メッセージ(2/3) + モードラベル
        self.frame_header.columnconfigure(4, weight=1)  # メッセージ(3/3) + モードボタン

    # メインエリア
    def _create_body(self):
        # 基本情報
        self._frame_basicsettings = tk.Frame(
            self.frame_body,
            bg=self.settings['bg_color']
        )
        self._frame_basicsettings.pack(expand=1, fill="both", side="top")
        # ラベル
        self._label_basicsettings = tk.Label(
            self._frame_basicsettings,
            text=self.settings['label_basicsettings'],
            font=self.settings['font_label_h1'],
            bg=self.settings['bg_color']
        )
        self._label_basicsettings.grid(row=0, column=0, sticky="w")
        # 対象フォルダー - ラベル
        self._label_targetfolder = tk.Label(
            self._frame_basicsettings,
            text=self.settings['label_targetfolder'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_targetfolder.grid(row=1, column=0, sticky="w")
        # 対象フォルダー - 入力ボックス
        self._textbox_folderpath = tk.Entry(
            self._frame_basicsettings,
            textvariable=self.settings['textbox_folderpath'],
            font=self.settings['font_text']
        )
        self._textbox_folderpath.grid(row=1, column=1, sticky="ew")
        self._textbox_folderpath.drop_target_register(tkdnd2.DND_FILES)
        self._textbox_folderpath.dnd_bind('<<Drop>>', self.drop)
        # 参照ボタン
        self._button_databrowse = tk.Button(
            self._frame_basicsettings,
            text=self.settings['button_databrose'],
            font=self.settings['font_button'],
            command=self._click_browse)
        self._button_databrowse.grid(row=1, column=2, sticky="ew")
        # レイアウト
        self._frame_basicsettings.columnconfigure(0, weight=1)
        self._frame_basicsettings.columnconfigure(1, weight=7)
        self._frame_basicsettings.columnconfigure(2, weight=1)

        # 帳票情報
        self._frame_reportsettings = tk.Frame(
            self.frame_body,
            bg=self.settings['bg_color']
        )
        self._frame_reportsettings.pack(expand=1, fill="both", side="bottom")
        # ラベル
        self._label_reportsettings = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_reportsettings'],
            font=self.settings['font_label_h1'],
            bg=self.settings['bg_color']
        )
        self._label_reportsettings.grid(row=0, column=0, sticky="w")
        # 作業対象 - ラベル
        self._label_targetrange = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_targetrange'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_targetrange.grid(row=1, column=0, sticky="w")
        # 作業対象 - コンボボックス
        self._combobox_targetrange = ttk.Combobox(
            self._frame_reportsettings,
            state='normal',
            values=self.settings['targetrange'],
            font=self.settings['font_text']
        )
        self._combobox_targetrange.grid(row=1, column=1, sticky="ew")
        # 作業日 - ラベル
        self._label_workdate = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_workdate'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_workdate.grid(row=2, column=0, sticky="w")
        # 作業日 - カレンダー
        self._dateentry_workdate = tkcal.DateEntry(
            self._frame_reportsettings,
            font=self.settings['font_cal'],
            date_pattern='mm/dd/yyyy',
            firstweekday='sunday',
            showweeknumbers=False)
        self._dateentry_workdate.grid(row=2, column=1, sticky="w")
        self._dateentry_workdate.delete(0, tk.END)
        # 部署名・作業者名 - ラベル
        self._label_workername = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_workername'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_workername.grid(row=3, column=0, sticky="w")
        # 部署名・作業者名 - コンボボックス
        self._combobox_workername = ttk.Combobox(
            self._frame_reportsettings,
            state='normal',
            values=self.settings['workername'],
            font=self.settings['font_text']
        )
        self._combobox_workername.grid(row=3, column=1, sticky="ew")
        # 作業端末名 - ラベル
        self._label_terminalname = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_terminalname'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_terminalname.grid(row=4, column=0, sticky="w")
        # 作業端末名 - コンボボックス
        self._combobox_terminalname = ttk.Combobox(
            self._frame_reportsettings,
            state='normal',
            values=self.settings['terminalname'],
            font=self.settings['font_text']
        )
        self._combobox_terminalname.grid(row=4, column=1, sticky="ew")
        # レイアウト
        self._frame_reportsettings.columnconfigure(0, weight=1)
        self._frame_reportsettings.columnconfigure(1, weight=6)

    def _create_footer(self):
        # F1ボタンの作成
        self._label_f1 = tk.Label(
            self.frame_footer, text='F1',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f1.grid(row=0, column=0, sticky="ew")
        self._button_f1 = tk.Button(
            self.frame_footer,
            text=' パッケージ ',
            font=self.settings['font_button'],
            command=self._click_f1
        )
        self._button_f1.grid(row=1, column=0, sticky="ew")
        # F2ボタンの作成
        self._label_f2 = tk.Label(
            self.frame_footer,
            text='F2',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f2.grid(row=0, column=1, sticky="ew")
        self._button_f2 = tk.Button(
            self.frame_footer,
            text=' チェック ',
            font=self.settings['font_button'],
            command=self._click_f2
        )
        self._button_f2.grid(row=1, column=1, sticky="ew")
        self._button_f2.configure(state='disable')
        # F3ボタンの作成
        self._label_f3 = tk.Label(
            self.frame_footer,
            text='F3',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f3.grid(row=0, column=2, sticky="ew")
        self._button_f3 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f3
        )
        self._button_f3.grid(row=1, column=2, sticky="ew")
        self._button_f3.configure(state='disable')
        # F4ボタンの作成
        self._label_f4 = tk.Label(
            self.frame_footer,
            text='F4',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f4.grid(row=0, column=3, sticky="ew")
        self._button_f4 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f4
        )
        self._button_f4.grid(row=1, column=3, sticky="ew")
        self._button_f4.configure(state='disable')
        # スペース
        self._label_space01 = tk.Label(
            self.frame_footer,
            text=' ',
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_space01.grid(row=0, column=4, rowspan=2, sticky="ew")
        # F5ボタンの作成
        self._label_f5 = tk.Label(
            self.frame_footer,
            text='F5',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f5.grid(row=0, column=5, sticky="ew")
        self._button_f5 = tk.Button(
            self.frame_footer,
            text='出力先を開く',
            font=self.settings['font_button'],
            command=self._click_f5
        )
        self._button_f5.grid(row=1, column=5, sticky="ew")
        self._button_f5.configure(state='disable')
        # F6ボタンの作成
        self._label_f6 = tk.Label(
            self.frame_footer,
            text='F6',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f6.grid(row=0, column=6, sticky="ew")
        self._button_f6 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f6
        )
        self._button_f6.grid(row=1, column=6, sticky="ew")
        self._button_f6.configure(state='disable')
        # F7ボタンの作成
        self._label_f7 = tk.Label(
            self.frame_footer,
            text='F7',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f7.grid(row=0, column=7, sticky="ew")
        self._button_f7 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f7
        )
        self._button_f7.grid(row=1, column=7, sticky="ew")
        self._button_f7.configure(state='disable')
        # F8ボタンの作成
        self._label_f8 = tk.Label(
            self.frame_footer,
            text='F8',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f8.grid(row=0, column=8, sticky="ew")
        self._button_f8 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f8
        )
        self._button_f8.grid(row=1, column=8, sticky="ew")
        self._button_f8.configure(state='disable')
        # スペース
        self._label_space02 = tk.Label(
            self.frame_footer,
            text=' ',
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_space02.grid(row=0, column=9, rowspan=2, sticky="ew")
        # F9ボタンの作成
        self._label_f9 = tk.Label(
            self.frame_footer,
            text='F9',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f9.grid(row=0, column=10, sticky="ew")
        self._button_f9 = tk.Button(
            self.frame_footer,
            text=' 設定 ',
            font=self.settings['font_button'],
            command=self._click_f9
        )
        self._button_f9.grid(row=1, column=10, sticky="ew")
        self._button_f9.configure(state='disable')
        # 実行ボタンの作成
        self._label_f10 = tk.Label(
            self.frame_footer,
            text='F10',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f10.grid(row=0, column=11, sticky="ew")
        self._button_f10 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f10
        )
        self._button_f10.grid(row=1, column=11, sticky="ew")
        self._button_f10.configure(state='disable')
        # 設定ボタンの作成
        self._label_f11 = tk.Label(
            self.frame_footer,
            text='F11',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f11.grid(row=0, column=12, sticky="ew")
        self._button_f11 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f11
        )
        self._button_f11.grid(row=1, column=12, sticky="ew")
        self._button_f11.configure(state='disable')
        # 閉じるボタンの作成
        self._label_f12 = tk.Label(
            self.frame_footer,
            text='F12',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f12.grid(row=0, column=13, sticky="ew")
        self._button_f12 = tk.Button(
            self.frame_footer,
            text='閉じる',
            font=self.settings['font_button'],
            command=self._click_f12
        )
        self._button_f12.grid(row=1, column=13, sticky="ew")

        self.frame_footer.columnconfigure(0, weight=2)
        self.frame_footer.columnconfigure(1, weight=2)
        self.frame_footer.columnconfigure(2, weight=2)
        self.frame_footer.columnconfigure(3, weight=2)
        self.frame_footer.columnconfigure(4, weight=1)
        self.frame_footer.columnconfigure(5, weight=2)
        self.frame_footer.columnconfigure(6, weight=2)
        self.frame_footer.columnconfigure(7, weight=2)
        self.frame_footer.columnconfigure(8, weight=2)
        self.frame_footer.columnconfigure(9, weight=1)
        self.frame_footer.columnconfigure(10, weight=2)
        self.frame_footer.columnconfigure(11, weight=2)
        self.frame_footer.columnconfigure(12, weight=2)
        self.frame_footer.columnconfigure(13, weight=2)

    # ドロップ制御
    def drop(self, event):
        self._textbox_folderpath.delete(0, tk.END)
        self._textbox_folderpath.insert(tk.END, event.data)

    # ボタン処理関連 --->

    # モード変更ボタン
    def _click_modechange(self):
        _result = 0

        # 設定ファイルのパスワードを読み込む
        correct_password = self.settings['password']
        # チェックモードに切り替える場合はパスワード認証
        if self._button_f2["state"] == 'disabled':
            _password = formAuth(self.master, "パスワード認証").result
            # パスワード認証でキャンセルした場合
            if _password is None:
                _result = 9999

            # パスワード認証が成功
            elif _password == correct_password:
                try:
                    # 2でチェックモードに切り替え
                    self._change_mode(2)
                    
                except Exception as err:
                    _result = -1301
            # パスワード認証が失敗
            else:
                tk.messagebox.showerror('パスワード認証失敗', 'パスワードが間違っています。')
                _result = 9999

        # パッケージ化モードに切り替える場合は、そのまま切り替え
        elif self._button_f2["state"] == 'normal':
            try:
                # 1のパッケージモードに戻す
                self._change_mode(1)
            except Exception as err:
                _result = -1302

        # 正常時はメッセージを出力しない
        if _result != 0:
            self._view_message(_result)
        
        return _result
    
    #   参照ボタン
    def _click_browse(self):
        # 初期のディレクトリ位置
        if not os.path.isdir(self._textbox_folderpath.get()):
            currentdir = os.getcwd()
            initdir = currentdir
        else:
            initdir = os.path.dirname(self._textbox_folderpath.get())
        # ダイアログ表示
        folder_path = filedialog.askdirectory(initialdir=initdir)
        if folder_path:
            self._textbox_folderpath.delete(0, tk.END)
            self._textbox_folderpath.insert(tk.END, folder_path)

    # F1ボタン
    def _click_f1(self):
        result = 0
        # 実行前の確認メッセージ
        if not (ctrlMessage.view_top_okcancel(9001)) > 0:
            result = 9002
        # 実行中のメッセージ
        if result == 0:
            self._view_message(9000)
        # 入力チェック
        if result == 0:
            entry_data = self._get_entry_data()
            result = self._validate_entry(entry_data)
        # 入力情報をCSVファイル
        if result == 0:
            output_path = self.settings['installdir'] + '/input/Packagelist_HeaderValues.csv'
            result = self._output_csv_headervalues(output_path)
        # ファイル数とファイルサイズをチェック
        if result == 0:
            max_size = int(self.settings['package_maxsize'])        # パッケージのファイルサイズ 1.5GB = 1,610,612,736 = 1.5 * 1GB(1,073,741,824B)
            max_files = int(self.settings['package_maxfiles'])
            # 既定のファイル数とファイルサイズでチェック
            target_path = entry_data.folderpath
            result = self._check_folder_limits(target_path, max_size, max_files)
        # パッケージ化
        if result == 0:
            # 実行するPowerShellスクリプトを指定
            script_path = self.settings['installdir'] + '/script/AdpackController.ps1'
            # PowerShellスクリプトの引数を指定
            hash_algorithm = self.settings['hash_algorithm']
            input_path = entry_data.folderpath
            output_path = '{}.zip'.format(input_path)
            output_path = self._get_unique_filename(output_path)
            # パッケージ化を実行
            result, zip_path = self._packaging_data(script_path, hash_algorithm, input_path, output_path)
        # パッケージ化後のManifestファイルをCSVファイル
        if result == 0:
            input_path = entry_data.folderpath + '/META-INF/Manifest.xml'
            output_path = self.settings['installdir'] + '/input/Packagelist_BodyValues.csv'
            result = self._output_csv_packagelist_bodyvalues(input_path, output_path)
        # パッケージリストを出力
        if result == 0:
            script_path = self.settings['installdir'] + '/script/PrintController.ps1'
            script_args = {
                'output_path': self.settings['installdir'] + '/output/PackageLists.pdf',
                'root_path': self.settings['installdir'],
                'datamapping_header': 'DataMapping_Header.csv',
                'dataMapping_body': 'DataMapping_Body.csv',
                'packform_template': 'Template_PackFormLists.xlsx',
                'packform_headerValues': 'Packagelist_HeaderValues.csv',
                'packform_bodyValues': 'Packagelist_BodyValues.csv',
                'checkform_template': 'Template_CheckFormLists.xlsx',
                'checkform_headerValues': 'Checklist_HeaderValues.csv',
                'checkform_bodyValues': 'Checklist_BodyValues.csv'
            }
            result = self._print_packagelists(script_path, script_args)
        # コピー処理（インストールフォルダーから対象フォルダーにコピー）
        if result == 0:
            # インストールフォルダー内のデータを指定
            from_path = self.settings['installdir'] + '/output/PackageLists.pdf'
            # ダウンロード先（コピー先）を指定
            _target_folder = Path(entry_data.folderpath)
            to_path = _target_folder.parent.as_posix() + '/' + _target_folder.name + '.pdf'
            to_path = self._get_unique_filename(to_path)
            # ダウンロード処理（コピー）
            result, pdf_path = self._copy_formdata(from_path, to_path)
        # コピー正常終了した場合は“格納先を開く”ボタンを有効化
        if result == 0:
            self._button_f5.configure(state='normal')
        # メッセージ出力
        self._view_message(result)
        # 特定ステータスコードの場合にメッセージを追記
        list_message = []
        # 正常終了した場合
        if result == 0:
            list_message.append('\r\n')
            list_message.append('　・ZIPファイル: ' + zip_path + '\r\n')
            list_message.append('　・PDFファイル: ' + pdf_path)
        # メッセージを追記
        self._append_message("".join(list_message))
    
    # F2ボタン
    def _click_f2(self):
        result = 0
        # 実行前の確認メッセージ
        if not (ctrlMessage.view_top_okcancel(9001)) > 0:
            result = 9002
        # 実行中に変更
        if result == 0:
            self._view_message(9000)
        # 入力チェック
        if result == 0:
            entry_data = self._get_entry_data()
            result = self._validate_entry(entry_data)
        # 入力情報をCSVファイル
        if result == 0:
            output_path = self.settings['installdir'] + '/input/Checklist_HeaderValues.csv'
            result = self._output_csv_headervalues(output_path)
        # ファイル数とファイルサイズをチェック
        if result == 0:
            target_path = entry_data.folderpath
            max_size = int(self.settings['check_maxsize'])
            max_files = int(self.settings['check_maxfiles'])
            result = self._check_folder_limits(target_path, max_size, max_files)
        # パッケージ化データのチェック
        ERROR_CODES = (-7104, -7107, -7110, -7111) 
        if result == 0:
            # 実行するPowerShellスクリプトを指定
            script_path = self.settings['installdir'] + '/script/MultiCheckController.ps1'
            # PowerShellスクリプトの引数を指定
            input_path = entry_data.folderpath
            output_path = self.settings['installdir'] + '/input/Checklist_ZipFileList.csv'
            result = self._check_data(script_path, input_path, output_path)

            # チェックでエラーが発生した場合はエラーリストを出力
            if result in ERROR_CODES:
                script_path = self.settings['installdir'] + '/script/PrintController.ps1'
                script_args = {
                    'output_path': self.settings['installdir'] + '/output/ErrorLists.pdf',
                    'root_path': self.settings['installdir'],
                    'datamapping_header': 'DataMapping_Header.csv',
                    'dataMapping_body': 'DataMapping_Body.csv',
                    'errorform_template': 'Template_ErrorFormLists.xlsx',
                    'errorform_headerValues': 'Checklist_HeaderValues.csv',
                    'errorform_bodyValues': 'Checklist_ZipFileList.csv'
                }
                _ = self._print_errorlists(script_path, script_args) # ステータスコードは破棄
        # チェック後の中間CSVファイルをもとに帳票本文の実データファイルを生成
        if result == 0:
            _zipfile_lists = ctrlCsv.read_csv(output_path)
            filepath_lists = []
            for _row in _zipfile_lists:
                _file_path = _row['Index']
                _file_path = entry_data.folderpath + '/' + _file_path
                # 拡張子を削除
                if _file_path.lower().endswith('.zip'):
                    _file_path = os.path.splitext(_file_path)[0]
                    
                filepath_lists.append(_file_path)

            output_path = self.settings['installdir'] + '/input/Checklist_BodyValues.csv'
            result = self._output_csv_checklist_bodyvalues(filepath_lists, output_path)
        # パッケージリストを出力
        if result == 0:
            script_path = self.settings['installdir'] + '/script/PrintController.ps1'
            script_args = {
                'output_path': self.settings['installdir'] + '/output/CheckLists.pdf',
                'root_path': self.settings['installdir'],
                'datamapping_header': 'DataMapping_Header.csv',
                'dataMapping_body': 'DataMapping_Body.csv',
                'checkform_template': 'Template_CheckFormLists.xlsx',
                'checkform_headerValues': 'Checklist_HeaderValues.csv',
                'checkform_bodyValues': 'Checklist_BodyValues.csv'
            }
            result = self._print_checklists(script_path, script_args)
        # ダウンロード処理
        if result == 0:
            # インストールフォルダー内のデータを指定
            from_path = self.settings['installdir'] + '/output/CheckLists.pdf'
            # ダウンロード先（コピー先）を指定
            _target_folder = Path(entry_data.folderpath)
            to_path = _target_folder.parent.as_posix() + '/' + _target_folder.name + '.pdf'
            to_path = self._get_unique_filename(to_path)
            result, pdf_path = self._copy_formdata(from_path, to_path)
            # 出力先を開く
            self._button_f5.configure(state='normal')
        elif result in ERROR_CODES:
            # インストールフォルダー内のデータを指定
            from_path = self.settings['installdir'] + '/output/ErrorLists.pdf'
            # ダウンロード先（コピー先）を指定
            _target_folder = Path(entry_data.folderpath)
            to_path = _target_folder.parent.as_posix() + '/' + _target_folder.name + '_Error.pdf'
            to_path = self._get_unique_filename(to_path)
            _, pdf_path = self._copy_formdata(from_path, to_path) # ステータスコードは破棄
            # 出力先を開く
            self._button_f5.configure(state='normal')
        # メッセージ出力
        self._view_message(result)
        
        # 特定ステータスコードの場合にメッセージを追記
        list_message = []
        # 正常終了した場合
        if result == 0:
            list_message.append('\r\n')
            list_message.append('　・PDFファイル: ' + pdf_path)
        # チェックで異常終了した場合
        elif result in ERROR_CODES:
            list_message.append('\r\n')
            list_message.append('　・PDFファイル: ' + pdf_path)
        # メッセージを追記
        self._append_message("".join(list_message))
    
    # F3ボタン
    def _click_f3(self):
        print('_click_f3')
    
    # F4ボタン
    def _click_f4(self):
        print('_click_f4')
    
    # F5ボタン
    def _click_f5(self):
        result = 0

        _entry_data = self._get_entry_data()
        target_path = _entry_data.folderpath

        _pathlib = Path(target_path)
        target_path = str(_pathlib.parent)
        
        result = self._open_output_path(target_path)

        # フォルダーを開く際の正常時はメッセージを出力しない
        if result != 0:
            self._view_message(result)
    
    # F6ボタン
    def _click_f6(self):
        print('_click_f6')

    # F7ボタン
    def _click_f7(self):
        print('_click_f7')

    # F8ボタン
    def _click_f8(self):
        print('_click_f8')

    # F9ボタン
    def _click_f9(self):
        _result = 0

        target_path = self.settings_filepath
        
        # 親フォルダーに設定
        _pathlib = Path(target_path)
        target_path = str(_pathlib.parent)

        _result = self._open_output_path(target_path)

        # 正常時はメッセージを出力しない
        if _result != 0:
            self._view_message(_result)

    # F10ボタン
    def _click_f10(self):
        print('_click_f10')

    # F11ボタン
    def _click_f11(self):
        print('_click_f11')

    # F12ボタン
    def _click_f12(self):
        self.master.destroy()
        sys.exit(0)

    # メソッド処理関連 --->

    def _toggle_buttons(self, correct_password):
        _result = 0

        # チェックモードに切り替える場合
        if self._button_f2["state"] == 'disabled':
            # パスワード認証
            # password = simpledialog.askstring('パスワード認証', 'パスワードを入力してください。', show='*')
            _password = formAuth(self.master, "パスワード認証").result
            # キャンセル
            if _password is None:
                _result = 9999

            # パスワード認証が成功
            elif _password == correct_password:
                try:
                    self._change_mode(1)
                    
                except Exception as err:
                    _result = -1001
            # パスワード認証が失敗
            else:
                tk.messagebox.showerror('パスワード認証失敗', 'パスワードが間違っています。')
                _result = 9999

        # パッケージ化モードに切り替える場合
        elif self._button_f2["state"] == 'normal':
            try:
                self._change_mode(2)
            except Exception as err:
                _result = -1002
        
        return _result
    
    def _change_mode(self, mode):
        # パッケージ化モード
        if mode == 1:
            # モードラベルの変更
            self._label_mode.configure(
                text=self.settings['label_mode_package'],
                fg=self.settings['fg_color_package'],
                bg=self.settings['bg_color_package']
            )
            # ボタンの有効／無効の設定
            self._button_f1.configure(state='normal')
            self._button_f2.configure(state='disabled')
            self._button_f5.configure(state='disabled')
            self._button_f9.configure(state='disabled')
        # チェックモード
        elif mode == 2:
            # モードラベルの変更
            self._label_mode.configure(
                text=self.settings['label_mode_check'],
                fg=self.settings['fg_color_check'],
                bg=self.settings['bg_color_check']
            )
            # ボタンの有効／無効の設定
            self._button_f1.configure(state='disabled')
            self._button_f2.configure(state='normal')
            self._button_f5.configure(state='disabled')
            self._button_f9.configure(state='normal')
    
    # 入力値の配列データを作成
    def _get_entry_data(self):
        entry_data = structureEntrydata.EntryData(
            folderpath = self._textbox_folderpath.get(),
            targetrange = self._combobox_targetrange.get(),
            workdate = self._dateentry_workdate.get(),
            workername = self._combobox_workername.get(),
            terminalname = self._combobox_terminalname.get()
        )
        return entry_data
    
    def _get_csv_data(self):
        _entry_data = self._get_entry_data()
        _csv_data = [
            ["OverallResult","WorkDate","WorkerName","TargetProcess","TerminalName","TargetFolder","Result"],
            ["All OK", _entry_data.workdate, _entry_data.workername, _entry_data.targetrange, _entry_data.terminalname, _entry_data.folderpath, "OK"]
        ]

        return _csv_data
    
    # 入力値のチェック
    def _validate_entry(self, entry_data: structureEntrydata.EntryData):
        _result = 0
        # 対象フォルダー - 入力チェック
        if entry_data.folderpath == "":
            _result = -1101
        # 対象フォルダー - フォルダーの存在チェック
        if _result == 0:
            if not os.path.isdir(entry_data.folderpath):
                _result = -1102

        # 作業対象 - コンボボックス
        if _result == 0:
            if entry_data.targetrange == "":
                _result = -1111
        # 作業日付
        if _result == 0:
            try:
                datetime.strptime(entry_data.workdate, "%m/%d/%Y")
            except ValueError:
                _result = -1121
        # 部署名・作業者名
        if _result == 0:
            if entry_data.workername == "":
                _result = -1131
        # 作業端末名
        if _result == 0:
            if entry_data.terminalname == "":
                _result = -1141
        
        return _result
    
    # 入力値を帳票のヘッダー用のCSVファイルとして出力
    def _output_csv_headervalues(self, output_path):
        _result = 0
        _csv_data = self._get_csv_data()
        _result = ctrlCsv.output_csv(output_path, _csv_data)

        return _result
    
    # 対象フォルダー内のファイル数とファイルサイズ(MB)のチェック
    def _check_folder_limits(self, target_path, max_size, max_files):
        _result = 0

        # 両方「0」の場合はチェックしない
        if max_size == 0 and max_files == 0:
            return _result
        
        total_size, file_count  = self._get_folder_stats(target_path)

        # ファイルサイズのチェック
        if max_size != 0 and total_size > max_size:
                _result = -1201
        if max_files != 0 and file_count > max_files:
                _result = -1202

        return _result
    
    # 対象フォルダーのファイル数とファイルサイズ(MB)を取得
    def _get_folder_stats(self, folder_path):
        total_bytes = 0
        file_count = 0

        for current_dir, _, file_lists in os.walk(folder_path):
            total_bytes += sum(os.path.getsize(os.path.join(current_dir, file)) for file in file_lists)
            file_count += len(file_lists)
        
        return total_bytes, file_count

    # パッケージ化
    def _packaging_data(self, script_path, hash_algorithm, input_path, output_path):
        args = [
            '-Pack',
            '-Hash', hash_algorithm,
            '-InputPath', input_path,
            '-OutputPath', output_path
        ]
        result = ctrlBatch.exe_powershell(script_path, *args)

        return result, output_path
    
    def _get_unique_filename(self, file_path, max_attempts=0):
        # ファイルが存在しない場合はそのまま返す
        if not os.path.exists(file_path):
            return file_path

        base, ext = os.path.splitext(file_path)

        # 試行回数なし（一意のファイルパスが設定できるで無制限に繰り返す）
        if max_attempts == 0:
            counter = 1
            while os.path.exists(f"{base}-{counter}{ext}"):
                counter += 1

            return f"{base}-{counter}{ext}"            
        
        # 試行回数あり
        else:
            for counter in range(1, max_attempts + 1):
                new_file_path = f"{base}-{counter}{ext}"
                if not os.path.exists(new_file_path):
                    return new_file_path
            raise Exception(f"試行回数 {max_attempts} を超えました。適切なファイル名を生成できませんでした。")
    
    def _output_csv_packagelist_bodyvalues(self, input_path, output_path):
        _result = 0

        _result = ctrlCsv.extract_xml_to_csv(input_path, output_path)

        return _result
    
    def _output_csv_checklist_bodyvalues(self, filepath_lists, output_path):
        _result = 0

        _result = ctrlCsv.extract_xmls_to_csv(filepath_lists, output_path)

        return _result

    # パッケージリストを印刷
    def _print_packagelists(self, script_path, script_args):
        _result = 0

        _result = ctrlExcel.testexcel()

        if _result == 0:
            args = [
                '-PackForm',
                '-OutputPath', script_args['output_path'],
                '-RootPath', script_args['root_path'],
                '-DataMapping_Header', script_args['datamapping_header'],
                '-DataMapping_Body', script_args['dataMapping_body'],
                '-PackForm_Template', script_args['packform_template'],
                '-PackForm_HeaderValues', script_args['packform_headerValues'],
                '-PackForm_BodyValues', script_args['packform_bodyValues']
            ]
            _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result

    def _copy_formdata(self, from_path, to_path):
        result = 0

        try:
            shutil.copy(from_path, to_path)
        except:
            result = -1203

        return result, to_path
    
    def _print_checklists(self, script_path, script_args):
        _result = 0

        _result = ctrlExcel.testexcel()

        if _result == 0:    
            args = [
                '-CheckForm',
                '-OutputPath', script_args['output_path'],
                '-RootPath', script_args['root_path'],
                '-DataMapping_Header', script_args['datamapping_header'],
                '-DataMapping_Body', script_args['dataMapping_body'],
                '-CheckForm_Template', script_args['checkform_template'],
                '-CheckForm_HeaderValues', script_args['checkform_headerValues'],
                '-CheckForm_BodyValues', script_args['checkform_bodyValues']
            ]
            _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result
    
    def _print_errorlists(self, script_path, script_args):
        _result = 0

        _result = ctrlExcel.testexcel()

        if _result == 0:    
            args = [
                '-ErrorForm',
                '-OutputPath', script_args['output_path'],
                '-RootPath', script_args['root_path'],
                '-DataMapping_Header', script_args['datamapping_header'],
                '-DataMapping_Body', script_args['dataMapping_body'],
                '-ErrorForm_Template', script_args['errorform_template'],
                '-ErrorForm_HeaderValues', script_args['errorform_headerValues'],
                '-ErrorForm_BodyValues', script_args['errorform_bodyValues']
            ]
            _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result

    def _open_output_path(self, target_path):
        _result = 0

        _result = ctrlBatch.open_folder(target_path)

        return _result
    
    # パッケージ化データのチェック
    def _check_data(self, script_path, input_path, output_path):
        args = [
            '-InputPath', input_path,
            '-OutputPath', output_path
        ]
        _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result

    # 処理結果とメッセージを表示
    def _view_message(self, result):
        # 処理結果ラベルに設定
        if result == 9999:
            self._label_result['text'] = self.settings['label_result_init']
            self._label_result['fg'] = self.settings['fb_color_normal']
        elif result == 9000:
            self._label_result['text'] = ' 実行中 '
            self._label_result['fg'] = self.settings['fb_color_ok']
        elif 9999 > result >= 0:
            self._label_result['text'] = '  完了  '
            self._label_result['fg'] = self.settings['fb_color_ok']
        else:
            self._label_result['text'] = '  異常  '
            self._label_result['fg'] = self.settings['fb_color_ng']
        # メッセージに設定
        datestr = ctrlString.now_label()
        messages = ctrlMessage.get_message(result)
        messages = "{} {}".format(datestr, messages[1])
        self._text_messages.configure(state='normal')
        self._text_messages.delete('1.0', 'end')
        self._text_messages.insert('1.0', messages)
        self._text_messages.configure(state='disabled')
        # 更新
        self._label_result.update()
        self._text_messages.update()
    
    def _append_message(self, messages):
        self._text_messages.configure(state='normal')
        self._text_messages.insert(tk.END, messages)
        self._text_messages.configure(state='disabled')
        # 更新
        self._text_messages.update()
