import tkinter as tk

class formBase(object):
    def __init__(self, master):
        # Main window
        self.master = master
        self.settings = formBase.base_settings()
        self.master.title(self.settings.get('title'))
        self.master.geometry(self.settings.get('window_size'))
        self.master.configure(bg=self.settings.get('bg_color'))

        # Area of header
        self.frame_header = tk.Frame(
            self.master,
            bg=self.settings['bg_color'],
            padx=12,
            pady=12
        )
        self.frame_header.pack(expand=1, fill="both", side="top")

        # Area of main
        self.frame_body = tk.Frame(
            self.master,
            bg=self.settings['bg_color'],
            padx=12,
            pady=12
        )
        self.frame_body.pack(expand=1, fill="both", side="top")

        # Area of button
        self.frame_footer = tk.Frame(
            self.master,
            bg=self.settings['bg_color'],
            padx=12,
            pady=12
        )
        self.frame_footer.pack(expand=1, fill="both", side="bottom")

    def base_settings():
        settings = {}
        settings["title"] = ''
        settings["window_size"] = '960x540'
        settings["font_label_h1"] = ['MSゴシック', 15, 'bold']
        settings["font_label_h2"] = ['MSゴシック', 14, 'bold']
        settings["font_label_body"] = ['MSゴシック', 12]
        settings['font_radio'] = ['MSゴシック', 12]
        settings['font_button'] = ['MSゴシック', 12]
        settings['font_text'] = ['MSゴシック', 15]
        settings['font_listbox'] = ['MSゴシック', 12]
        settings['font_combobox'] = ['MSゴシック', 12]
        settings['font_dateentry'] = ['MSゴシック', 12]
        settings['font_cal'] = ['MSゴシック', 12]
        # ファイル入力処理
        settings["filetype_contents"] = [('XLSXファイル', '*.xlsx'), ('すべて', '*')]
        #   ウィンドウ設定
        settings["bg_color"] = '#f5f5f5'
        #   ラベル設定
        settings["fb_color_ok"] = '#00008b'
        settings["fb_color_ng"] = '#ff0000'
        settings["fb_color_normal"] = '#000000'

        return settings

    def start(self):
        self.master.resizable(False, False)
        self.master.mainloop()

    def exit(self):
        self.master.exit()
