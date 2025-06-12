import tkinter as tk
import tkinterdnd2 as tkdnd2
import ctypes

# Own module
from classes.form import formPackageMain
from classes.control import ctrlMessage

if __name__ == '__main__':
    try:
        # 多重起動の確認（Win限定）
        kernel32 = ctypes.windll.Kernel32
        mutex = kernel32.CreateMutexA(0, 1, "mutex-unique-name")
        _result = kernel32.WaitForSingleObject(mutex, 0)

        if _result == 0:
            # master = tk.Tk()
            master = tkdnd2.Tk()
            main = formPackageMain(master)
            main.start()
        else:
            main = tk.Tk()
            main.withdraw()
            main.lift()
            ctrlMessage.view_top_warning(-9001)

        # # 多重起動の確認なしにする場合
        #     master = tkdnd2.TkinterDnD.Tk()
        #     main = formPackageMain(master)
        #     main.start()
    
    except Exception as err:
        ctrlMessage.print_error(err)
        input("Press Enter to exit...")
