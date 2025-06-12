# pyinstaller -n PyTkinterToPSScript --add-data "classes/config/settings.xml;classes/config" --add-data "classes/image/logo.png;classes/image" --hidden-import babel.numbers --clean --onefile --noconsole --collect-data tkinterdnd2 --paths="C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript-UserDefined" main.py
pyinstaller -n PyTkinterToPSScript --add-data "classes/image/logo.png;classes/image" --hidden-import babel.numbers --clean --onefile --noconsole --collect-data tkinterdnd2 --paths="C:\Users\Administrator\Documents\Git\python\PyTkinterToPSScript-UserDefined" main.py
# たぶんリリース後に更新する場合
# pyinstaller .\PyTkinterToPSScript.spec

# 作成した実行ファイルをコピー
$copyFrom = ".\dist\PyTkinterToPSScript.exe"
$copyTo = ".\classes\exe\PyTkinterToPSScript.exe"
Copy-Item -Path $copyFrom -Destination $copyTo -Force
