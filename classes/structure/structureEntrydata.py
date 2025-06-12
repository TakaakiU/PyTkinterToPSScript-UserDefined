from dataclasses import dataclass

# 入力データの構造体
@dataclass
class EntryData:
    folderpath: str
    targetrange: str
    workdate: str
    workername: str
    terminalname: str
