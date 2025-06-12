# カスタムクラス定義
class PrintListConfig {
    [System.String]$Path
    [System.Object[]]$HeaderConstants
    [System.Object[]]$HeaderValues
    [System.Object[]]$MainConstants
    [System.Object[]]$MainValues

    # 定数（クラス内で管理）
    static [System.String]$TEMPLATESHEET1 = '1ページ目'
    static [System.Int32]$TEMPLATEROWS1 = 30
    static [System.String]$TEMPLATESHEET2 = '2ページ目以降'
    static [System.Int32]$TEMPLATEROWS2 = 35
    static [System.String]$HEADERRANGE = 'B2:AA9'
    static [System.String]$MAINRANGE = 'B7:AA41'

    PrintListConfig([System.String]$Path, [System.Object[]]$HeaderConstants, [System.Object[]]$HeaderValues, [System.Object[]]$MainConstants, [System.Object[]]$MainValues) {
        $this.Path = $Path
        $this.HeaderConstants = $HeaderConstants
        $this.HeaderValues = $HeaderValues
        $this.MainConstants = $MainConstants
        $this.MainValues = $MainValues
    }
}
