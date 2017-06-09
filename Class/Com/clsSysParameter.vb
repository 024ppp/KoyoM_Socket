Public Class clsSysParameter
    'ＤＢ処理の追加
    Public Shared gInsert As String = "0"
    'ＤＢ処理の更新
    Public Shared gUpdate As String = "1"
    'ＤＢ処理の削除
    Public Shared gDelete As String = "2"
    '履歴
    Public Shared gRireki As String = "3"
    'クリア
    Public Shared gClear As String = "1"
    '実行
    Public Shared gRun As String = "2"
    '検索
    Public Shared gKensaku As String = "3"
    'エクセル出力
    Public Shared gExcel As String = "4"
    '再表示
    Public Shared gRefresh As String = "5"
    'PDF
    Public Shared gPDF As String = "6"
    '四捨五入
    Public Shared gSisyagonyuu As String = "1"
    '切り捨て
    Public Shared gKirisute As String = "2"
    '切り上げ
    Public Shared gKiriage As String = "3"
    'エクセルテンプレートフォルダ名
    Public Shared gTEMP_FD As String = System.IO.Directory.GetParent(Application.StartupPath).ToString & "\Templates"

    '担当者ID
    Public Shared gTanCD As String = ""
    '担当者名
    Public Shared gTanNM As String = ""
    '担当者かな
    Public Shared gTanKN As String = ""
    '普通式
    Public Shared gFutuSiki As String = "(W/(T*B*L))*1000"
End Class
