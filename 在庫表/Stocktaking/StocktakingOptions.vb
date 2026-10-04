''' <summary>
''' 棚卸表の出力オプション
''' </summary>
Public Class StocktakingOptions
    ''' <summary>営業所コード（"08" 等）</summary>
    Public Property OfficeCode As String
    ''' <summary>一覧シート1行目・シート名に使う表記（"08.小山営業所" 等）</summary>
    Public Property OfficeLabel As String
    ''' <summary>印刷用シートのヘッダーに使う表記（"08：小山営業所" 等）</summary>
    Public Property OfficeHeader As String
    ''' <summary>印刷用シートに帳簿在庫数の列を出すか</summary>
    Public Property ShowStockOnPrintSheet As Boolean
    ''' <summary>出力日時（在庫の基準時点として表示）</summary>
    Public Property OutputAt As DateTime
End Class
