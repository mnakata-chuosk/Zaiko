''' <summary>
''' 棚卸表の出力オプション
''' </summary>
Public Class StocktakingOptions
    ''' <summary>選択できる [棚番・数量] の組数の範囲</summary>
    Public Const MIN_SHELF_SLOTS As Integer = 3
    Public Const MAX_SHELF_SLOTS As Integer = 5
    ''' <summary>画面の初期値</summary>
    Public Const DEFAULT_SHELF_SLOTS As Integer = 3

    ''' <summary>営業所コード（"08" 等）</summary>
    Public Property OfficeCode As String
    ''' <summary>一覧シート1行目・シート名に使う表記（"08.小山営業所" 等）</summary>
    Public Property OfficeLabel As String
    ''' <summary>印刷用シートのヘッダーに使う表記（"08：小山営業所" 等）</summary>
    Public Property OfficeHeader As String
    ''' <summary>[棚番・数量] の組数（一覧・印刷用で共通）</summary>
    Public Property ShelfSlots As Integer = DEFAULT_SHELF_SLOTS
    ''' <summary>印刷用シートに帳簿在庫数の列を出すか</summary>
    Public Property ShowStockOnPrintSheet As Boolean
    ''' <summary>出力日時（在庫の基準時点として表示）</summary>
    Public Property OutputAt As DateTime
End Class
