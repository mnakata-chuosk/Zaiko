''' <summary>
''' 在庫1行（営業所内の 倉庫 × 得意先 × 商品CD × ステータス 単位）
''' </summary>
Public Class StockRow
    ''' <summary>商品CD（AXIS では型式単位で一意）</summary>
    Public Property ItemCode As String
    ''' <summary>型式</summary>
    Public Property Model As String
    ''' <summary>倉庫区分</summary>
    Public Property Warehouse As String
    ''' <summary>得意先CD（AXIS 内部コード。棚番の得意先指定との照合に使う）</summary>
    Public Property CustomerCode As String
    ''' <summary>得意先短縮CD（表示用）</summary>
    Public Property CustomerShortCode As String
    ''' <summary>標準仕入先の短縮CD（表示用）</summary>
    Public Property SupplierShortCode As String
    ''' <summary>ステータス区分（通常は空文字）</summary>
    Public Property StatusCode As String
    ''' <summary>ステータス名称</summary>
    Public Property StatusName As String
    ''' <summary>棚卸表（現場記入）に載せるか。s在庫表ステータス 未登録は True</summary>
    Public Property IsCountTarget As Boolean
    ''' <summary>帳簿在庫数（出力時点）</summary>
    Public Property Quantity As Decimal
    ''' <summary>在庫評価単価（最終仕入単価、無ければ現在庫の円貨単価の最大値）</summary>
    Public Property UnitPrice As Decimal
    ''' <summary>単価を現在庫の円貨単価で補完したか</summary>
    Public Property IsPriceSupplemented As Boolean
End Class

''' <summary>
''' 棚番1件（m商品棚番明細）
''' </summary>
Public Class ShelfRow
    Public Property ItemCode As String
    ''' <summary>得意先CD。空文字は得意先指定なし（共用棚）</summary>
    Public Property CustomerCode As String
    Public Property ShelfNo As String
    ''' <summary>帳票出力優先順（DB の値）</summary>
    Public Property Priority As Decimal?
End Class
