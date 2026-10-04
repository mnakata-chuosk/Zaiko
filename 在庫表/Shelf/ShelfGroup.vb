''' <summary>
''' m商品棚番明細 の1件
''' </summary>
Public Class ShelfRecord
    ''' <summary>商品棚番明細番号（商品CD内の連番。営業所をまたいで共通）</summary>
    Public Property DetailNo As Integer
    Public Property ShelfNo As String
End Class

''' <summary>
''' 棚番の編集単位（営業所内の 商品CD × 得意先指定）。
''' 得意先指定なし（共用棚）は CustomerCode = ""。
''' </summary>
Public Class ShelfGroup

    ''' <summary>画面・Excel で横並びに編集できる棚番の数</summary>
    Public Const EDIT_SLOTS As Integer = 5

    Public Property ItemCode As String
    Public Property Model As String
    ''' <summary>得意先CD（AXIS 内部コード）。共用棚は空文字</summary>
    Public Property CustomerCode As String = ""
    Public Property CustomerShortCode As String = ""
    Public Property CustomerName As String = ""
    ''' <summary>在庫数（共用棚は型式の合計、得意先指定はその得意先分）</summary>
    Public Property StockQuantity As Decimal

    ''' <summary>読み込み時点の明細（帳票出力優先順 → 明細番号の順）</summary>
    Public ReadOnly Property Records As New List(Of ShelfRecord)

    ''' <summary>読み込み時点の件数と SysStartTime の最大値（同時編集の検出用）</summary>
    Public Property LoadedCount As Integer
    Public Property LoadedStamp As DateTime?

    ''' <summary>編集後の棚番（Nothing は未変更）。6件目以降も含む全件</summary>
    Public Property PendingShelves As List(Of String)

    Public ReadOnly Property Key As String
        Get
            Return MakeKey(ItemCode, CustomerCode)
        End Get
    End Property

    Public Shared Function MakeKey(itemCode As String, customerCode As String) As String
        Return itemCode & "|" & customerCode
    End Function

    Public ReadOnly Property IsShared As Boolean
        Get
            Return CustomerCode = ""
        End Get
    End Property

    ''' <summary>読み込み時点の棚番（全件）</summary>
    Public ReadOnly Property OriginalShelves As List(Of String)
        Get
            Return Records.Select(Function(r) r.ShelfNo).ToList()
        End Get
    End Property

    ''' <summary>現在の棚番（編集中なら編集後、未編集なら読み込み時点）</summary>
    Public ReadOnly Property CurrentShelves As List(Of String)
        Get
            Return If(PendingShelves, OriginalShelves)
        End Get
    End Property

    ''' <summary>編集で内容が変わったか</summary>
    Public ReadOnly Property IsChanged As Boolean
        Get
            Return PendingShelves IsNot Nothing AndAlso Not PendingShelves.SequenceEqual(OriginalShelves)
        End Get
    End Property

    ''' <summary>
    ''' 横並び5枠の編集結果を反映する。6件目以降は現在の内容をそのまま残す。空欄は詰める。
    ''' </summary>
    Public Sub ApplySlots(slots As IEnumerable(Of String))
        Dim edited = slots.Select(Function(s) If(s, "").Trim()).Where(Function(s) s <> "").ToList()
        edited.AddRange(CurrentShelves.Skip(EDIT_SLOTS))
        PendingShelves = edited
    End Sub

End Class
