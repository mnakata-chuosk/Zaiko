''' <summary>
''' m商品棚番明細 の1件
''' </summary>
Public Class ShelfRecord
    ''' <summary>商品棚番明細番号（商品CD内の連番。営業所をまたいで共通）</summary>
    Public Property DetailNo As Integer
    Public Property ShelfNo As String
    ''' <summary>帳票出力優先順（DB の値）</summary>
    Public Property Priority As Decimal?
    ''' <summary>棚番の位置（ShelfPositions.Assign で決めた 1～）</summary>
    Public Property Position As Integer
End Class

''' <summary>
''' 棚番の編集単位（営業所内の 商品CD × 得意先指定）。
''' 得意先指定なし（共用棚）は CustomerCode = ""。
''' 棚番は位置（帳票出力優先順）をキーに持ち、空き位置は詰めない。
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

    ''' <summary>読み込み時点の明細（位置の割り当て済み）</summary>
    Public ReadOnly Property Records As New List(Of ShelfRecord)

    ''' <summary>読み込み時点の件数と SysStartTime の最大値（同時編集の検出用）</summary>
    Public Property LoadedCount As Integer
    Public Property LoadedStamp As DateTime?

    ''' <summary>編集後の棚番（位置 → 棚番。Nothing は未変更）。6 以降の位置も含む</summary>
    Public Property PendingMap As SortedDictionary(Of Integer, String)

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

    ''' <summary>明細を追加し終えたら呼ぶ。帳票出力優先順から位置を決める</summary>
    Public Sub AssignPositions()
        Dim positions = ShelfPositions.Assign(Records.Select(Function(r) r.Priority).ToList())
        For i As Integer = 0 To Records.Count - 1
            Records(i).Position = positions(i)
        Next
    End Sub

    ''' <summary>読み込み時点の棚番（位置 → 棚番。空の棚番は含めない）</summary>
    Public ReadOnly Property OriginalMap As SortedDictionary(Of Integer, String)
        Get
            Dim map As New SortedDictionary(Of Integer, String)
            For Each r In Records
                If r.ShelfNo <> "" Then map(r.Position) = r.ShelfNo
            Next
            Return map
        End Get
    End Property

    ''' <summary>現在の棚番（編集中なら編集後、未編集なら読み込み時点）</summary>
    Public ReadOnly Property CurrentMap As SortedDictionary(Of Integer, String)
        Get
            Return If(PendingMap, OriginalMap)
        End Get
    End Property

    ''' <summary>編集で内容が変わったか</summary>
    Public ReadOnly Property IsChanged As Boolean
        Get
            Return PendingMap IsNot Nothing AndAlso Not ShelfPositions.SameMap(PendingMap, OriginalMap)
        End Get
    End Property

    ''' <summary>位置 6 以降の棚番の件数</summary>
    Public ReadOnly Property ExtraCount As Integer
        Get
            Return CurrentMap.Keys.Where(Function(p) p > EDIT_SLOTS).Count()
        End Get
    End Property

    ''' <summary>
    ''' 棚番1～5 の編集結果を反映する（空欄はその位置の棚番を削除。詰めない）。位置 6 以降はそのまま残す。
    ''' </summary>
    Public Sub ApplySlots(slots As IList(Of String))
        Dim map As New SortedDictionary(Of Integer, String)
        For Each kv In CurrentMap.Where(Function(x) x.Key > EDIT_SLOTS)
            map(kv.Key) = kv.Value
        Next
        For i As Integer = 0 To Math.Min(slots.Count, EDIT_SLOTS) - 1
            Dim v = If(slots(i), "").Trim()
            If v <> "" Then map(i + 1) = v
        Next
        PendingMap = map
    End Sub

End Class
