''' <summary>
''' 棚卸表の1型式分（営業所内の商品CD単位）。
''' 棚卸対象の在庫行が2行以上なら「親行＋子行」、1行なら単独行として出力する。
''' </summary>
Public Class StocktakingItem

    ''' <summary>通し番号（一覧シート・印刷用シート共通）</summary>
    Public Property No As Integer
    Public Property ItemCode As String
    Public Property Model As String

    ''' <summary>棚卸対象の在庫行（得意先 → 倉庫 → ステータス順）</summary>
    Public ReadOnly Property Targets As New List(Of StockRow)
    ''' <summary>棚卸対象外ステータスの在庫行（一覧シートのみ参考表示）</summary>
    Public ReadOnly Property Excluded As New List(Of StockRow)
    ''' <summary>得意先指定なしの棚番（位置 → 棚番。位置は帳票出力優先順、空き位置は詰めない）</summary>
    Public ReadOnly Property SharedShelves As New SortedDictionary(Of Integer, String)
    ''' <summary>得意先指定ありの棚番（キー: 得意先CD、値: 位置 → 棚番）</summary>
    Public ReadOnly Property CustomerShelves As New Dictionary(Of String, SortedDictionary(Of Integer, String))

    ''' <summary>
    ''' 親行＋子行で出力するか。棚卸対象が2行以上、または1行でも得意先指定棚がある場合
    ''' （共用棚＝親行、得意先指定棚＝子行で区別できるようにするため）
    ''' </summary>
    Public ReadOnly Property IsGroup As Boolean
        Get
            Return Targets.Count >= 2 OrElse (Targets.Count = 1 AndAlso ShelvesOf(Targets(0)).Count > 0)
        End Get
    End Property

    ''' <summary>棚卸対象の在庫行の合計</summary>
    Public ReadOnly Property TargetQuantity As Decimal
        Get
            Return Targets.Sum(Function(r) r.Quantity)
        End Get
    End Property

    ''' <summary>
    ''' 並び替え用の仕入先（棚卸対象の在庫行で最小の仕入先短縮CD。対象が無ければ全行から）
    ''' </summary>
    Public ReadOnly Property SortSupplier As String
        Get
            Dim src As IEnumerable(Of StockRow) = If(Targets.Count > 0, Targets, Excluded)
            Return src.Select(Function(r) r.SupplierShortCode).OrderBy(Function(s) s, StringComparer.Ordinal).FirstOrDefault()
        End Get
    End Property

    ''' <summary>全行で値が同じならその値、異なれば空文字</summary>
    Public Function CommonValue(selector As Func(Of StockRow, String)) As String
        Dim values = Targets.Select(selector).Distinct().ToList()
        Return If(values.Count = 1, values(0), "")
    End Function

    ''' <summary>得意先指定の棚番（無ければ空）</summary>
    Public Function ShelvesOf(row As StockRow) As SortedDictionary(Of Integer, String)
        Dim map As SortedDictionary(Of Integer, String) = Nothing
        If CustomerShelves.TryGetValue(row.CustomerCode, map) Then Return map
        Return New SortedDictionary(Of Integer, String)
    End Function

    ''' <summary>単独行で出力するときの棚番（共用棚と得意先指定棚を ShelfPositions.MergeSingle で合成）</summary>
    Public Function SingleRowShelves() As SortedDictionary(Of Integer, String)
        Dim custMap = If(Targets.Count = 1 AndAlso CustomerShelves.ContainsKey(Targets(0).CustomerCode), CustomerShelves(Targets(0).CustomerCode), Nothing)
        Dim result As New SortedDictionary(Of Integer, String)
        For Each kv In ShelfPositions.MergeSingle(SharedShelves, custMap)
            result(kv.Key) = kv.Value.Shelf
        Next
        Return result
    End Function

End Class

''' <summary>
''' 在庫行と棚番から棚卸表の型式一覧を組み立てる
''' </summary>
Public Module StocktakingBuilder

    ''' <summary>
    ''' 商品CD単位にまとめ、仕入先 → 型式の順に並べて通し番号を振る
    ''' </summary>
    Public Function Build(rows As List(Of StockRow), shelves As List(Of ShelfRow)) As List(Of StocktakingItem)
        Dim shelfLookup = shelves.ToLookup(Function(s) s.ItemCode)

        Dim items As New List(Of StocktakingItem)
        For Each grp In rows.GroupBy(Function(r) r.ItemCode)
            Dim item As New StocktakingItem With {
                .ItemCode = grp.Key,
                .Model = grp.First().Model
            }

            Dim ordered = grp.OrderBy(Function(r) r.CustomerShortCode, StringComparer.Ordinal) _
                             .ThenBy(Function(r) r.Warehouse, StringComparer.Ordinal) _
                             .ThenBy(Function(r) r.StatusCode, StringComparer.Ordinal)
            For Each r In ordered
                If r.IsCountTarget Then
                    item.Targets.Add(r)
                Else
                    item.Excluded.Add(r)
                End If
            Next

            ' 得意先指定ごとに帳票出力優先順から位置を決める（棚番編集画面と同じ規則）
            For Each byCustomer In shelfLookup(grp.Key).GroupBy(Function(s) s.CustomerCode)
                Dim list = byCustomer.ToList()
                Dim positions = ShelfPositions.Assign(list.Select(Function(s) s.Priority).ToList())
                Dim map = If(byCustomer.Key = "", item.SharedShelves, New SortedDictionary(Of Integer, String))
                For i As Integer = 0 To list.Count - 1
                    map(positions(i)) = list(i).ShelfNo
                Next
                If byCustomer.Key <> "" Then item.CustomerShelves(byCustomer.Key) = map
            Next

            items.Add(item)
        Next

        ' 仕入先なし（送料・諸口等）は末尾
        Dim sorted = items.OrderBy(Function(i) If(String.IsNullOrEmpty(i.SortSupplier), 1, 0)) _
                          .ThenBy(Function(i) i.SortSupplier, StringComparer.Ordinal) _
                          .ThenBy(Function(i) i.Model, StringComparer.Ordinal) _
                          .ThenBy(Function(i) i.ItemCode, StringComparer.Ordinal) _
                          .ToList()

        For i As Integer = 0 To sorted.Count - 1
            sorted(i).No = i + 1
        Next

        Return sorted
    End Function

End Module
