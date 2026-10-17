''' <summary>
''' 棚卸表の集計単位（ブロック）。1ブロック＝集計対象の1行（＋内訳行）。
''' ・共用棚ブロック … 得意先指定のない棚。専用棚を持たない得意先の在庫をまとめる
''' ・専用棚ブロック … 得意先専用棚。その得意先の在庫だけ（共用棚の在庫には含めない）
''' </summary>
Public Class StockBlock
    ''' <summary>通し番号（内訳行も同じ番号）</summary>
    Public Property No As Integer
    ''' <summary>専用棚ブロックか</summary>
    Public Property IsDedicated As Boolean
    ''' <summary>専用棚の得意先CD（共用棚は空）</summary>
    Public Property CustomerCode As String = ""
    ''' <summary>棚番（位置 → 棚番）</summary>
    Public Property Shelves As SortedDictionary(Of Integer, String)
    ''' <summary>このブロックの印刷対象の在庫行（棚で数える。得意先 → 倉庫 → ステータス順）</summary>
    Public ReadOnly Property Rows As New List(Of StockRow)
    ''' <summary>印刷対象外ステータスの在庫行（棚に無く、保管先に確認して一覧シートに数を入れる）</summary>
    Public ReadOnly Property OffShelfRows As New List(Of StockRow)

    ''' <summary>全ステータスの在庫行（印刷対象 → 印刷対象外の順）</summary>
    Public ReadOnly Property AllRows As IEnumerable(Of StockRow)
        Get
            Return Rows.Concat(OffShelfRows)
        End Get
    End Property

    ''' <summary>印刷用シート・棚番別シートに載せるか（印刷対象の在庫があるか）</summary>
    Public ReadOnly Property HasPrintRows As Boolean
        Get
            Return Rows.Count > 0
        End Get
    End Property

    ''' <summary>表示に使う在庫行（印刷対象があればそれ、無ければ印刷対象外）</summary>
    Private ReadOnly Property DisplayRows As List(Of StockRow)
        Get
            Return If(Rows.Count > 0, Rows, OffShelfRows)
        End Get
    End Property

    ''' <summary>代表の在庫行（短縮CDが最も若い得意先）</summary>
    Public ReadOnly Property Representative As StockRow
        Get
            Return DisplayRows(0)
        End Get
    End Property

    ''' <summary>得意先の数（登録数。棚を共用している得意先の数なので印刷対象の在庫で数える）</summary>
    Public ReadOnly Property CustomerCount As Integer
        Get
            Return DisplayRows.Select(Function(r) r.CustomerCode).Distinct().Count()
        End Get
    End Property

    ''' <summary>印刷対象の在庫数（棚にあるはずの数）</summary>
    Public ReadOnly Property Quantity As Decimal
        Get
            Return Rows.Sum(Function(r) r.Quantity)
        End Get
    End Property

    ''' <summary>全ステータスの在庫数（一覧シートの AX在庫数）</summary>
    Public ReadOnly Property TotalQuantity As Decimal
        Get
            Return AllRows.Sum(Function(r) r.Quantity)
        End Get
    End Property

    ''' <summary>単価（内訳の単価の最大値）</summary>
    Public ReadOnly Property UnitPrice As Decimal
        Get
            Return AllRows.Max(Function(r) r.UnitPrice)
        End Get
    End Property

    ''' <summary>表示行で値が同じならその値、異なれば代表行の値（useRepresentative=False なら空文字）</summary>
    Public Function CommonValue(selector As Func(Of StockRow, String), Optional useRepresentative As Boolean = False) As String
        Return CommonValueOf(DisplayRows, selector, If(useRepresentative, selector(Representative), ""))
    End Function

    ''' <summary>全ステータスの行で値が同じならその値、異なれば空文字</summary>
    Public Function CommonValueAll(selector As Func(Of StockRow, String)) As String
        Return CommonValueOf(AllRows, selector, "")
    End Function

    Friend Shared Function CommonValueOf(rows As IEnumerable(Of StockRow), selector As Func(Of StockRow, String), otherwise As String) As String
        Dim values = rows.Select(selector).Distinct().ToList()
        Return If(values.Count = 1, values(0), otherwise)
    End Function
End Class

''' <summary>
''' 棚卸表の1型式分（営業所内の商品CD単位）
''' </summary>
Public Class StocktakingItem

    Public Property ItemCode As String
    Public Property Model As String

    ''' <summary>印刷対象の在庫行（得意先 → 倉庫 → ステータス順）</summary>
    Public ReadOnly Property Targets As New List(Of StockRow)
    ''' <summary>印刷対象外ステータスの在庫行（集計行には合算し、保管先に確認した数を一覧シートに入れる）</summary>
    Public ReadOnly Property Excluded As New List(Of StockRow)
    ''' <summary>得意先指定なしの棚番（位置 → 棚番。位置は帳票出力優先順、空き位置は詰めない）</summary>
    Public ReadOnly Property SharedShelves As New SortedDictionary(Of Integer, String)
    ''' <summary>得意先指定ありの棚番（キー: 得意先CD、値: 位置 → 棚番）</summary>
    Public ReadOnly Property CustomerShelves As New Dictionary(Of String, SortedDictionary(Of Integer, String))

    ''' <summary>集計単位（共用棚 → 専用棚の順）</summary>
    Public ReadOnly Property Blocks As New List(Of StockBlock)

    ''' <summary>
    ''' 並び替え用の仕入先（印刷対象の在庫行で最小の仕入先短縮CD。対象が無ければ印刷対象外から）
    ''' </summary>
    Public ReadOnly Property SortSupplier As String
        Get
            Dim src As IEnumerable(Of StockRow) = If(Targets.Count > 0, Targets, Excluded)
            Return src.Select(Function(r) r.SupplierShortCode).OrderBy(Function(s) s, StringComparer.Ordinal).FirstOrDefault()
        End Get
    End Property

    ''' <summary>
    ''' 在庫行を集計単位に分ける。専用棚（棚番あり）を持つ得意先はその得意先だけの専用棚ブロック、
    ''' それ以外の得意先は共用棚ブロックにまとめる。印刷対象外の在庫も同じ規則で得意先のブロックに入れる。
    ''' </summary>
    Public Sub BuildBlocks()
        Blocks.Clear()
        Dim dedicated = Targets.Concat(Excluded).Select(Function(r) r.CustomerCode).Distinct() _
                               .Where(Function(c) CustomerShelves.ContainsKey(c) AndAlso CustomerShelves(c).Count > 0) _
                               .ToList()

        Dim sharedBlock As New StockBlock With {.IsDedicated = False, .Shelves = SharedShelves}
        sharedBlock.Rows.AddRange(Targets.Where(Function(r) Not dedicated.Contains(r.CustomerCode)))
        sharedBlock.OffShelfRows.AddRange(Excluded.Where(Function(r) Not dedicated.Contains(r.CustomerCode)))
        If sharedBlock.AllRows.Any() Then Blocks.Add(sharedBlock)

        For Each code In dedicated
            Dim b As New StockBlock With {.IsDedicated = True, .CustomerCode = code, .Shelves = CustomerShelves(code)}
            b.Rows.AddRange(Targets.Where(Function(r) r.CustomerCode = code))
            b.OffShelfRows.AddRange(Excluded.Where(Function(r) r.CustomerCode = code))
            Blocks.Add(b)
        Next
    End Sub

End Class

''' <summary>
''' 在庫行と棚番から棚卸表の型式一覧を組み立てる
''' </summary>
Public Module StocktakingBuilder

    ''' <summary>
    ''' 商品CD単位にまとめて集計単位に分け、仕入先 → 型式の順に並べて集計単位ごとに通し番号を振る
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

            item.BuildBlocks()
            items.Add(item)
        Next

        ' 仕入先なし（送料・諸口等）は末尾
        Dim sorted = items.OrderBy(Function(i) If(String.IsNullOrEmpty(i.SortSupplier), 1, 0)) _
                          .ThenBy(Function(i) i.SortSupplier, StringComparer.Ordinal) _
                          .ThenBy(Function(i) i.Model, StringComparer.Ordinal) _
                          .ThenBy(Function(i) i.ItemCode, StringComparer.Ordinal) _
                          .ToList()

        Dim no As Integer = 0
        For Each item In sorted
            For Each b In item.Blocks
                no += 1
                b.No = no
            Next
        Next

        Return sorted
    End Function

End Module
