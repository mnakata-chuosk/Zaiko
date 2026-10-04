Imports System.Text.RegularExpressions

''' <summary>
''' Excel 読み込みで変わる1グループ分の棚番
''' </summary>
Public Class ShelfImportChange
    Public Property Group As ShelfGroup
    ''' <summary>編集画面にまだ無いグループ（新しく行を追加する）</summary>
    Public Property IsNewGroup As Boolean
    Public Property Before As List(Of String)
    Public Property After As List(Of String)
    ''' <summary>確認を促す内容（メモらしき棚番など）</summary>
    Public Property Warnings As New List(Of String)
    Public Property ExcelRows As New SortedSet(Of Integer)
End Class

''' <summary>
''' 棚卸表から読み取った棚番を、編集画面の現在の内容と突き合わせて変更点を求める。
''' 内容が同じグループは結果に含めない。
''' </summary>
''' <remarks>
''' ・親行 … 共用棚。ファイル上の棚番列数より後ろの登録済み棚番はそのまま残す
''' ・子行 … その得意先の指定棚
''' ・単独行 … 共用棚と得意先指定棚が混在して出力されているため、
'''            元々その得意先の指定棚だった棚番は指定棚に、それ以外（新しく書かれた棚番を含む）は共用棚に振り分ける
''' ・追記行 … 型式（完全一致）と得意先短縮CD（その型式の取扱得意先）から特定し、棚番を追加する
''' </remarks>
Public Class ShelfImportPlanner

    Private ReadOnly _officeCode As String
    Private ReadOnly _groups As Dictionary(Of String, ShelfGroup)
    Private ReadOnly _work As New Dictionary(Of String, List(Of String))
    Private ReadOnly _newGroups As New Dictionary(Of String, ShelfGroup)
    Private ReadOnly _rowsByKey As New Dictionary(Of String, SortedSet(Of Integer))
    Private ReadOnly _customerCache As New Dictionary(Of String, List(Of ShelfRepository.Customer))
    Private ReadOnly _modelCache As New Dictionary(Of String, Dictionary(Of String, String))

    ''' <summary>反映できなかった行（ファイルの行番号付きメッセージ）</summary>
    Public ReadOnly Property Errors As New List(Of String)

    Public Sub New(officeCode As String, groups As IEnumerable(Of ShelfGroup))
        _officeCode = officeCode
        _groups = groups.ToDictionary(Function(g) g.Key)
    End Sub

    Public Function Plan(file As ShelfImportFile) As List(Of ShelfImportChange)
        For Each row In file.Rows
            Try
                Select Case row.Kind
                    Case ShelfImportRow.RowKind.Parent
                        Replace(GetGroup(row.ItemCode, "", row), row.Shelves, file.Slots, row.ExcelRow)

                    Case ShelfImportRow.RowKind.Child
                        Dim cust = ResolveCustomer(row.ItemCode, row)
                        If cust Is Nothing Then Continue For
                        Dim key = ShelfGroup.MakeKey(row.ItemCode, cust.Code)
                        ' 指定棚が無く、ファイルにも書かれていない子行は何もしない
                        If row.Shelves.Count = 0 AndAlso Not Exists(key) Then Continue For
                        Replace(GetGroup(row.ItemCode, cust.Code, row, cust), row.Shelves, file.Slots, row.ExcelRow)

                    Case ShelfImportRow.RowKind.SingleRow
                        PlanSingle(row, file.Slots)

                    Case ShelfImportRow.RowKind.Append
                        PlanAppend(row)
                End Select
            Catch ex As Exception
                Errors.Add($"{row.ExcelRow}行目：{ex.Message}")
            End Try
        Next

        Dim result As New List(Of ShelfImportChange)
        For Each kv In _work
            Dim g As ShelfGroup = Nothing
            Dim isNew = Not _groups.TryGetValue(kv.Key, g)
            If isNew Then g = _newGroups(kv.Key)

            Dim before = g.CurrentShelves
            If kv.Value.SequenceEqual(before) Then Continue For

            Dim change As New ShelfImportChange With {
                .Group = g, .IsNewGroup = isNew, .Before = before, .After = kv.Value
            }
            For Each r In _rowsByKey(kv.Key)
                change.ExcelRows.Add(r)
            Next
            For Each s In kv.Value.Where(Function(x) Not before.Contains(x) AndAlso LooksLikeMemo(x)).Distinct()
                change.Warnings.Add($"「{s}」は棚番ではなくメモの可能性")
            Next
            If kv.Value.Count > ShelfGroup.EDIT_SLOTS Then
                change.Warnings.Add($"棚番が{kv.Value.Count}件（画面には{ShelfGroup.EDIT_SLOTS}件まで表示）")
            End If
            result.Add(change)
        Next

        Return result.OrderBy(Function(c) c.ExcelRows.Min).ToList()
    End Function

    ''' <summary>単独行：元の指定棚は指定棚に、それ以外は共用棚に振り分ける</summary>
    Private Sub PlanSingle(row As ShelfImportRow, slots As Integer)
        Dim sharedKey = ShelfGroup.MakeKey(row.ItemCode, "")
        Dim cust = If(row.CustomerShortCode = "" AndAlso row.CustomerCode = "", Nothing, ResolveCustomer(row.ItemCode, row, quiet:=True))
        Dim custKey = If(cust Is Nothing, Nothing, ShelfGroup.MakeKey(row.ItemCode, cust.Code))

        Dim sharedNow = Current(sharedKey)
        Dim custNow = If(custKey IsNot Nothing AndAlso Exists(custKey), Current(custKey), New List(Of String))

        ' 出力時と同じ並び（共用棚 → 共用棚に無い指定棚）
        Dim combined = sharedNow.Concat(custNow.Where(Function(s) Not sharedNow.Contains(s))).ToList()
        Dim proposed = row.Shelves.Concat(combined.Skip(slots)).Distinct().ToList()

        Dim custNew = proposed.Where(Function(s) custNow.Contains(s) AndAlso Not sharedNow.Contains(s)).ToList()
        Dim sharedNew = proposed.Where(Function(s) Not custNew.Contains(s)).ToList()

        SetWork(GetGroup(row.ItemCode, "", row), sharedNew, row.ExcelRow)
        If custKey IsNot Nothing AndAlso Exists(custKey) Then
            SetWork(GetGroup(row.ItemCode, cust.Code, row, cust), custNew, row.ExcelRow)
        End If
    End Sub

    ''' <summary>追記行：型式と得意先短縮CDから特定して棚番を追加する</summary>
    Private Sub PlanAppend(row As ShelfImportRow)
        If row.Model = "" Then
            Errors.Add($"{row.ExcelRow}行目：追記行に品名（型式）がありません")
            Return
        End If
        If row.Shelves.Count = 0 Then
            Errors.Add($"{row.ExcelRow}行目：追記行「{row.Model}」に棚番がありません")
            Return
        End If

        Dim items As Dictionary(Of String, String) = Nothing
        If Not _modelCache.TryGetValue(row.Model, items) Then
            items = ShelfRepository.FindItemsByModel(row.Model)
            _modelCache(row.Model) = items
        End If
        If items.Count = 0 Then
            Errors.Add($"{row.ExcelRow}行目：型式「{row.Model}」が商品マスタに見つかりません")
            Return
        End If
        If items.Count > 1 Then
            Errors.Add($"{row.ExcelRow}行目：型式「{row.Model}」に該当する商品が複数あります（{String.Join("、", items.Keys)}）。編集画面で行追加してください")
            Return
        End If

        Dim itemCode = items.Keys.First()
        row.ItemCode = itemCode
        row.Model = items(itemCode)

        Dim cust As ShelfRepository.Customer = Nothing
        If row.CustomerShortCode <> "" Then
            cust = ResolveCustomer(itemCode, row)
            If cust Is Nothing Then Return
        End If

        Dim g = GetGroup(itemCode, If(cust?.Code, ""), row, cust)
        Dim list = Current(g.Key)
        SetWork(g, list.Concat(row.Shelves.Where(Function(s) Not list.Contains(s))).ToList(), row.ExcelRow)
    End Sub

    ''' <summary>先頭 slots 件をファイルの内容で置き換え、それ以降の登録済み棚番は残す</summary>
    Private Sub Replace(g As ShelfGroup, shelves As List(Of String), slots As Integer, excelRow As Integer)
        Dim proposed = shelves.Concat(Current(g.Key).Skip(slots)).Distinct().ToList()
        SetWork(g, proposed, excelRow)
    End Sub

    Private Sub SetWork(g As ShelfGroup, shelves As List(Of String), excelRow As Integer)
        _work(g.Key) = shelves
        If Not _rowsByKey.ContainsKey(g.Key) Then _rowsByKey(g.Key) = New SortedSet(Of Integer)
        _rowsByKey(g.Key).Add(excelRow)
    End Sub

    Private Function Exists(key As String) As Boolean
        Return _groups.ContainsKey(key) OrElse _newGroups.ContainsKey(key)
    End Function

    ''' <summary>このファイルの処理途中の内容（未処理なら編集画面の現在の内容）</summary>
    Private Function Current(key As String) As List(Of String)
        Dim list As List(Of String) = Nothing
        If _work.TryGetValue(key, list) Then Return list
        Dim g As ShelfGroup = Nothing
        If _groups.TryGetValue(key, g) OrElse _newGroups.TryGetValue(key, g) Then Return g.CurrentShelves
        Return New List(Of String)
    End Function

    ''' <summary>編集画面のグループ。無ければ新しいグループを用意する</summary>
    Private Function GetGroup(itemCode As String, customerCode As String, row As ShelfImportRow,
                              Optional cust As ShelfRepository.Customer = Nothing) As ShelfGroup
        Dim key = ShelfGroup.MakeKey(itemCode, customerCode)
        Dim g As ShelfGroup = Nothing
        If _groups.TryGetValue(key, g) OrElse _newGroups.TryGetValue(key, g) Then Return g

        g = New ShelfGroup With {
            .ItemCode = itemCode,
            .Model = row.Model,
            .CustomerCode = customerCode,
            .CustomerShortCode = If(cust?.ShortCode, ""),
            .CustomerName = If(cust?.Name, "")
        }
        _newGroups.Add(key, g)
        Return g
    End Function

    ''' <summary>
    ''' 行の得意先を特定する。非表示列の得意先CDがあればそれを、無ければ型式の取扱得意先から短縮CDで探す
    ''' </summary>
    Private Function ResolveCustomer(itemCode As String, row As ShelfImportRow, Optional quiet As Boolean = False) As ShelfRepository.Customer
        Dim customers As List(Of ShelfRepository.Customer) = Nothing
        If Not _customerCache.TryGetValue(itemCode, customers) Then
            customers = ShelfRepository.LoadCustomers(_officeCode, itemCode)
            _customerCache(itemCode) = customers
        End If

        Dim found As List(Of ShelfRepository.Customer)
        If row.CustomerCode <> "" Then
            found = customers.Where(Function(c) c.Code = row.CustomerCode).ToList()
            If found.Count = 0 Then
                ' 取扱得意先から外れていても、行の得意先CDはそのまま使う
                found.Add(New ShelfRepository.Customer With {.Code = row.CustomerCode, .ShortCode = row.CustomerShortCode})
            End If
        Else
            found = customers.Where(Function(c) c.ShortCode = row.CustomerShortCode).ToList()
        End If

        If found.Count = 1 Then Return found(0)
        If Not quiet Then
            Errors.Add($"{row.ExcelRow}行目：「{row.Model}」の得意先「{row.CustomerShortCode}」を特定できません" &
                       If(found.Count > 1, "（同じ短縮CDが複数）", "（取扱得意先にありません）"))
        End If
        Return Nothing
    End Function

    Private Shared ReadOnly MEMO_PATTERN As New Regex("済|納入|予定|返品|\d{1,2}/\d{1,2}")

    ''' <summary>棚番ではなくメモが書かれていそうか</summary>
    Private Shared Function LooksLikeMemo(shelf As String) As Boolean
        Return shelf.Length > 12 OrElse MEMO_PATTERN.IsMatch(shelf)
    End Function

End Class
