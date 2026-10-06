Imports System.Text.RegularExpressions

''' <summary>
''' Excel 読み込みで変わる1グループ分の棚番
''' </summary>
Public Class ShelfImportChange
    Public Property Group As ShelfGroup
    ''' <summary>編集画面にまだ無いグループ（新しく行を追加する）</summary>
    Public Property IsNewGroup As Boolean
    ''' <summary>位置 → 棚番</summary>
    Public Property Before As SortedDictionary(Of Integer, String)
    Public Property After As SortedDictionary(Of Integer, String)
    ''' <summary>確認を促す内容（メモらしき棚番など）</summary>
    Public Property Warnings As New List(Of String)
    Public Property ExcelRows As New SortedSet(Of Integer)
End Class

''' <summary>
''' 棚卸表から読み取った棚番を、編集画面の現在の内容と突き合わせて変更点を求める。
''' 内容が同じグループは結果に含めない。棚番は位置（棚番1～）で扱い、空欄は「その位置に棚番なし」。
''' </summary>
''' <remarks>
''' ・親行 … 共用棚。ファイルの棚番列（位置 1～列数）をそのまま反映。列数より後ろの位置は変更しない
''' ・子行 … その得意先の指定棚。親行と同じ
''' ・単独行 … 共用棚と得意先指定棚を合成して出力しているため（ShelfPositions.MergeSingle）、
'''            出力時と同じ合成をやり直し、変わったセルは元のグループ・元の位置に反映する。
'''            出力時に空いていたセルへの記入は共用棚のその位置に入れる
''' ・単独行で棚区分が「専用」… その行の棚番をその得意先の専用棚にする（同じ位置の共用棚は外す）
''' ・追記行 … 型式（完全一致）で特定し、空いている位置に棚番を追加する。棚区分が「専用」なら得意先短縮CDの専用棚
''' </remarks>
Public Class ShelfImportPlanner

    Private ReadOnly _officeCode As String
    Private ReadOnly _groups As Dictionary(Of String, ShelfGroup)
    Private ReadOnly _work As New Dictionary(Of String, SortedDictionary(Of Integer, String))
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
                        Overwrite(GetGroup(row.ItemCode, "", row), row)

                    Case ShelfImportRow.RowKind.Child
                        Dim cust = ResolveCustomer(row.ItemCode, row)
                        If cust Is Nothing Then Continue For
                        ' 指定棚が無く、ファイルにも書かれていない子行は何もしない
                        If Not row.HasShelf AndAlso Not Exists(ShelfGroup.MakeKey(row.ItemCode, cust.Code)) Then Continue For
                        Overwrite(GetGroup(row.ItemCode, cust.Code, row, cust), row)

                    Case ShelfImportRow.RowKind.SingleRow
                        If row.IsDedicated Then
                            PlanSingleDedicated(row)
                        Else
                            PlanSingle(row)
                        End If

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

            Dim before = g.CurrentMap
            If ShelfPositions.SameMap(kv.Value, before) Then Continue For

            Dim change As New ShelfImportChange With {
                .Group = g, .IsNewGroup = isNew, .Before = before, .After = kv.Value
            }
            For Each r In _rowsByKey(kv.Key)
                change.ExcelRows.Add(r)
            Next
            For Each s In kv.Value.Values.Where(Function(x) Not before.ContainsValue(x) AndAlso LooksLikeMemo(x)).Distinct()
                change.Warnings.Add($"「{s}」は棚番ではなくメモの可能性")
            Next
            If kv.Value.Keys.Any(Function(p) p > ShelfGroup.EDIT_SLOTS) AndAlso Not before.Keys.Any(Function(p) p > ShelfGroup.EDIT_SLOTS) Then
                change.Warnings.Add($"棚番{ShelfGroup.EDIT_SLOTS + 1}以降に入る棚番あり（画面には{ShelfGroup.EDIT_SLOTS}まで表示）")
            End If
            result.Add(change)
        Next

        Return result.OrderBy(Function(c) c.ExcelRows.Min).ToList()
    End Function

    ''' <summary>ファイルの棚番列（位置 1～列数）で上書きする。空欄はその位置を削除</summary>
    Private Sub Overwrite(g As ShelfGroup, row As ShelfImportRow)
        Dim map As New SortedDictionary(Of Integer, String)(Current(g.Key))
        For i As Integer = 0 To row.Shelves.Count - 1
            Dim pos = i + 1
            If row.Shelves(i) = "" Then
                map.Remove(pos)
            Else
                map(pos) = row.Shelves(i)
            End If
        Next
        SetWork(g, map, row.ExcelRow)
    End Sub

    ''' <summary>単独行：出力時の合成をやり直し、変わったセルを元のグループ・元の位置へ反映する</summary>
    Private Sub PlanSingle(row As ShelfImportRow)
        Dim sharedGroup = GetGroup(row.ItemCode, "", row)
        Dim sharedMap As New SortedDictionary(Of Integer, String)(Current(sharedGroup.Key))

        Dim cust = If(row.CustomerShortCode = "" AndAlso row.CustomerCode = "", Nothing, ResolveCustomer(row.ItemCode, row, quiet:=True))
        Dim custKey = If(cust Is Nothing, Nothing, ShelfGroup.MakeKey(row.ItemCode, cust.Code))
        Dim hasCustGroup = custKey IsNot Nothing AndAlso Exists(custKey)
        Dim custMap As New SortedDictionary(Of Integer, String)(If(hasCustGroup, Current(custKey), New SortedDictionary(Of Integer, String)))

        Dim merged = ShelfPositions.MergeSingle(sharedMap, If(hasCustGroup, custMap, Nothing))
        Dim newShared As New SortedDictionary(Of Integer, String)(sharedMap)
        Dim newCust As New SortedDictionary(Of Integer, String)(custMap)

        For i As Integer = 0 To row.Shelves.Count - 1
            Dim pos = i + 1
            Dim v = row.Shelves(i)
            Dim m As ShelfPositions.MergedShelf = Nothing
            If merged.TryGetValue(pos, m) Then
                If v = m.Shelf Then Continue For
                Dim target = If(m.IsCustomer, newCust, newShared)
                If v = "" Then
                    target.Remove(m.SourcePosition)
                Else
                    target(m.SourcePosition) = v
                End If
            ElseIf v <> "" Then
                newShared(pos) = v
            End If
        Next

        SetWork(sharedGroup, newShared, row.ExcelRow)
        If hasCustGroup Then SetWork(GetGroup(row.ItemCode, cust.Code, row, cust), newCust, row.ExcelRow)
    End Sub

    ''' <summary>
    ''' 単独行で棚区分が「専用」：この行の棚番（位置 1～列数）をその得意先の専用棚にする。
    ''' 単独行の棚番は共用棚として出力しているため、同じ位置の共用棚は外す（専用棚へ移す）。
    ''' </summary>
    Private Sub PlanSingleDedicated(row As ShelfImportRow)
        Dim cust = ResolveCustomer(row.ItemCode, row)
        If cust Is Nothing Then Return

        Dim sharedGroup = GetGroup(row.ItemCode, "", row)
        Dim newShared As New SortedDictionary(Of Integer, String)(Current(sharedGroup.Key))
        For pos As Integer = 1 To row.Shelves.Count
            newShared.Remove(pos)
        Next
        SetWork(sharedGroup, newShared, row.ExcelRow)

        Overwrite(GetGroup(row.ItemCode, cust.Code, row, cust), row)
    End Sub

    ''' <summary>追記行：型式と得意先短縮CDから特定し、空いている位置に棚番を追加する</summary>
    Private Sub PlanAppend(row As ShelfImportRow)
        If row.Model = "" Then
            Errors.Add($"{row.ExcelRow}行目：追記行に品名（型式）がありません")
            Return
        End If
        If Not row.HasShelf Then
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

        ' 棚区分の列があれば「専用」のときだけ得意先専用棚、無い古いファイルは得意先が書いてあれば専用棚
        Dim dedicated = If(row.ShelfKind Is Nothing, row.CustomerShortCode <> "", row.IsDedicated)
        If dedicated AndAlso row.CustomerShortCode = "" Then
            Errors.Add($"{row.ExcelRow}行目：追記行「{row.Model}」は棚区分が専用ですが得意先がありません")
            Return
        End If

        Dim cust As ShelfRepository.Customer = Nothing
        If dedicated Then
            cust = ResolveCustomer(itemCode, row)
            If cust Is Nothing Then Return
        End If

        Dim g = GetGroup(itemCode, If(cust?.Code, ""), row, cust)
        Dim map As New SortedDictionary(Of Integer, String)(Current(g.Key))
        For Each s In row.Shelves.Where(Function(x) x <> "")
            If map.ContainsValue(s) Then Continue For
            Dim pos As Integer = 1
            While map.ContainsKey(pos)
                pos += 1
            End While
            map(pos) = s
        Next
        SetWork(g, map, row.ExcelRow)
    End Sub

    Private Sub SetWork(g As ShelfGroup, map As SortedDictionary(Of Integer, String), excelRow As Integer)
        _work(g.Key) = map
        If Not _rowsByKey.ContainsKey(g.Key) Then _rowsByKey(g.Key) = New SortedSet(Of Integer)
        _rowsByKey(g.Key).Add(excelRow)
    End Sub

    Private Function Exists(key As String) As Boolean
        Return _groups.ContainsKey(key) OrElse _newGroups.ContainsKey(key)
    End Function

    ''' <summary>このファイルの処理途中の内容（未処理なら編集画面の現在の内容）</summary>
    Private Function Current(key As String) As SortedDictionary(Of Integer, String)
        Dim map As SortedDictionary(Of Integer, String) = Nothing
        If _work.TryGetValue(key, map) Then Return map
        Dim g As ShelfGroup = Nothing
        If _groups.TryGetValue(key, g) OrElse _newGroups.TryGetValue(key, g) Then Return g.CurrentMap
        Return New SortedDictionary(Of Integer, String)
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
