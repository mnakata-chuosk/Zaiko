''' <summary>
''' 棚番の位置（棚番1～5 の列）の扱い。帳票出力優先順をそのまま位置として使い、空き位置は詰めない。
''' （「棚番3は外部倉庫」のように位置に意味を持たせる運用があるため）
''' </summary>
Public Module ShelfPositions

    ''' <summary>単独行に合成した棚番1件</summary>
    Public Structure MergedShelf
        Public Shelf As String
        ''' <summary>得意先指定棚から来たか（False は共用棚）</summary>
        Public IsCustomer As Boolean
        ''' <summary>元のグループでの位置</summary>
        Public SourcePosition As Integer
    End Structure

    ''' <summary>
    ''' 明細の並び（帳票出力優先順 → 明細番号の順）に対して位置を割り当てる。
    ''' 優先順が 1 以上の整数で未使用ならその位置、NULL・重複・小数は空いている最小の位置に入れる。
    ''' </summary>
    Public Function Assign(priorities As IList(Of Decimal?)) As Integer()
        Dim result(priorities.Count - 1) As Integer
        Dim taken As New HashSet(Of Integer)

        For i As Integer = 0 To priorities.Count - 1
            Dim p = priorities(i)
            If p.HasValue AndAlso p.Value >= 1D AndAlso p.Value = Decimal.Truncate(p.Value) AndAlso p.Value <= Integer.MaxValue Then
                Dim pos = CInt(p.Value)
                If taken.Add(pos) Then result(i) = pos
            End If
        Next

        For i As Integer = 0 To priorities.Count - 1
            If result(i) <> 0 Then Continue For
            Dim pos As Integer = 1
            While taken.Contains(pos)
                pos += 1
            End While
            taken.Add(pos)
            result(i) = pos
        Next

        Return result
    End Function

    ''' <summary>
    ''' 単独行（共用棚と得意先指定棚を1行に出す）の棚番を合成する。
    ''' 共用棚は元の位置、得意先指定棚は元の位置が空いていればそこ、埋まっていれば空いている最小の位置に置く。
    ''' </summary>
    Public Function MergeSingle(sharedMap As IDictionary(Of Integer, String),
                                customerMap As IDictionary(Of Integer, String)) As SortedDictionary(Of Integer, MergedShelf)
        Dim result As New SortedDictionary(Of Integer, MergedShelf)
        For Each kv In sharedMap
            result(kv.Key) = New MergedShelf With {.Shelf = kv.Value, .IsCustomer = False, .SourcePosition = kv.Key}
        Next
        If customerMap Is Nothing Then Return result

        For Each kv In customerMap
            If result.Values.Any(Function(m) m.Shelf = kv.Value) Then Continue For   ' 共用棚と同じ棚番は重ねて出さない
            Dim pos As Integer = kv.Key
            If result.ContainsKey(pos) Then
                pos = 1
                While result.ContainsKey(pos)
                    pos += 1
                End While
            End If
            result(pos) = New MergedShelf With {.Shelf = kv.Value, .IsCustomer = True, .SourcePosition = kv.Key}
        Next
        Return result
    End Function

    ''' <summary>位置付きの棚番を「A / （空）/ C」のように表示する</summary>
    Public Function Describe(map As IDictionary(Of Integer, String)) As String
        If map.Count = 0 Then Return "（なし）"
        Dim last = map.Keys.Max()
        Return String.Join(" / ", Enumerable.Range(1, last).Select(Function(p) If(map.ContainsKey(p), map(p), "（空）")))
    End Function

    ''' <summary>位置付きの棚番が同じ内容か</summary>
    Public Function SameMap(a As IDictionary(Of Integer, String), b As IDictionary(Of Integer, String)) As Boolean
        If a.Count <> b.Count Then Return False
        For Each kv In a
            Dim v As String = Nothing
            If Not b.TryGetValue(kv.Key, v) OrElse v <> kv.Value Then Return False
        Next
        Return True
    End Function

End Module
