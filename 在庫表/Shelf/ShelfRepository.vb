Imports System.Text
Imports ChuoUtils

''' <summary>
''' 棚番（AXIS.dbo.m商品棚番明細）の取得・更新
''' </summary>
''' <remarks>
''' m商品棚番明細 はシステムバージョン管理テーブル（履歴は m商品棚番明細_Log に自動記録）。
''' 更新は 商品CD × 営業所 × 得意先指定 の単位で1トランザクションとし、
''' 読み込み時点から件数・SysStartTime が変わっていれば AXIS 側で編集されたとみなして更新しない。
''' </remarks>
Public NotInheritable Class ShelfRepository

    Private Sub New()
    End Sub

    ''' <summary>保存結果</summary>
    Public Enum SaveOutcome
        Saved
        ''' <summary>読み込み後に他で更新されていたため保存しなかった</summary>
        Conflict
    End Enum

    Private Const SHELF_SQL As String = "
SELECT b.商品CD, b.商品棚番明細番号, ISNULL(b.得意先CD, '') AS 得意先CD, ISNULL(b.棚番, '') AS 棚番, b.帳票出力優先順, b.SysStartTime,
       ISNULL(m.型式, '') AS 型式, ISNULL(t.短縮CD, '') AS 短縮CD, ISNULL(t.略称, ISNULL(t.取引先名, '')) AS 得意先名
FROM AXIS.dbo.m商品棚番明細 b
LEFT JOIN AXIS.dbo.m商品管理 m ON m.商品CD = b.商品CD
LEFT JOIN AXIS.dbo.M取引先管理 t ON t.取引先CD = b.得意先CD
WHERE b.営業所区分 = @OFFICE
ORDER BY b.商品CD, ISNULL(b.得意先CD, ''), ISNULL(b.帳票出力優先順, 999999), b.商品棚番明細番号"

    Private Const STOCK_SQL As String = "
SELECT z.商品CD, ISNULL(z.得意先CD, '') AS 得意先CD, SUM(z.在庫数) AS 在庫数,
       ISNULL(m.型式, '') AS 型式, ISNULL(t.短縮CD, '') AS 短縮CD, ISNULL(t.略称, ISNULL(t.取引先名, '')) AS 得意先名
FROM AXIS.dbo.FT_在庫数明細取得(GETDATE(), '1') z
INNER JOIN AXIS.dbo.m商品管理 m ON m.商品CD = z.商品CD AND m.在庫管理区分 = '1'
LEFT JOIN AXIS.dbo.M取引先管理 t ON t.取引先CD = z.得意先CD
WHERE z.営業所区分 = @OFFICE
GROUP BY z.商品CD, ISNULL(z.得意先CD, ''), m.型式, t.短縮CD, t.略称, t.取引先名
HAVING SUM(z.在庫数) <> 0"

    Private Const CUSTOMER_SQL As String = "
SELECT h.得意先CD, ISNULL(t.短縮CD, '') AS 短縮CD, ISNULL(t.略称, ISNULL(t.取引先名, '')) AS 得意先名
FROM AXIS.dbo.m商品取扱得意先明細 h
LEFT JOIN AXIS.dbo.M取引先管理 t ON t.取引先CD = h.得意先CD
WHERE h.商品CD = @ITEM AND h.営業所区分 = @OFFICE
ORDER BY t.短縮CD"

    Private Const MODEL_SQL As String = "
SELECT 商品CD, ISNULL(型式, '') AS 型式
FROM AXIS.dbo.m商品管理
WHERE 型式 = @MODEL AND ISNULL(使用可否区分, '0') = '0'"

    ''' <summary>得意先（取扱得意先明細）</summary>
    Public Class Customer
        Public Property Code As String
        Public Property ShortCode As String
        Public Property Name As String
    End Class

    ''' <summary>
    ''' 営業所の棚番を 商品CD × 得意先指定 の単位で取得する。
    ''' 在庫があるのに共用棚の行が無い型式は、棚番を追加できるよう空の共用棚の行を加える。
    ''' </summary>
    Public Shared Function LoadGroups(officeCode As String) As List(Of ShelfGroup)
        Dim prm As New Dictionary(Of String, Object) From {{"@OFFICE", officeCode}}
        Dim groups As New Dictionary(Of String, ShelfGroup)

        For Each row As DataRow In Query(SHELF_SQL, prm).Rows
            Dim g = GetOrAdd(groups, Str(row("商品CD")), Str(row("得意先CD")), Str(row("型式")), Str(row("短縮CD")), Str(row("得意先名")))
            g.Records.Add(New ShelfRecord With {
                .DetailNo = CInt(row("商品棚番明細番号")),
                .ShelfNo = Str(row("棚番")),
                .Priority = If(TypeOf row("帳票出力優先順") Is DBNull, CType(Nothing, Decimal?), Convert.ToDecimal(row("帳票出力優先順")))
            })
            g.LoadedCount += 1
            Dim stamp = CDate(row("SysStartTime"))
            If Not g.LoadedStamp.HasValue OrElse stamp > g.LoadedStamp.Value Then g.LoadedStamp = stamp
        Next

        For Each g In groups.Values
            g.AssignPositions()
        Next

        For Each row As DataRow In Query(STOCK_SQL, prm).Rows
            Dim itemCode = Str(row("商品CD"))
            Dim qty = Convert.ToDecimal(row("在庫数"))
            Dim model = Str(row("型式"))

            ' 共用棚の行は型式の在庫合計
            GetOrAdd(groups, itemCode, "", model, "", "").StockQuantity += qty

            ' 得意先指定の行がある得意先はその在庫も表示
            Dim key = ShelfGroup.MakeKey(itemCode, Str(row("得意先CD")))
            If key <> ShelfGroup.MakeKey(itemCode, "") AndAlso groups.ContainsKey(key) Then
                groups(key).StockQuantity += qty
            End If
        Next

        Return groups.Values _
            .OrderBy(Function(g) g.Model, StringComparer.Ordinal) _
            .ThenBy(Function(g) g.ItemCode, StringComparer.Ordinal) _
            .ThenBy(Function(g) g.CustomerShortCode, StringComparer.Ordinal) _
            .ToList()
    End Function

    Private Shared Function GetOrAdd(groups As Dictionary(Of String, ShelfGroup), itemCode As String, customerCode As String,
                                     model As String, shortCode As String, customerName As String) As ShelfGroup
        Dim key = ShelfGroup.MakeKey(itemCode, customerCode)
        Dim g As ShelfGroup = Nothing
        If Not groups.TryGetValue(key, g) Then
            g = New ShelfGroup With {
                .ItemCode = itemCode, .Model = model,
                .CustomerCode = customerCode, .CustomerShortCode = shortCode, .CustomerName = customerName
            }
            groups.Add(key, g)
        End If
        Return g
    End Function

    ''' <summary>型式・営業所の取扱得意先</summary>
    Public Shared Function LoadCustomers(officeCode As String, itemCode As String) As List(Of Customer)
        Dim prm As New Dictionary(Of String, Object) From {{"@OFFICE", officeCode}, {"@ITEM", itemCode}}
        Return Query(CUSTOMER_SQL, prm).Rows.Cast(Of DataRow)() _
            .Select(Function(r) New Customer With {.Code = Str(r("得意先CD")), .ShortCode = Str(r("短縮CD")), .Name = Str(r("得意先名"))}) _
            .ToList()
    End Function

    ''' <summary>型式（完全一致）から使用中の商品を探す。キーは商品CD、値は型式</summary>
    Public Shared Function FindItemsByModel(model As String) As Dictionary(Of String, String)
        Dim prm As New Dictionary(Of String, Object) From {{"@MODEL", model}}
        Return Query(MODEL_SQL, prm).Rows.Cast(Of DataRow)() _
            .ToDictionary(Function(r) Str(r("商品CD")), Function(r) Str(r("型式")))
    End Function

    ''' <summary>
    ''' 1グループ分の棚番を保存する（PendingMap の内容に置き換える）。
    ''' 位置ごとに、既存明細があれば棚番・帳票出力優先順（=位置）を書き換え、無ければ追加（明細番号は商品CD内の最大＋1）、
    ''' 棚番が無くなった位置の明細は削除する。
    ''' </summary>
    Public Shared Function Save(officeCode As String, userId As String, group As ShelfGroup) As SaveOutcome
        Dim cmd = BuildSaveCommand(officeCode, userId, group)
        Dim dt = Query(cmd.Sql, cmd.Params)
        Return If(dt.Rows.Count > 0 AndAlso CInt(dt.Rows(0)("結果")) = 0, SaveOutcome.Saved, SaveOutcome.Conflict)
    End Function

    ''' <summary>保存用の SQL バッチとパラメータを組み立てる（結果セット「結果」: 0=保存、1=競合）</summary>
    Public Shared Function BuildSaveCommand(officeCode As String, userId As String, group As ShelfGroup) As (Sql As String, Params As Dictionary(Of String, Object))
        Dim newMap = group.PendingMap
        Dim sql As New StringBuilder()
        Dim prm As New Dictionary(Of String, Object) From {
            {"@OFFICE", officeCode},
            {"@ITEM", group.ItemCode},
            {"@CUST", group.CustomerCode},
            {"@CUSTNULL", If(group.IsShared, CObj(DBNull.Value), group.CustomerCode)},
            {"@USER", userId},
            {"@EXP_CNT", group.LoadedCount},
            {"@EXP_STAMP", If(group.LoadedStamp.HasValue, CObj(group.LoadedStamp.Value), DBNull.Value)}
        }

        sql.AppendLine("SET NOCOUNT ON; SET XACT_ABORT ON;")
        sql.AppendLine("BEGIN TRAN;")
        sql.AppendLine("DECLARE @cnt INT, @stamp DATETIME2;")
        sql.AppendLine("SELECT @cnt = COUNT(*), @stamp = MAX(SysStartTime) FROM AXIS.dbo.m商品棚番明細 WITH (UPDLOCK, HOLDLOCK)")
        sql.AppendLine("  WHERE 商品CD = @ITEM AND 営業所区分 = @OFFICE AND ISNULL(得意先CD, '') = @CUST;")
        sql.AppendLine("IF @cnt <> @EXP_CNT OR ISNULL(@stamp, '19000101') <> ISNULL(@EXP_STAMP, '19000101')")
        sql.AppendLine("BEGIN ROLLBACK; SELECT 1 AS 結果; RETURN; END")

        Dim byPosition = group.Records.ToDictionary(Function(r) r.Position)
        Dim positions = byPosition.Keys.Union(newMap.Keys).OrderBy(Function(p) p).ToList()
        For Each pos In positions
            Dim rec As ShelfRecord = Nothing
            byPosition.TryGetValue(pos, rec)
            Dim shelf As String = Nothing
            newMap.TryGetValue(pos, shelf)

            If shelf IsNot Nothing Then
                prm.Add($"@S{pos}", shelf)
                prm.Add($"@P{pos}", CDec(pos))
            End If
            If rec IsNot Nothing Then prm.Add($"@D{pos}", rec.DetailNo)

            If rec IsNot Nothing AndAlso shelf IsNot Nothing Then
                ' 既存明細の書き換え（変化がある行のみ）
                sql.AppendLine($"UPDATE AXIS.dbo.m商品棚番明細 SET 棚番 = @S{pos}, 帳票出力優先順 = @P{pos}, 更新担当者区分 = @USER, 更新日時 = GETDATE()")
                sql.AppendLine($"  WHERE 商品CD = @ITEM AND 商品棚番明細番号 = @D{pos} AND (ISNULL(棚番, '') <> @S{pos} OR ISNULL(帳票出力優先順, -1) <> @P{pos});")

            ElseIf shelf IsNot Nothing Then
                ' 追加
                sql.AppendLine("INSERT INTO AXIS.dbo.m商品棚番明細 (商品CD, 商品棚番明細番号, 営業所区分, 得意先CD, 棚番, 帳票出力優先順, 登録担当者区分, 登録日時, 更新担当者区分, 更新日時)")
                sql.AppendLine($"  SELECT @ITEM, ISNULL(MAX(商品棚番明細番号), 0) + 1, @OFFICE, @CUSTNULL, @S{pos}, @P{pos}, @USER, GETDATE(), @USER, GETDATE()")
                sql.AppendLine("  FROM AXIS.dbo.m商品棚番明細 WITH (UPDLOCK, HOLDLOCK) WHERE 商品CD = @ITEM;")

            Else
                ' 削除（この位置の棚番が無くなった）
                sql.AppendLine($"DELETE FROM AXIS.dbo.m商品棚番明細 WHERE 商品CD = @ITEM AND 商品棚番明細番号 = @D{pos};")
            End If
        Next
        sql.AppendLine("COMMIT;")
        sql.AppendLine("SELECT 0 AS 結果;")

        Return (sql.ToString(), prm)
    End Function

    ''' <summary>パラメータ付きで実行し、最初の結果セットを返す（例外は呼び出し元へ）</summary>
    Private Shared Function Query(sql As String, prm As Dictionary(Of String, Object)) As DataTable
        Using sobj As New SqlObject(SqlObject.DbName.AXIS)
            sobj.SetCommandText(SqlObject.CMD_SELECT, sql)
            For Each p In prm
                ' 比較用の SysStartTime は NULL でも datetime2 として渡す
                Dim dbType = If(p.Key = "@EXP_STAMP", SqlDbType.DateTime2, ToDbType(p.Value))
                sobj.AddParameter(SqlObject.CMD_SELECT, p.Key, dbType, p.Value)
            Next
            ' GetDataTable は例外を MsgBox で握りつぶすため Fill を使う
            sobj.Fill()
            Dim ds As DataSet = sobj.GetDataSet()
            Return If(ds.Tables.Count > 0, ds.Tables(0).Copy(), New DataTable)
        End Using
    End Function

    Private Shared Function ToDbType(value As Object) As SqlDbType
        If TypeOf value Is Integer Then Return SqlDbType.Int
        If TypeOf value Is Decimal Then Return SqlDbType.Decimal
        If TypeOf value Is DateTime Then Return SqlDbType.DateTime2
        Return SqlDbType.NVarChar
    End Function

    Private Shared Function Str(value As Object) As String
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return ""
        Return value.ToString().Trim()
    End Function

End Class
