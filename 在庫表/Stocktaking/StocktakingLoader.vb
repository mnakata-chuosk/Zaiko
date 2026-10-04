Imports ChuoUtils

''' <summary>
''' 棚卸表用データを AXIS から取得する
''' </summary>
Public NotInheritable Class StocktakingLoader

    Private Sub New()
    End Sub

    ''' <summary>
    ''' 在庫の取得。
    ''' ・FT_在庫数明細取得(GETDATE(),'1') を 倉庫×得意先×商品CD×ステータス で合算（通貨・単価違いの行をまとめる）
    ''' ・在庫数 0 と在庫管理区分 '1' 以外は除外
    ''' ・単価は最新締の t締在庫実績明細.最終仕入単価（円）。NULL/0 の場合は現在庫（在庫数&gt;0 の単価行）の円貨単価の最大値
    ''' ・棚卸表出力区分は APP DB の s在庫表ステータス。未登録のステータスは '1'（載せる）
    ''' </summary>
    Private Const STOCK_SQL As String = "
SET NOCOUNT ON;
IF OBJECT_ID('tempdb..#在庫') IS NOT NULL DROP TABLE #在庫;

SELECT 商品CD, ISNULL(倉庫区分, '') AS 倉庫区分, ISNULL(得意先CD, '') AS 得意先CD, ISNULL(ステータス区分, '') AS ステータス区分,
       在庫数, 円貨単価
INTO #在庫
FROM AXIS.dbo.FT_在庫数明細取得(GETDATE(), '1')
WHERE 営業所区分 = @OFFICE;

WITH 在庫 AS (
    SELECT 商品CD, 倉庫区分, 得意先CD, ステータス区分, SUM(在庫数) AS 在庫数
    FROM #在庫
    GROUP BY 商品CD, 倉庫区分, 得意先CD, ステータス区分
    HAVING SUM(在庫数) <> 0
), 補完単価 AS (
    SELECT 商品CD, 得意先CD, ステータス区分, MAX(円貨単価) AS 単価
    FROM #在庫
    WHERE 在庫数 > 0
    GROUP BY 商品CD, 得意先CD, ステータス区分
)
SELECT z.商品CD, z.倉庫区分, z.得意先CD, z.ステータス区分, z.在庫数,
       ISNULL(m.型式, '') AS 型式,
       ISNULL(tk.短縮CD, '') AS 得意先短縮CD,
       ISNULL(sh.短縮CD, '') AS 仕入先短縮CD,
       ISNULL(st.名称, '') AS ステータス名,
       ISNULL(sf.棚卸表出力区分, '1') AS 棚卸表出力区分,
       CASE WHEN ISNULL(j.最終仕入単価, 0) <> 0 THEN j.最終仕入単価 ELSE ISNULL(p.単価, 0) END AS 単価,
       CASE WHEN ISNULL(j.最終仕入単価, 0) <> 0 THEN 0 ELSE 1 END AS 単価補完
FROM 在庫 z
INNER JOIN AXIS.dbo.m商品管理 m
        ON m.商品CD = z.商品CD AND m.在庫管理区分 = '1'
LEFT JOIN AXIS.dbo.m商品取扱得意先明細 h
       ON h.商品CD = z.商品CD AND h.得意先CD = z.得意先CD AND h.営業所区分 = @OFFICE
LEFT JOIN AXIS.dbo.M取引先管理 tk ON tk.取引先CD = z.得意先CD
LEFT JOIN AXIS.dbo.M取引先管理 sh ON sh.取引先CD = h.標準仕入先CD
LEFT JOIN AXIS.dbo.vkステータス区分 st ON st.区分 = z.ステータス区分
LEFT JOIN アプリ.dbo.s在庫表ステータス sf ON sf.ステータス区分 = z.ステータス区分
LEFT JOIN AXIS.dbo.t締在庫実績明細 j
       ON j.締CD = (SELECT MAX(締CD) FROM AXIS.dbo.t締管理)
      AND j.商品CD = z.商品CD AND j.営業所区分 = @OFFICE AND j.得意先CD = z.得意先CD
      AND ISNULL(j.ステータス区分, '') = z.ステータス区分
LEFT JOIN 補完単価 p
       ON p.商品CD = z.商品CD AND p.得意先CD = z.得意先CD AND p.ステータス区分 = z.ステータス区分;

DROP TABLE #在庫;"

    ''' <summary>棚番の取得（帳票出力優先順 → 明細番号の順）</summary>
    Private Const SHELF_SQL As String = "
SELECT 商品CD, ISNULL(得意先CD, '') AS 得意先CD, 棚番
FROM AXIS.dbo.m商品棚番明細
WHERE 営業所区分 = @OFFICE AND ISNULL(棚番, '') <> ''
ORDER BY 商品CD, ISNULL(帳票出力優先順, 999999), 商品棚番明細番号"

    ''' <summary>営業所の在庫を取得する</summary>
    Public Shared Function LoadStock(officeCode As String) As List(Of StockRow)
        Dim result As New List(Of StockRow)

        For Each row As DataRow In Query(STOCK_SQL, officeCode).Rows
            result.Add(New StockRow With {
                .ItemCode = Str(row("商品CD")),
                .Model = Str(row("型式")),
                .Warehouse = Str(row("倉庫区分")),
                .CustomerCode = Str(row("得意先CD")),
                .CustomerShortCode = Str(row("得意先短縮CD")),
                .SupplierShortCode = Str(row("仕入先短縮CD")),
                .StatusCode = Str(row("ステータス区分")),
                .StatusName = Str(row("ステータス名")),
                .IsCountTarget = Str(row("棚卸表出力区分")) = "1",
                .Quantity = Dec(row("在庫数")),
                .UnitPrice = Dec(row("単価")),
                .IsPriceSupplemented = Dec(row("単価補完")) = 1D
            })
        Next

        Return result
    End Function

    ''' <summary>営業所の棚番を取得する</summary>
    Public Shared Function LoadShelves(officeCode As String) As List(Of ShelfRow)
        Dim result As New List(Of ShelfRow)

        For Each row As DataRow In Query(SHELF_SQL, officeCode).Rows
            result.Add(New ShelfRow With {
                .ItemCode = Str(row("商品CD")),
                .CustomerCode = Str(row("得意先CD")),
                .ShelfNo = Str(row("棚番"))
            })
        Next

        Return result
    End Function

    ''' <summary>営業所コードを @OFFICE に渡して SELECT を実行する（例外は呼び出し元へ）</summary>
    Private Shared Function Query(sql As String, officeCode As String) As DataTable
        Using sobj As New SqlObject(SqlObject.DbName.AXIS)
            sobj.SetCommandText(SqlObject.CMD_SELECT, sql)
            sobj.AddParameter(SqlObject.CMD_SELECT, "@OFFICE", SqlDbType.NVarChar, officeCode)
            ' GetDataTable は例外を MsgBox で握りつぶすため Fill を使う
            sobj.Fill()
            Dim ds As DataSet = sobj.GetDataSet()
            Return If(ds.Tables.Count > 0, ds.Tables(0).Copy(), New DataTable)
        End Using
    End Function

    Private Shared Function Str(value As Object) As String
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return ""
        Return value.ToString().Trim()
    End Function

    Private Shared Function Dec(value As Object) As Decimal
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return 0D
        Return Convert.ToDecimal(value)
    End Function

End Class
