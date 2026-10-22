Option Strict Off  ' PageSetup・Outline 等の Excel COM 遅延バインディングに必要

Imports ChuoUtils
Imports ChuoUtils.ExcelTools

''' <summary>
''' 棚卸表（一覧シート・印刷用シート・棚番別シート）を Excel に出力し、ユーザーに表示する
''' </summary>
''' <remarks>
''' 集計単位（StockBlock）ごとに1行（集計対象）を出す。AX在庫数は全ステータス（印刷対象外を含む）の合算。
''' ・共用棚の行 … 得意先は代表得意先（短縮CDが最も若い得意先）。専用棚を持たない得意先の在庫の合計
''' ・専用棚の行 … 「専用」欄に○。その得意先の在庫だけ（共用棚の在庫には含めない）
''' ・内訳行（灰色・同じNo）… 「得意先内訳を表示」なら得意先・倉庫・ステータス別、しないならステータス別の在庫（「内訳在庫」列）。集計対象外
''' ・印刷対象外の行 … 棚に無いステータスの在庫。棚卸総数の欄に保管先で確認した数を入れ、集計行の棚卸総数に足す
''' 集計用の列（棚卸総数・AX在庫数・差数・単価・差額）は集計対象の行にだけ値を入れ、二重計上を防ぐ。
''' 両シートとも末尾に追記用の空行を設け、リスト外の現物はそこに書き込む。
''' </remarks>
Public NotInheritable Class StocktakingWriter

    ''' <summary>[棚番・数量] の組数（3～5、一覧・印刷用で共通）</summary>
    Private ReadOnly SLOTS As Integer
    ''' <summary>リスト外の現物を書き込むための追記行数</summary>
    Private Const APPEND_ROWS As Integer = 10
    Private Const APPEND_LABEL As String = ListHeaders.APPEND_NO
    ''' <summary>印刷用シート（品物ごと）。棚番別と合わせて現場で書き込む紙</summary>
    Private Const PRINT_SHEET_NAME As String = "型式別"
    Private Const SHELF_SHEET_NAME As String = "棚番別"

    ' 一覧シートの列
    Private Const L_NO As Integer = 1
    Private Const L_WH As Integer = 2
    Private Const L_CUST As Integer = 3
    Private Const L_SUP As Integer = 4
    Private Const L_ITEM As Integer = 5
    Private Const L_NAME As Integer = 6
    Private Const L_STATUS As Integer = 7
    Private Const L_MARK As Integer = 8                         ' 専用（専用棚の行に○。空欄以外は専用として読み込む）
    Private Const L_REG As Integer = 9
    Private Const L_SHELF As Integer = 10                       ' 棚番1（数量1 は +1、以降 2 列おき）
    ' 棚番列より右は組数で位置が変わる
    Private ReadOnly L_TOTAL As Integer                         ' 棚卸総数
    Private ReadOnly L_BOOK As Integer                          ' AX在庫数
    Private ReadOnly L_DETAIL As Integer                        ' 内訳在庫（内訳行のみ）
    Private ReadOnly L_PLUS As Integer                          ' ＋（売上漏れ）
    Private ReadOnly L_MINUS As Integer                         ' －（仕入漏れ）
    Private ReadOnly L_DIFF As Integer                          ' 差数
    Private ReadOnly L_PRICE As Integer                         ' 単価
    Private ReadOnly L_AMOUNT As Integer                        ' 差額
    Private ReadOnly L_NOTE As Integer                          ' 備考
    Private ReadOnly L_CUST_CODE As Integer                     ' 得意先CD（非表示。AXIS内部コード）
    Private ReadOnly L_ROW_TYPE As Integer                      ' 行種別（非表示。共用／専用／内訳／専用内訳／ステータス内訳／印刷対象外）
    Private ReadOnly L_ORDER As Integer                         ' 並び順（非表示。出力時の並びに戻す用）
    Private ReadOnly L_COLS As Integer
    Private Const L_HEADER_ROW As Integer = 2
    Private Const L_FIRST_ROW As Integer = 3

    ' 色（旧アプリの重複明細と同じ灰色）
    Private Shared ReadOnly COLOR_DETAIL As Integer = RGB(217, 217, 217)
    Private Shared ReadOnly COLOR_EXCLUDED_BACK As Integer = RGB(242, 242, 242)
    Private Shared ReadOnly COLOR_EXCLUDED_FONT As Integer = RGB(128, 128, 128)
    Private Shared ReadOnly COLOR_NG As Integer = RGB(255, 199, 206)
    Private Shared ReadOnly COLOR_INPUT As Integer = RGB(255, 242, 204)

    ''' <summary>連続した行範囲</summary>
    Private Structure RowSpan
        Public First As Integer
        Public Last As Integer
        Public Sub New(first As Integer, last As Integer)
            Me.First = first
            Me.Last = last
        End Sub
    End Structure

    Private Sub New(shelfSlots As Integer)
        SLOTS = shelfSlots
        Dim baseCol = L_SHELF + SLOTS * 2
        L_TOTAL = baseCol
        L_BOOK = baseCol + 1
        L_DETAIL = baseCol + 2
        L_PLUS = baseCol + 3
        L_MINUS = baseCol + 4
        L_DIFF = baseCol + 5
        L_PRICE = baseCol + 6
        L_AMOUNT = baseCol + 7
        L_NOTE = baseCol + 8
        L_CUST_CODE = baseCol + 9
        L_ROW_TYPE = baseCol + 10
        L_ORDER = baseCol + 11
        L_COLS = L_ORDER
    End Sub

    ''' <summary>
    ''' 棚卸表を作成して Excel を表示する。表示後の Excel はユーザーに移譲する。
    ''' </summary>
    ''' <param name="savePath">指定時は表示せずにこのパスへ保存して閉じる（一括出力・検証用）</param>
    Public Shared Sub Write(items As List(Of StocktakingItem), opt As StocktakingOptions, Optional savePath As String = Nothing)
        Dim w As New StocktakingWriter(opt.ShelfSlots)
        Dim xl As New ExcelObject()
        Try
            xl.SheetName = opt.OfficeLabel & ListHeaders.SHEET_SUFFIX
            w.WriteListSheet(xl, items, opt)

            xl.AddSheet(PRINT_SHEET_NAME)
            w.WritePrintSheet(xl, items, opt)
            xl.SetFreezePanes(2, 1)

            xl.AddSheet(SHELF_SHEET_NAME)
            w.WriteShelfSheet(xl, items, opt)

            ' Show() は最後に SetFreezePanes したシート（=先頭シート）にだけ固定を再適用するため、一覧シートを最後にする
            xl.SetSheet(1)
            xl.SetFreezePanes(L_FIRST_ROW, L_NAME + 1)

            If String.IsNullOrEmpty(savePath) Then
                xl.Show()
            Else
                xl.Calculate()
                xl.SetHome()
                ' 開いた人の Excel で数式（専用の印による昇格の計算など）が自動で再計算されるように
                xl.Calculation = xlCalculationAutomatic
                xl.SaveAs(savePath)
                xl.Dispose()
            End If
        Catch
            xl.Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 印刷用シートの内訳行（印刷対象の在庫だけ）。
    ''' 得意先内訳を表示するなら得意先 × 倉庫 × ステータス別、しないならステータス別に合算。1行にしかならないときは出さない
    ''' </summary>
    Private Shared Function PrintDetailRows(b As StockBlock, opt As StocktakingOptions) As List(Of StockRow)
        Dim result As List(Of StockRow)
        If opt.ShowBreakdown Then
            result = b.Rows.ToList()
        Else
            result = b.Rows.GroupBy(Function(r) r.StatusCode).OrderBy(Function(g) g.Key, StringComparer.Ordinal).Select(Function(g) MergeRows(g.ToList())).ToList()
        End If
        If result.Count <= 1 Then result.Clear()
        Return result
    End Function

    ''' <summary>一覧シートの内訳行1行分</summary>
    Private Structure DetailEntry
        Public Row As StockRow
        Public RowType As String
        Public Sub New(row As StockRow, rowType As String)
            Me.Row = row
            Me.RowType = rowType
        End Sub
    End Structure

    ''' <summary>
    ''' 一覧シートの集計行の下に出す内訳行。
    ''' ・得意先内訳を表示 … 得意先 × 倉庫 × ステータス別（今までどおり）
    ''' ・表示しない … ステータス別に合算（ステータスが2種類以上のとき）
    ''' 印刷対象外の在庫は数を入れる欄が要るので、内訳の設定にかかわらず必ず行を出す
    ''' </summary>
    Private Shared Function DetailEntries(b As StockBlock, opt As StocktakingOptions) As List(Of DetailEntry)
        Dim result As New List(Of DetailEntry)
        Dim byStatus = Function(rows As List(Of StockRow)) rows.GroupBy(Function(r) r.StatusCode).OrderBy(Function(g) g.Key, StringComparer.Ordinal).Select(Function(g) MergeRows(g.ToList()))

        If opt.ShowBreakdown Then
            Dim printType As String = If(b.IsDedicated, ListHeaders.ROW_DEDICATED_DETAIL, ListHeaders.ROW_DETAIL)
            result.AddRange(b.Rows.Select(Function(r) New DetailEntry(r, printType)))
            result.AddRange(b.OffShelfRows.Select(Function(r) New DetailEntry(r, ListHeaders.ROW_EXCLUDED)))
        Else
            result.AddRange(byStatus(b.Rows).Select(Function(r) New DetailEntry(r, ListHeaders.ROW_STATUS)))
            result.AddRange(byStatus(b.OffShelfRows).Select(Function(r) New DetailEntry(r, ListHeaders.ROW_EXCLUDED)))
        End If

        ' 印刷対象外が無く、内訳が1行だけなら集計行と同じなので出さない
        If result.Count <= 1 AndAlso Not result.Any(Function(e) e.RowType = ListHeaders.ROW_EXCLUDED) Then result.Clear()
        Return result
    End Function

    ''' <summary>同じステータスの在庫行を1行に合算する（倉庫・得意先・仕入先は全行で同じときだけ表示）</summary>
    Private Shared Function MergeRows(rows As List(Of StockRow)) As StockRow
        If rows.Count = 1 Then Return rows(0)
        Dim first = rows(0)
        Return New StockRow With {
            .ItemCode = first.ItemCode,
            .Model = first.Model,
            .Warehouse = StockBlock.CommonValueOf(rows, Function(r) r.Warehouse, ""),
            .CustomerCode = StockBlock.CommonValueOf(rows, Function(r) r.CustomerCode, ""),
            .CustomerShortCode = StockBlock.CommonValueOf(rows, Function(r) r.CustomerShortCode, ""),
            .CustomerName = StockBlock.CommonValueOf(rows, Function(r) r.CustomerName, ""),
            .SupplierShortCode = StockBlock.CommonValueOf(rows, Function(r) r.SupplierShortCode, ""),
            .SupplierName = StockBlock.CommonValueOf(rows, Function(r) r.SupplierName, ""),
            .StatusCode = first.StatusCode,
            .StatusName = first.StatusName,
            .IsCountTarget = first.IsCountTarget,
            .Quantity = rows.Sum(Function(r) r.Quantity),
            .UnitPrice = rows.Max(Function(r) r.UnitPrice)
        }
    End Function

    ' ============================================================
    ' 一覧シート
    ' ============================================================
    Private Sub WriteListSheet(xl As ExcelObject, items As List(Of StocktakingItem), opt As StocktakingOptions)
        Dim lines As New List(Of Object())
        Dim groupSpans As New List(Of RowSpan)     ' 行グループ（集計行の下の内訳行）

        Dim rowNo As Integer = L_FIRST_ROW
        For Each item In items
            For Each b In item.Blocks
                Dim entries = DetailEntries(b, opt)
                Dim details = If(entries.Count > 0, New RowSpan(rowNo + 1, rowNo + entries.Count), New RowSpan(0, 0))
                ' 共用棚の内訳行（得意先別）は専用欄に印を付けるとその場で専用棚の行として計算される
                Dim promotable As Boolean = opt.ShowBreakdown AndAlso Not b.IsDedicated AndAlso entries.Count > 0
                Dim hasOffShelf As Boolean = entries.Any(Function(e) e.RowType = ListHeaders.ROW_EXCLUDED)
                lines.Add(ListTopLine(item, b, rowNo, details, promotable, hasOffShelf))
                rowNo += 1
                For Each e In entries
                    lines.Add(ListDetailLine(item, e.Row, b.No, e.RowType, rowNo))
                    rowNo += 1
                Next
                If entries.Count > 0 Then groupSpans.Add(details)
            Next
        Next

        Dim firstAppendRow As Integer = rowNo
        For i As Integer = 1 To APPEND_ROWS
            lines.Add(ListAppendLine(rowNo))
            rowNo += 1
        Next
        Dim lastRow As Integer = rowNo - 1

        ' 並び順（並べ替えた後に出力時の並びへ戻すための通し番号）
        For i As Integer = 0 To lines.Count - 1
            lines(i)(L_ORDER - 1) = i + 1
        Next

        ' 1行目・2行目 + 明細を配列に詰める
        Dim data(lastRow - 1, L_COLS - 1) As Object
        data(0, L_NO - 1) = $"{opt.OutputAt:yyyy/MM/dd HH:mm} 時点"
        data(0, L_NAME - 1) = opt.OfficeLabel
        data(0, L_PLUS - 1) = "売上漏れ"
        data(0, L_MINUS - 1) = "仕入漏れ"
        data(0, L_AMOUNT - 1) = $"=SUBTOTAL(9,{Col(L_AMOUNT)}{L_FIRST_ROW}:{Col(L_AMOUNT)}{lastRow})"

        Dim headers As New List(Of String) From {
            ListHeaders.NO, "倉庫", ListHeaders.CUSTOMER, "仕入先", ListHeaders.ITEM_CODE, ListHeaders.MODEL, "ステータス", ListHeaders.MARK, ListHeaders.REGISTERED}
        For i As Integer = 1 To SLOTS
            headers.Add(ListHeaders.SHELF)
            headers.Add("数量")
        Next
        headers.AddRange({"棚卸総数", "AX在庫数", "内訳在庫", "＋", "－", "差数", "単価", "差額", ListHeaders.NOTE,
                          ListHeaders.CUSTOMER_CODE, ListHeaders.ROW_TYPE, ListHeaders.ORDER})
        For c As Integer = 0 To L_COLS - 1
            data(L_HEADER_ROW - 1, c) = headers(c)
        Next

        For i As Integer = 0 To lines.Count - 1
            For c As Integer = 0 To L_COLS - 1
                data(L_FIRST_ROW - 1 + i, c) = lines(i)(c)
            Next
        Next

        ' 表示形式（コードの先頭ゼロを残すため値の設定前に行う）
        ' 得意先・得意先CD は共用棚の行で数式（代表得意先）を使うため標準書式のまま、値は ' 付きの文字列で書く
        xl.SetColumnNumberFormat(L_WH, nfString)
        xl.SetColumnsNumberFormat(L_SUP, L_MARK, nfString)
        For i As Integer = 0 To SLOTS - 1
            xl.SetColumnNumberFormat(L_SHELF + i * 2, nfString)
            xl.SetColumnNumberFormat(L_SHELF + i * 2 + 1, nfInteger)
        Next
        xl.SetColumnsNumberFormat(L_TOTAL, L_DIFF, nfInteger)
        xl.SetColumnNumberFormat(L_PRICE, If(HasFraction(items), nfDecimal2, nfInteger))
        xl.SetColumnNumberFormat(L_AMOUNT, nfInteger)
        xl.SetColumnNumberFormat(L_ROW_TYPE, nfString)

        xl.SetValue(data)

        ' 配置・列幅
        For Each c In {L_NO, L_WH, L_CUST, L_SUP, L_STATUS, L_MARK, L_REG}
            xl.SetColumnAlign(c, xlCenter)
        Next
        xl.SetRowAlign(L_HEADER_ROW, xlCenter)
        xl.SetCellAlign(1, L_PLUS, xlCenter)
        xl.SetCellAlign(1, L_MINUS, xlCenter)

        Dim widths As New List(Of Double) From {5, 5.5, 5.5, 5.5, 12, 30, 8, 4.5, 5}
        For i As Integer = 1 To SLOTS
            widths.Add(8)
            widths.Add(7)
        Next
        widths.AddRange({8.5, 8.5, 8.5, 7, 7, 7.5, 10, 12, 24, 10, 8, 8})
        xl.SetColumnsWidth(widths.ToArray())
        For Each c In {L_CUST_CODE, L_ROW_TYPE, L_ORDER}
            xl.SetColumnHidden(c)
        Next

        Dim sheetName As String = xl.SheetName
        xl.WithSheet(sheetName,
            Sub(sh)
                ' 行の色は非表示の行種別で決める（並べ替えても追従する）
                Dim typeCol = $"INDEX(${Col(L_ROW_TYPE)}:${Col(L_ROW_TYPE)},ROW())"
                Dim body = sh.Range(sh.Cells(L_FIRST_ROW, 1), sh.Cells(lastRow, L_NOTE))
                ' 内訳は灰色。ただし共用棚の内訳行で専用欄に印があれば外す（専用棚に昇格する行）
                Dim markCol = $"INDEX(${Col(L_MARK)}:${Col(L_MARK)},ROW())"
                Dim fcDetail = body.FormatConditions.Add(Type:=xlExpression,
                    Formula1:=$"=OR(AND({typeCol}=""{ListHeaders.ROW_DETAIL}"",{markCol}=""""),{typeCol}=""{ListHeaders.ROW_DEDICATED_DETAIL}"",{typeCol}=""{ListHeaders.ROW_STATUS}"")")
                fcDetail.Interior.Color = COLOR_DETAIL
                Dim fcExcluded = body.FormatConditions.Add(Type:=xlExpression, Formula1:=$"={typeCol}=""{ListHeaders.ROW_EXCLUDED}""")
                fcExcluded.Interior.Color = COLOR_EXCLUDED_BACK
                fcExcluded.Font.Color = COLOR_EXCLUDED_FONT

                ' 印刷対象外の行の棚卸総数は入力欄（保管先で確認した数）
                Dim totalCells = sh.Range(sh.Cells(L_FIRST_ROW, L_TOTAL), sh.Cells(lastRow, L_TOTAL))
                Dim fcInput = totalCells.FormatConditions.Add(Type:=xlExpression, Formula1:=$"={typeCol}=""{ListHeaders.ROW_EXCLUDED}""")
                fcInput.Interior.Color = COLOR_INPUT
                fcInput.Font.Color = RGB(0, 0, 0)
                fcInput.SetFirstPriority()

                ' 得意先が複数（登録数2以上）の共用棚の行は専用にできない：専用欄に書いたら赤で知らせる
                Dim markCells = sh.Range(sh.Cells(L_FIRST_ROW, L_MARK), sh.Cells(lastRow, L_MARK))
                Dim fcNg = markCells.FormatConditions.Add(Type:=xlExpression,
                    Formula1:=$"=AND({typeCol}=""{ListHeaders.ROW_SHARED}"",{markCol}<>"""",N(INDEX(${Col(L_REG)}:${Col(L_REG)},ROW()))>=2)")
                fcNg.Interior.Color = COLOR_NG
                fcNg.SetFirstPriority()

                ' 追記行の枠
                sh.Range(sh.Cells(firstAppendRow, 1), sh.Cells(lastRow, L_NOTE)).Borders.LineStyle = xlContinuous

                ' 行グループのボタンを集計行側に出す
                sh.Outline.SummaryRow = 0   ' xlAbove
            End Sub)

        For Each s In groupSpans
            xl.SetRowsGroup(s.First, s.Last)
        Next
        ' Excel は並べ替え範囲の先頭行がグループの親でないとグループを無視して並べ替えるため、明細全体を外側のグループで囲む
        If groupSpans.Count > 0 Then xl.SetRowsGroup(L_FIRST_ROW, lastRow)

        ' 商品CD と ＋/－ は折りたたみ
        xl.SetColumnGroup(L_ITEM)
        xl.SetColumnsGroup(L_PLUS, L_MINUS)
        xl.ColumnLevels = 1

        ' 非表示列（得意先CD・行種別・並び順）まで含める。並べ替えたときに一緒に動かすため
        xl.SetAutoFilter(L_HEADER_ROW, 1, lastRow, L_COLS)

        xl.WithSheet(sheetName,
            Sub(sh)
                With sh.PageSetup
                    .Orientation = xlLandscape
                    .PaperSize = xlPaperA4
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                    .PrintTitleRows = $"$1:${L_HEADER_ROW}"
                    .TopMargin = 40
                    .BottomMargin = 40
                    .LeftMargin = 20
                    .RightMargin = 20
                    .CenterFooter = "&P/&N"
                End With
            End Sub)
    End Sub

    Private Function NewListLine() As Object()
        Return New Object(L_COLS - 1) {}
    End Function

    ''' <summary>集計対象の行（共用棚・専用棚）。AX在庫数は全ステータス（印刷対象外を含む）の合算</summary>
    ''' <param name="details">直下の内訳行の範囲（無ければ First=0）</param>
    ''' <param name="promotable">得意先別の内訳行があり、専用欄の印で専用棚に昇格できる共用棚の行か</param>
    ''' <param name="hasOffShelf">印刷対象外の内訳行（確認した数を入れる行）があるか</param>
    Private Function ListTopLine(item As StocktakingItem, b As StockBlock, row As Integer, details As RowSpan,
                                 promotable As Boolean, hasOffShelf As Boolean) As Object()
        Dim line = NewListLine()
        Dim rep = b.Representative

        line(L_NO - 1) = b.No
        line(L_WH - 1) = b.CommonValue(Function(r) r.Warehouse)
        line(L_CUST - 1) = Txt(rep.CustomerShortCode)                     ' 代表得意先（短縮CDが最も若い得意先）
        line(L_SUP - 1) = b.CommonValue(Function(r) r.SupplierShortCode, useRepresentative:=True)
        line(L_ITEM - 1) = item.ItemCode
        line(L_NAME - 1) = item.Model
        line(L_STATUS - 1) = b.CommonValueAll(Function(r) StatusLabel(r))  ' ステータスが混ざるときは空欄（内訳行に出る）
        line(L_MARK - 1) = If(b.IsDedicated, ListHeaders.MARK_VALUE, Nothing)
        line(L_REG - 1) = b.CustomerCount
        SetListShelves(line, b.Shelves)

        ' 棚卸総数＝棚の数量＋印刷対象外の内訳行に入れた数（保管先で確認した数）
        If hasOffShelf Then
            Dim types = $"{Col(L_ROW_TYPE)}{details.First}:{Col(L_ROW_TYPE)}{details.Last}"
            Dim totals = $"{Col(L_TOTAL)}{details.First}:{Col(L_TOTAL)}{details.Last}"
            line(L_TOTAL - 1) = $"={QtySum(row)}+SUMIF({types},""{ListHeaders.ROW_EXCLUDED}"",{totals})"
        Else
            line(L_TOTAL - 1) = $"={QtySum(row)}"
        End If

        If promotable Then
            ' 内訳行のうち専用欄に印の無い得意先だけを集計（印を付けた得意先は専用棚の行として別に計算される）
            Dim mark = $"{Col(L_MARK)}{details.First}:{Col(L_MARK)}{details.Last}"
            Dim types = $"{Col(L_ROW_TYPE)}{details.First}:{Col(L_ROW_TYPE)}{details.Last}"
            line(L_BOOK - 1) = $"=SUMIF({mark},"""",{Col(L_DETAIL)}{details.First}:{Col(L_DETAIL)}{details.Last})"
            line(L_PRICE - 1) = $"=MAXIFS({Col(L_PRICE)}{details.First}:{Col(L_PRICE)}{details.Last},{mark},"""")"
            ' 登録数＝印の無い得意先の数（棚を使う得意先なので印刷対象の内訳行だけ。同じ得意先の倉庫・ステータス違いは1つと数える）
            Dim codes = $"{Col(L_CUST_CODE)}{details.First}:{Col(L_CUST_CODE)}{details.Last}"
            ' 代表得意先＝印の無い最初の内訳行（内訳は短縮CD順）。内訳を専用にしたら次の得意先に替わる
            line(L_CUST - 1) = FirstUnmarked(L_CUST, details, mark)
            line(L_CUST_CODE - 1) = FirstUnmarked(L_CUST_CODE, details, mark)
            line(L_REG - 1) = $"=ROUND(SUMPRODUCT(({mark}="""")*({types}=""{ListHeaders.ROW_DETAIL}"")/COUNTIFS({codes},{codes},{mark},{mark}&"""",{types},{types}&"""")),0)"
        Else
            line(L_BOOK - 1) = CDbl(b.TotalQuantity)
            line(L_PRICE - 1) = CDbl(b.UnitPrice)
        End If
        line(L_DIFF - 1) = $"={Col(L_TOTAL)}{row}-{Col(L_BOOK)}{row}+{Col(L_PLUS)}{row}-{Col(L_MINUS)}{row}"
        line(L_AMOUNT - 1) = $"={Col(L_DIFF)}{row}*{Col(L_PRICE)}{row}"
        If line(L_CUST_CODE - 1) Is Nothing Then line(L_CUST_CODE - 1) = Txt(If(b.IsDedicated, b.CustomerCode, rep.CustomerCode))
        line(L_ROW_TYPE - 1) = If(b.IsDedicated, ListHeaders.ROW_DEDICATED, ListHeaders.ROW_SHARED)
        Return line
    End Function

    ''' <summary>
    ''' 内訳行（集計対象外。在庫は「内訳在庫」列に出す）。
    ''' 単価は常に表示。共用棚の内訳行は、専用欄に印を付けると AX在庫数・棚卸総数・差数・差額が出て専用棚の行として計算される。
    ''' 印刷対象外の行は棚卸総数の欄が入力欄（保管先で確認した数。集計行の棚卸総数に足される）
    ''' </summary>
    Private Function ListDetailLine(item As StocktakingItem, r As StockRow, no As Object, rowType As String, row As Integer) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = no
        line(L_WH - 1) = r.Warehouse
        line(L_CUST - 1) = Txt(r.CustomerShortCode)
        line(L_SUP - 1) = r.SupplierShortCode
        line(L_ITEM - 1) = item.ItemCode
        line(L_NAME - 1) = item.Model
        line(L_STATUS - 1) = StatusLabel(r)
        line(L_DETAIL - 1) = CDbl(r.Quantity)
        line(L_PRICE - 1) = CDbl(r.UnitPrice)                              ' 内訳行にも単価を表示（集計しない列）
        line(L_CUST_CODE - 1) = Txt(r.CustomerCode)
        line(L_ROW_TYPE - 1) = rowType

        If rowType = ListHeaders.ROW_DETAIL Then
            Dim marked = $"{Col(L_MARK)}{row}<>"""""
            line(L_TOTAL - 1) = $"=IF({marked},{QtySum(row)},"""")"
            line(L_BOOK - 1) = $"=IF({marked},{Col(L_DETAIL)}{row},"""")"
            line(L_DIFF - 1) = $"=IF({marked},{Col(L_TOTAL)}{row}-{Col(L_BOOK)}{row}+{Col(L_PLUS)}{row}-{Col(L_MINUS)}{row},"""")"
            line(L_AMOUNT - 1) = $"=IF({marked},{Col(L_DIFF)}{row}*{Col(L_PRICE)}{row},"""")"
        ElseIf rowType = ListHeaders.ROW_EXCLUDED Then
            line(L_NOTE - 1) = ListHeaders.EXCLUDED_NOTE
        End If
        Return line
    End Function

    ''' <summary>標準書式の列に先頭ゼロ付きのコードを文字列として書く（' を付ける）</summary>
    Private Shared Function Txt(value As String) As Object
        Return If(String.IsNullOrEmpty(value), Nothing, "'" & value)
    End Function

    ''' <summary>内訳行のうち専用欄が空の最初の行の値を返す数式</summary>
    Private Function FirstUnmarked(column As Integer, details As RowSpan, markRange As String) As String
        Return $"=IFERROR(INDEX({Col(column)}{details.First}:{Col(column)}{details.Last},MATCH(TRUE,INDEX({markRange}="""",0),0)),"""")"
    End Function

    ''' <summary>棚番を位置どおりの記入欄へ（表示数より後ろの位置にある棚番は備考に件数を出す）</summary>
    Private Sub SetListShelves(line As Object(), shelves As IDictionary(Of Integer, String))
        For Each kv In shelves.Where(Function(x) x.Key <= SLOTS)
            line(L_SHELF - 1 + (kv.Key - 1) * 2) = kv.Value
        Next
        Dim hidden = shelves.Keys.Where(Function(p) p > SLOTS).Count()
        If hidden > 0 Then
            line(L_NOTE - 1) = $"棚番{SLOTS + 1}以降に{hidden}件（表示省略）"
        End If
    End Sub

    ''' <summary>追記用の空行（数式のみ）</summary>
    Private Function ListAppendLine(row As Integer) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = APPEND_LABEL
        line(L_TOTAL - 1) = $"={QtySum(row)}"
        line(L_DIFF - 1) = $"={Col(L_TOTAL)}{row}-{Col(L_BOOK)}{row}+{Col(L_PLUS)}{row}-{Col(L_MINUS)}{row}"
        line(L_AMOUNT - 1) = $"={Col(L_DIFF)}{row}*{Col(L_PRICE)}{row}"
        Return line
    End Function

    ''' <summary>数量列の合計式（SUM(K3,M3,...)）</summary>
    Private Function QtySum(row As Integer) As String
        Dim cells = Enumerable.Range(0, SLOTS).Select(Function(i) $"{Col(L_SHELF + i * 2 + 1)}{row}")
        Return $"SUM({String.Join(",", cells)})"
    End Function

    ' ============================================================
    ' 印刷用シート
    ' ============================================================
    Private Sub WritePrintSheet(xl As ExcelObject, items As List(Of StocktakingItem), opt As StocktakingOptions)
        Const P_NO As Integer = 1
        Const P_WH As Integer = 2
        Const P_CUST As Integer = 3
        Const P_SUP As Integer = 4
        Const P_NAME As Integer = 5
        Const P_MARK As Integer = 6                               ' 専用（専用棚の行に○）
        Const P_STATUS As Integer = 7                             ' ステータス（混在する行は空欄。内訳行に出る）
        Const P_REG As Integer = 8
        Dim pStock As Integer = If(opt.ShowStockOnPrintSheet, 9, 0)
        Dim pShelf As Integer = If(opt.ShowStockOnPrintSheet, 10, 9)
        Dim pOrder As Integer = pShelf + SLOTS * 2               ' 並び順（非表示）
        Dim pCols As Integer = pOrder

        Dim lines As New List(Of Object())
        Dim detailSpans As New List(Of RowSpan)

        Dim newLine = Function(no As Object, wh As String, cust As String, sup As String, name As String, mark As String, status As String,
                               reg As Object, stock As Decimal, shelves As IDictionary(Of Integer, String)) As Object()
                          Dim line(pCols - 1) As Object
                          line(P_NO - 1) = no
                          line(P_WH - 1) = wh
                          line(P_CUST - 1) = cust
                          line(P_SUP - 1) = sup
                          line(P_NAME - 1) = name
                          line(P_MARK - 1) = mark
                          line(P_STATUS - 1) = status
                          line(P_REG - 1) = reg
                          If pStock > 0 Then line(pStock - 1) = CDbl(stock)
                          If shelves IsNot Nothing Then
                              For Each kv In shelves.Where(Function(x) x.Key <= SLOTS)
                                  line(pShelf - 1 + (kv.Key - 1) * 2) = kv.Value
                              Next
                          End If
                          Return line
                      End Function

        Dim rowNo As Integer = 2
        For Each item In items
            ' 印刷対象外だけのブロックは棚に無いので載せない（一覧シートで数を入れる）
            For Each b In item.Blocks.Where(Function(x) x.HasPrintRows)
                Dim rep = b.Representative
                lines.Add(newLine(b.No, b.CommonValue(Function(r) r.Warehouse), rep.CustomerShortCode,
                                  b.CommonValue(Function(r) r.SupplierShortCode, useRepresentative:=True),
                                  item.Model, If(b.IsDedicated, ListHeaders.MARK_VALUE, Nothing), b.CommonValue(Function(r) StatusLabel(r)),
                                  b.CustomerCount, b.Quantity, b.Shelves))
                rowNo += 1

                Dim printDetails = PrintDetailRows(b, opt)
                If printDetails.Count > 0 Then
                    Dim first As Integer = rowNo
                    For Each r In printDetails
                        lines.Add(newLine(b.No, r.Warehouse, r.CustomerShortCode, r.SupplierShortCode,
                                          item.Model, Nothing, StatusLabel(r), Nothing, r.Quantity, Nothing))
                        rowNo += 1
                    Next
                    detailSpans.Add(New RowSpan(first, rowNo - 1))
                End If
            Next
        Next

        For i As Integer = 1 To APPEND_ROWS
            Dim line(pCols - 1) As Object
            line(P_NO - 1) = APPEND_LABEL
            lines.Add(line)
            rowNo += 1
        Next
        Dim lastRow As Integer = rowNo - 1

        For i As Integer = 0 To lines.Count - 1
            lines(i)(pOrder - 1) = i + 1
        Next

        Dim data(lastRow - 1, pCols - 1) As Object
        Dim headers As New List(Of String) From {"No", "倉庫", "得意先", "仕入先", ListHeaders.MODEL, "専用", "ステータス", "登録数"}
        If pStock > 0 Then headers.Add("在庫数")
        For i As Integer = 1 To SLOTS
            headers.Add("棚番")
            headers.Add("数量")
        Next
        headers.Add(ListHeaders.ORDER)
        For c As Integer = 0 To pCols - 1
            data(0, c) = headers(c)
        Next
        For i As Integer = 0 To lines.Count - 1
            For c As Integer = 0 To pCols - 1
                data(1 + i, c) = lines(i)(c)
            Next
        Next

        xl.SetColumnsNumberFormat(P_WH, P_STATUS, nfString)
        If pStock > 0 Then xl.SetColumnNumberFormat(pStock, nfInteger)
        For i As Integer = 0 To SLOTS - 1
            xl.SetColumnNumberFormat(pShelf + i * 2, nfString)
        Next

        xl.SetValue(data)

        ' 列幅：旧アプリと同じ総幅（記入欄 97/105）から倉庫列の分を引き、棚番:数量 = 1:3 で配分。
        ' 専用・ステータス列は細く（棚番別と同じく7掛け）。横幅が A4 横に収まらない分は下で記入欄を縮める
        Dim widths As New List(Of Double) From {3, 4, 4, 4, 13, 1.75, 2.45, 3}
        If pStock > 0 Then widths.Add(8)
        Dim slotWidth As Double = If(pStock > 0, 93.0, 101.0) / SLOTS
        For i As Integer = 1 To SLOTS
            widths.Add(slotWidth * 0.25)
            widths.Add(slotWidth * 0.75)
        Next
        widths.Add(8)
        xl.SetColumnsWidth(widths.ToArray())
        xl.SetColumnHidden(pOrder)

        For Each c In {P_NO, P_WH, P_CUST, P_SUP, P_MARK, P_STATUS, P_REG}
            xl.SetColumnAlign(c, xlCenter)
        Next
        xl.SetRowAlign(1, xlCenter)

        For Each s In detailSpans
            xl.SetCellsBackColor(s.First, 1, s.Last, pCols, COLOR_DETAIL)
        Next

        ' 内訳行を親の行のグループにする（閉じて並べ替えると親と一緒に動く）
        xl.WithSheet(PRINT_SHEET_NAME, Sub(sh) sh.Outline.SummaryRow = 0)   ' xlAbove
        For Each s In detailSpans
            xl.SetRowsGroup(s.First, s.Last)
        Next
        ' 一覧シートと同じく、明細全体を外側のグループで囲む（先頭行に内訳がなくてもグループごと並べ替わるように）
        If detailSpans.Count > 0 Then xl.SetRowsGroup(2, lastRow)

        xl.SetAutoFilter(1, 1, lastRow, pCols)

        Dim headerText As String = opt.OfficeHeader
        Dim outputText As String = $"出力：{opt.OutputAt:yyyy/MM/dd HH:mm}"
        xl.WithSheet(PRINT_SHEET_NAME,
            Sub(sh)
                Dim all = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, pCols))
                all.RowHeight = 37
                all.ShrinkToFit = True                       ' 文字サイズをセル幅に合わせる
                all.Borders.LineStyle = xlContinuous         ' 格子
                sh.Columns(P_NAME).WrapText = True           ' 型式はセル幅で折り返し

                ' 旧アプリと同じ大きさ（倍率100%）で刷れるよう、横幅が A4 横の印刷範囲を超える分だけ記入欄（棚番・数量）を狭める
                ' （横1ページに収める設定なので、はみ出すと Excel が全体を縮小して文字・枠が小さくなる）
                Const PRINTABLE_WIDTH As Double = 842 - 6 - 6 - 4   ' A4横の幅 − 左右余白 − ゆとり（pt）
                For pass As Integer = 1 To 3
                    Dim total As Double = 0
                    Dim slotTotal As Double = 0
                    For c As Integer = 1 To pOrder - 1
                        total += CDbl(sh.Columns(c).Width)
                        If c >= pShelf Then slotTotal += CDbl(sh.Columns(c).Width)
                    Next
                    If total <= PRINTABLE_WIDTH Then Exit For
                    Dim factor As Double = (slotTotal - (total - PRINTABLE_WIDTH)) / slotTotal
                    For c As Integer = pShelf To pOrder - 1
                        sh.Columns(c).ColumnWidth = CDbl(sh.Columns(c).ColumnWidth) * factor
                    Next
                Next

                With sh.PageSetup
                    .Orientation = xlLandscape
                    .PaperSize = xlPaperA4
                    ' 余白は最小限（ヘッダー分だけ上を空ける）にし、横1ページに収める
                    .TopMargin = 24
                    .BottomMargin = 6
                    .LeftMargin = 6
                    .RightMargin = 6
                    .HeaderMargin = 6
                    .FooterMargin = 0
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                    .CenterHorizontally = True
                    .PrintTitleRows = "$1:$1"
                    .LeftHeader = outputText
                    .CenterHeader = headerText
                    .RightHeader = "ページ：&P/&N"
                End With
            End Sub)
    End Sub

    ' ============================================================
    ' 棚番別シート
    ' ============================================================
    ''' <summary>
    ''' 棚番1つにつき1行（棚番 × 集計行）。複数の棚にある型式は棚の数だけ行が出る。
    ''' 列は No → 棚番 → 棚数 → 型式 → 専用 → ステータス → (在庫数) → 数量 → 倉庫 → 得意先 → 仕入先。
    ''' 棚数はその集計行に登録されている棚番の数（他の棚にもあるかが分かるように）。
    ''' 得意先・仕入先は略称。コードを表示する設定なら「短縮CD（改行）略称」の1セル2行。
    ''' 横向き印刷・MS ゴシック 12pt（型式の空白位置を揃えるため等幅）。明細行は「○個入り×何箱」の計算も書ける高さ。
    ''' 並べ替え・絞り込みは現場でオートフィルタを使う（棚番 → 仕入先 → 型式の順で出力）。
    ''' 数えた数は一覧シートの同じ No の行で、同じ棚番の数量欄に書き写す
    ''' </summary>
    Private Sub WriteShelfSheet(xl As ExcelObject, items As List(Of StocktakingItem), opt As StocktakingOptions)
        Const MAX_POSITION As Integer = 5                         ' 棚番5まで
        Const S_NO As Integer = 1
        Const S_SHELF As Integer = 2
        Const S_COUNT As Integer = 3                              ' 棚数（登録されている棚番の数。棚番5より後ろも数える）
        Const S_NAME As Integer = 4
        Const S_MARK As Integer = 5
        Const S_STATUS As Integer = 6
        Dim sStock As Integer = If(opt.ShowStockOnPrintSheet, 7, 0)
        Dim sQty As Integer = If(opt.ShowStockOnPrintSheet, 8, 7)
        Dim sWh As Integer = sQty + 1                             ' 倉庫・得意先・仕入先は一番右
        Dim sCust As Integer = sQty + 2
        Dim sSup As Integer = sQty + 3
        Dim sOrder As Integer = sSup + 1                          ' 並び順（非表示）
        Dim sCols As Integer = sOrder

        ' コードを出すときは「短縮CD（改行）略称」
        Dim party = Function(code As String, name As String) As String
                        If Not opt.ShowCodesOnShelfSheet OrElse String.IsNullOrEmpty(code) Then Return name
                        Return If(String.IsNullOrEmpty(name), code, code & vbLf & name)
                    End Function

        ' 1行＝棚番 × 集計行 × 倉庫 × ステータス。同じ棚・同じ型式でも倉庫やステータスが違えば別の行にし、在庫数もそれぞれ出す
        ' 棚番 → 仕入先（短縮CD）→ 型式 → 倉庫 → ステータス の順（棚番なしは最後）
        Dim entries As New List(Of Tuple(Of String, StocktakingItem, StockBlock, List(Of StockRow)))
        For Each item In items
            For Each b In item.Blocks.Where(Function(x) x.HasPrintRows)
                Dim groups = b.Rows.GroupBy(Function(x) x.Warehouse & vbTab & x.StatusCode) _
                                   .Select(Function(g) g.ToList()) _
                                   .OrderBy(Function(g) g(0).Warehouse, StringComparer.Ordinal) _
                                   .ThenBy(Function(g) g(0).StatusCode, StringComparer.Ordinal) _
                                   .ToList()
                Dim shelves = b.Shelves.Where(Function(kv) kv.Key <= MAX_POSITION).Select(Function(kv) kv.Value).ToList()
                If shelves.Count = 0 Then shelves.Add("")
                For Each shelf In shelves
                    For Each g In groups
                        entries.Add(Tuple.Create(shelf, item, b, g))
                    Next
                Next
            Next
        Next
        Dim supplierOf = Function(rows As List(Of StockRow)) StockBlock.CommonValueOf(rows, Function(x) x.SupplierShortCode, rows(0).SupplierShortCode)
        Dim ordered = entries.OrderBy(Function(e) If(e.Item1 = "", 1, 0)) _
                             .ThenBy(Function(e) e.Item1, StringComparer.Ordinal) _
                             .ThenBy(Function(e) supplierOf(e.Item4), StringComparer.Ordinal) _
                             .ThenBy(Function(e) e.Item2.Model, StringComparer.Ordinal) _
                             .ThenBy(Function(e) e.Item3.No) _
                             .ThenBy(Function(e) e.Item4(0).Warehouse, StringComparer.Ordinal) _
                             .ThenBy(Function(e) e.Item4(0).StatusCode, StringComparer.Ordinal) _
                             .ToList()

        Dim lastRow As Integer = 1 + ordered.Count + APPEND_ROWS
        Dim data(lastRow - 1, sCols - 1) As Object
        Dim headers As New List(Of String) From {"No", "棚番", "棚数", ListHeaders.MODEL, "専用", "ステ"}   ' 見出しは縮小しないので細い列に収まる「ステ」
        If sStock > 0 Then headers.Add("在庫数")
        headers.AddRange({"数量", "倉庫", "得意先", "仕入先", ListHeaders.ORDER})
        For c As Integer = 0 To sCols - 1
            data(0, c) = headers(c)
        Next

        For i As Integer = 0 To ordered.Count - 1
            Dim e = ordered(i)
            Dim b = e.Item3
            Dim rows = e.Item4                                    ' この行の在庫（同じ倉庫・同じステータス）
            Dim rep = rows(0)                                     ' 代表得意先（行は得意先の短縮CD順）
            Dim others As Integer = rows.Select(Function(x) x.CustomerCode).Distinct().Count() - 1
            Dim r As Integer = i + 1
            data(r, S_NO - 1) = b.No
            data(r, S_SHELF - 1) = e.Item1
            data(r, S_COUNT - 1) = If(b.Shelves.Count > 0, CObj(b.Shelves.Count), Nothing)
            data(r, S_NAME - 1) = e.Item2.Model
            data(r, S_MARK - 1) = If(b.IsDedicated, ListHeaders.MARK_VALUE, Nothing)
            data(r, S_STATUS - 1) = StatusLabel(rep)
            If sStock > 0 Then data(r, sStock - 1) = CDbl(rows.Sum(Function(x) x.Quantity))
            data(r, sWh - 1) = rep.Warehouse                      ' 倉庫区分は設定に関係なく出す
            ' 共用棚に複数の得意先がいるときは代表得意先＋「他n」
            data(r, sCust - 1) = party(rep.CustomerShortCode, If(others > 0, $"{rep.CustomerName} 他{others}", rep.CustomerName))
            Dim sup = supplierOf(rows)
            data(r, sSup - 1) = party(sup, rows.Where(Function(x) x.SupplierShortCode = sup).Select(Function(x) x.SupplierName).FirstOrDefault())
            data(r, sOrder - 1) = r
        Next
        For i As Integer = 1 To APPEND_ROWS
            Dim r As Integer = ordered.Count + i
            data(r, S_NO - 1) = APPEND_LABEL
            data(r, sOrder - 1) = r
        Next

        xl.SetColumnsNumberFormat(S_SHELF, S_STATUS, nfString)
        xl.SetColumnsNumberFormat(sWh, sSup, nfString)
        If sStock > 0 Then xl.SetColumnNumberFormat(sStock, nfInteger)

        xl.SetValue(data)

        ' 文字を大きめ（MS ゴシック 12pt）にするので幅も広めに。数量は計算も書けるよう広く、専用とステータスは同じ幅
        ' 型式 45（大きい文字で見やすく）、数量 30、棚数 4、専用・ステ 5、得意先・仕入先 12
        Dim widths As New List(Of Double) From {6, 13, 4, 45, 5, 5}
        If sStock > 0 Then widths.Add(11)
        widths.AddRange({30, 7, 12, 12, 8})
        xl.SetColumnsWidth(widths.ToArray())
        xl.SetColumnHidden(sOrder)

        For Each c In {S_NO, S_COUNT, S_MARK, S_STATUS, sWh}
            xl.SetColumnAlign(c, xlCenter)
        Next
        xl.SetRowAlign(1, xlCenter)                               ' 見出し行は中央揃え
        xl.SetAutoFilter(1, 1, lastRow, sCols)

        Dim zoomPercent As Integer = 100
        Dim headerText As String = opt.OfficeHeader & "（棚番別）"
        Dim outputText As String = $"出力：{opt.OutputAt:yyyy/MM/dd HH:mm}"
        xl.WithSheet(SHELF_SHEET_NAME,
            Sub(sh)
                Dim all = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, sCols))
                ' 型式の空白位置が揃うよう等幅の MS ゴシックにする
                all.Font.Name = "ＭＳ ゴシック"
                all.Font.Size = 12
                all.Borders.LineStyle = xlContinuous
                ' 明細行は「○個入り×何箱」の計算も書ける高さで、印刷1ページにちょうど ROWS_PER_PAGE 行。見出し行は縮小しない
                Const ROWS_PER_PAGE As Integer = 11
                Const HEADER_HEIGHT As Double = 26
                Const PAGE_WIDTH As Double = 842 - 6 - 6        ' A4横 − 左右余白（pt）
                ' Excel の実際の印刷寸法は列幅・行高の値より1割ほど小さい（実測 約0.9）。計算だけで合わせると左右・下に余白が残るので、
                ' 倍率は大きめから始めて横にはみ出す間は下げ、行高は改ページ位置を見て決める（どちらも出力する PC のプリンターで判定）
                Const PRINT_RATIO As Double = 0.91
                Dim body = sh.Range(sh.Cells(2, 1), sh.Cells(lastRow, sCols))
                body.ShrinkToFit = True
                body.VerticalAlignment = xlCenter
                ' 得意先・仕入先は「短縮CD（改行）略称」になるので折り返す（折り返しと縮小は両立しないため縮小は外す）
                Dim names = sh.Range(sh.Cells(2, sCust), sh.Cells(lastRow, sSup))
                names.ShrinkToFit = False
                names.WrapText = True
                ' 型式は現場で見つけやすいよう大きく（長い型式は縮小表示で収める）
                sh.Range(sh.Cells(2, S_NAME), sh.Cells(lastRow, S_NAME)).Font.Size = 20
                If sStock > 0 Then sh.Range(sh.Cells(2, sStock), sh.Cells(lastRow, sStock)).Font.Size = 16
                sh.Rows(1).RowHeight = HEADER_HEIGHT

                Dim totalWidth As Double = 0
                For c As Integer = 1 To sOrder - 1
                    totalWidth += CDbl(sh.Columns(c).Width)
                Next
                zoomPercent = CInt(Math.Min(100, Math.Floor(PAGE_WIDTH / (totalWidth * PRINT_RATIO) * 100)))
                With sh.PageSetup
                    .Orientation = xlLandscape
                    .PaperSize = xlPaperA4
                    .TopMargin = 24
                    .BottomMargin = 6
                    .LeftMargin = 6
                    .RightMargin = 6
                    .HeaderMargin = 6
                    .FooterMargin = 0
                    .Zoom = zoomPercent                      ' 横1ページに収まる倍率（上で計算）
                    .CenterHorizontally = True
                    .PrintTitleRows = "$1:$1"
                    .LeftHeader = outputText
                    .CenterHeader = headerText
                    .RightHeader = "ページ：&P/&N"
                End With

                ' 改ページ位置は設定を変えても自動では計算し直されない（特に倍率）。総ページ数を読むと計算し直されるので、読む前に必ず呼ぶ
                Dim repaginate = Sub()
                                     Dim pages = sh.PageSetup.Pages.Count
                                 End Sub

                ' 横1ページに収まるまで倍率を下げる
                repaginate()
                Do While sh.VPageBreaks.Count > 0 AndAlso zoomPercent > 30
                    zoomPercent -= 1
                    sh.PageSetup.Zoom = zoomPercent
                    repaginate()
                Loop

                ' 1ページに ROWS_PER_PAGE 行以上入る一番高い行高を、改ページ位置で二分探索（行高は 0.75pt 単位）
                Dim fits = Function(steps As Integer) As Boolean
                               body.RowHeight = steps * 0.75
                               repaginate()
                               If sh.HPageBreaks.Count = 0 Then Return True
                               Return CInt(sh.HPageBreaks(1).Location.Row) - 2 >= ROWS_PER_PAGE   ' 見出し1行を除く
                           End Function
                Dim lo As Integer = 40      ' 30pt
                Dim hi As Integer = 120     ' 90pt
                Do While lo < hi
                    Dim mid As Integer = (lo + hi + 1) \ 2
                    If fits(mid) Then lo = mid Else hi = mid - 1
                Loop
                body.RowHeight = lo * 0.75
            End Sub)
    End Sub

    ' ============================================================
    ' 共通
    ' ============================================================

    ''' <summary>ステータスの表示名（通常は空文字。名称が取れないコードはコードを表示）</summary>
    Private Shared Function StatusLabel(r As StockRow) As String
        If r.StatusCode = "" Then Return ""
        Return If(r.StatusName <> "", r.StatusName, $"ステータス{r.StatusCode}")
    End Function

    ''' <summary>単価に小数を含む行があるか（単価列の表示形式の判定）</summary>
    Private Shared Function HasFraction(items As List(Of StocktakingItem)) As Boolean
        Return items.SelectMany(Function(i) i.Targets.Concat(i.Excluded)).Any(Function(r) r.UnitPrice <> Decimal.Truncate(r.UnitPrice))
    End Function

    Private Shared Function Col(columnIndex As Integer) As String
        Return GetAlphabets(columnIndex)
    End Function

End Class
