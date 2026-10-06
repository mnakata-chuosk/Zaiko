Option Strict Off  ' PageSetup・Outline 等の Excel COM 遅延バインディングに必要

Imports ChuoUtils
Imports ChuoUtils.ExcelTools

''' <summary>
''' 棚卸表（一覧シート・印刷用シート）を Excel に出力し、ユーザーに表示する
''' </summary>
''' <remarks>
''' 集計単位（StockBlock）ごとに1行（集計対象）を出す。
''' ・共用棚の行 … 得意先は代表得意先（短縮CDが最も若い得意先）。専用棚を持たない得意先の在庫の合計
''' ・専用棚の行 … 「専用」欄に○。その得意先の在庫だけ（共用棚の在庫には含めない）
''' ・内訳行（灰色・同じNo）… 得意先・倉庫・ステータス別の在庫（「内訳在庫」列）。集計対象外。出力するかは設定
''' 集計用の列（棚卸総数・AX在庫数・差数・単価・差額）は集計対象の行にだけ値を入れ、二重計上を防ぐ。
''' 両シートとも末尾に追記用の空行を設け、リスト外の現物はそこに書き込む。
''' </remarks>
Public NotInheritable Class StocktakingWriter

    ''' <summary>[棚番・数量] の組数（3～5、一覧・印刷用で共通）</summary>
    Private ReadOnly SLOTS As Integer
    ''' <summary>リスト外の現物を書き込むための追記行数</summary>
    Private Const APPEND_ROWS As Integer = 10
    Private Const APPEND_LABEL As String = ListHeaders.APPEND_NO
    Private Const PRINT_SHEET_NAME As String = "印刷用"

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
    Private ReadOnly L_ROW_TYPE As Integer                      ' 行種別（非表示。共用／専用／内訳／対象外）
    Private ReadOnly L_ORDER As Integer                         ' 並び順（非表示。出力時の並びに戻す用）
    Private ReadOnly L_COLS As Integer
    Private Const L_HEADER_ROW As Integer = 2
    Private Const L_FIRST_ROW As Integer = 3

    ' 色（旧アプリの重複明細と同じ灰色）
    Private Shared ReadOnly COLOR_DETAIL As Integer = RGB(217, 217, 217)
    Private Shared ReadOnly COLOR_EXCLUDED_BACK As Integer = RGB(242, 242, 242)
    Private Shared ReadOnly COLOR_EXCLUDED_FONT As Integer = RGB(128, 128, 128)

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

            ' Show() は最後に SetFreezePanes したシート（=先頭シート）にだけ固定を再適用するため、一覧シートを最後にする
            xl.SetSheet(1)
            xl.SetFreezePanes(L_FIRST_ROW, L_NAME + 1)

            If String.IsNullOrEmpty(savePath) Then
                xl.Show()
            Else
                xl.Calculate()
                xl.SetHome()
                xl.SaveAs(savePath)
                xl.Dispose()
            End If
        Catch
            xl.Dispose()
            Throw
        End Try
    End Sub

    ''' <summary>内訳行を出すブロックか（設定がオンで、在庫行が2行以上）</summary>
    Private Shared Function HasDetail(b As StockBlock, opt As StocktakingOptions) As Boolean
        Return opt.ShowBreakdown AndAlso b.Rows.Count >= 2
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
                Dim topRow As Integer = rowNo
                lines.Add(ListTopLine(item, b, rowNo))
                rowNo += 1
                If HasDetail(b, opt) Then
                    For Each r In b.Rows
                        lines.Add(ListDetailLine(item, r, b.No, ListHeaders.ROW_DETAIL))
                        rowNo += 1
                    Next
                    groupSpans.Add(New RowSpan(topRow + 1, rowNo - 1))
                End If
            Next

            ' 棚卸対象外ステータスは内訳を出すときだけ参考表示
            If opt.ShowBreakdown AndAlso item.Excluded.Count > 0 Then
                Dim first As Integer = rowNo
                Dim no As Object = If(item.Blocks.Count > 0, CObj(item.Blocks(0).No), Nothing)
                For Each r In item.Excluded
                    Dim line = ListDetailLine(item, r, no, ListHeaders.ROW_EXCLUDED)
                    line(L_NOTE - 1) = ListHeaders.EXCLUDED_NOTE
                    lines.Add(line)
                    rowNo += 1
                Next
                If item.Blocks.Count > 0 Then groupSpans.Add(New RowSpan(first, rowNo - 1))
            End If
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
        xl.SetColumnsNumberFormat(L_WH, L_MARK, nfString)
        For i As Integer = 0 To SLOTS - 1
            xl.SetColumnNumberFormat(L_SHELF + i * 2, nfString)
            xl.SetColumnNumberFormat(L_SHELF + i * 2 + 1, nfInteger)
        Next
        xl.SetColumnsNumberFormat(L_TOTAL, L_DIFF, nfInteger)
        xl.SetColumnNumberFormat(L_PRICE, If(HasFraction(items), nfDecimal2, nfInteger))
        xl.SetColumnNumberFormat(L_AMOUNT, nfInteger)
        xl.SetColumnsNumberFormat(L_CUST_CODE, L_ROW_TYPE, nfString)

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
                Dim fcDetail = body.FormatConditions.Add(Type:=xlExpression, Formula1:=$"={typeCol}=""{ListHeaders.ROW_DETAIL}""")
                fcDetail.Interior.Color = COLOR_DETAIL
                Dim fcExcluded = body.FormatConditions.Add(Type:=xlExpression, Formula1:=$"={typeCol}=""{ListHeaders.ROW_EXCLUDED}""")
                fcExcluded.Interior.Color = COLOR_EXCLUDED_BACK
                fcExcluded.Font.Color = COLOR_EXCLUDED_FONT

                ' 追記行の枠
                sh.Range(sh.Cells(firstAppendRow, 1), sh.Cells(lastRow, L_NOTE)).Borders.LineStyle = xlContinuous

                ' 行グループのボタンを集計行側に出す
                sh.Outline.SummaryRow = 0   ' xlAbove
            End Sub)

        For Each s In groupSpans
            xl.SetRowsGroup(s.First, s.Last)
        Next

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

    ''' <summary>集計対象の行（共用棚・専用棚）</summary>
    Private Function ListTopLine(item As StocktakingItem, b As StockBlock, row As Integer) As Object()
        Dim line = NewListLine()
        Dim rep = b.Representative
        Dim isSingle As Boolean = b.Rows.Count = 1

        line(L_NO - 1) = b.No
        line(L_WH - 1) = If(isSingle, rep.Warehouse, b.CommonValue(Function(r) r.Warehouse))
        line(L_CUST - 1) = rep.CustomerShortCode                          ' 代表得意先（短縮CDが最も若い得意先）
        line(L_SUP - 1) = If(isSingle, rep.SupplierShortCode, b.CommonValue(Function(r) r.SupplierShortCode, useRepresentative:=True))
        line(L_ITEM - 1) = item.ItemCode
        line(L_NAME - 1) = item.Model
        line(L_STATUS - 1) = If(isSingle, StatusLabel(rep), b.CommonValue(Function(r) StatusLabel(r)))
        line(L_MARK - 1) = If(b.IsDedicated, ListHeaders.MARK_VALUE, Nothing)
        line(L_REG - 1) = b.CustomerCount
        SetListShelves(line, b.Shelves)
        line(L_TOTAL - 1) = $"={QtySum(row)}"
        line(L_BOOK - 1) = CDbl(b.Quantity)
        line(L_PRICE - 1) = CDbl(b.UnitPrice)
        line(L_DIFF - 1) = $"={Col(L_TOTAL)}{row}-{Col(L_BOOK)}{row}+{Col(L_PLUS)}{row}-{Col(L_MINUS)}{row}"
        line(L_AMOUNT - 1) = $"={Col(L_DIFF)}{row}*{Col(L_PRICE)}{row}"
        line(L_CUST_CODE - 1) = If(b.IsDedicated, b.CustomerCode, rep.CustomerCode)
        line(L_ROW_TYPE - 1) = If(b.IsDedicated, ListHeaders.ROW_DEDICATED, ListHeaders.ROW_SHARED)
        Return line
    End Function

    ''' <summary>内訳行・対象外行（集計対象外。在庫は「内訳在庫」列に出す）</summary>
    Private Function ListDetailLine(item As StocktakingItem, r As StockRow, no As Object, rowType As String) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = no
        line(L_WH - 1) = r.Warehouse
        line(L_CUST - 1) = r.CustomerShortCode
        line(L_SUP - 1) = r.SupplierShortCode
        line(L_ITEM - 1) = item.ItemCode
        line(L_NAME - 1) = item.Model
        line(L_STATUS - 1) = StatusLabel(r)
        line(L_DETAIL - 1) = CDbl(r.Quantity)
        line(L_CUST_CODE - 1) = r.CustomerCode
        line(L_ROW_TYPE - 1) = rowType
        Return line
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
        Const P_REG As Integer = 6
        Dim pStock As Integer = If(opt.ShowStockOnPrintSheet, 7, 0)
        Dim pShelf As Integer = If(opt.ShowStockOnPrintSheet, 8, 7)
        Dim pOrder As Integer = pShelf + SLOTS * 2               ' 並び順（非表示）
        Dim pCols As Integer = pOrder

        Dim lines As New List(Of Object())
        Dim detailSpans As New List(Of RowSpan)

        Dim newLine = Function(no As Object, wh As String, cust As String, sup As String, name As String,
                               reg As Object, stock As Decimal, shelves As IDictionary(Of Integer, String)) As Object()
                          Dim line(pCols - 1) As Object
                          line(P_NO - 1) = no
                          line(P_WH - 1) = wh
                          line(P_CUST - 1) = cust
                          line(P_SUP - 1) = sup
                          line(P_NAME - 1) = name
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
            For Each b In item.Blocks
                Dim rep = b.Representative
                Dim isSingle As Boolean = b.Rows.Count = 1
                Dim name = item.Model
                If isSingle AndAlso StatusLabel(rep) <> "" Then name &= $"【{StatusLabel(rep)}】"
                If b.IsDedicated Then name &= "【専用】"

                lines.Add(newLine(b.No,
                                  If(isSingle, rep.Warehouse, b.CommonValue(Function(r) r.Warehouse)),
                                  rep.CustomerShortCode,
                                  If(isSingle, rep.SupplierShortCode, b.CommonValue(Function(r) r.SupplierShortCode, useRepresentative:=True)),
                                  name, b.CustomerCount, b.Quantity, b.Shelves))
                rowNo += 1

                If HasDetail(b, opt) Then
                    Dim first As Integer = rowNo
                    For Each r In b.Rows
                        lines.Add(newLine(b.No, r.Warehouse, r.CustomerShortCode, r.SupplierShortCode,
                                          PrintName(item, r), Nothing, r.Quantity, Nothing))
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
        Dim headers As New List(Of String) From {"No", "倉庫", "得意先", "仕入先", "品名", "登録数"}
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

        xl.SetColumnsNumberFormat(P_WH, P_NAME, nfString)
        If pStock > 0 Then xl.SetColumnNumberFormat(pStock, nfInteger)
        For i As Integer = 0 To SLOTS - 1
            xl.SetColumnNumberFormat(pShelf + i * 2, nfString)
        Next

        xl.SetValue(data)

        ' 列幅：旧アプリと同じ総幅（記入欄 97/105）から倉庫列の分を引き、棚番:数量 = 1:3 で配分
        Dim widths As New List(Of Double) From {3, 4, 4, 4, 19, 3}
        If pStock > 0 Then widths.Add(8)
        Dim slotWidth As Double = If(pStock > 0, 93.0, 101.0) / SLOTS
        For i As Integer = 1 To SLOTS
            widths.Add(slotWidth * 0.25)
            widths.Add(slotWidth * 0.75)
        Next
        widths.Add(8)
        xl.SetColumnsWidth(widths.ToArray())
        xl.SetColumnHidden(pOrder)

        For Each c In {P_NO, P_WH, P_CUST, P_SUP, P_REG}
            xl.SetColumnAlign(c, xlCenter)
        Next
        xl.SetRowAlign(1, xlCenter)

        For Each s In detailSpans
            xl.SetCellsBackColor(s.First, 1, s.Last, pCols, COLOR_DETAIL)
        Next

        xl.SetAutoFilter(1, 1, lastRow, pCols)

        Dim headerText As String = opt.OfficeHeader
        Dim outputText As String = $"出力：{opt.OutputAt:yyyy/MM/dd HH:mm}"
        xl.WithSheet(PRINT_SHEET_NAME,
            Sub(sh)
                Dim all = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, pCols))
                all.RowHeight = 37
                all.ShrinkToFit = True                       ' 文字サイズをセル幅に合わせる
                all.Borders.LineStyle = xlContinuous         ' 格子
                sh.Columns(P_NAME).WrapText = True           ' 品名はセル幅で折り返し

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

    ''' <summary>印刷用の品名（通常以外のステータスは【】で付記）</summary>
    Private Shared Function PrintName(item As StocktakingItem, r As StockRow) As String
        Dim label As String = StatusLabel(r)
        Return If(label = "", item.Model, $"{item.Model}【{label}】")
    End Function

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
        Return items.SelectMany(Function(i) i.Targets).Any(Function(r) r.UnitPrice <> Decimal.Truncate(r.UnitPrice))
    End Function

    Private Shared Function Col(columnIndex As Integer) As String
        Return GetAlphabets(columnIndex)
    End Function

End Class
