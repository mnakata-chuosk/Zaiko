Option Strict Off  ' PageSetup・Outline 等の Excel COM 遅延バインディングに必要

Imports ChuoUtils
Imports ChuoUtils.ExcelTools

''' <summary>
''' 棚卸表（一覧シート・印刷用シート）を Excel に出力し、ユーザーに表示する
''' </summary>
''' <remarks>
''' 一覧シート … 帳簿との突き合わせ用。型式を親行、得意先・倉庫・ステータス別を子行とし、
'''               共用棚で数えた数量を子行の「振分数」に人が振り分ける。親行の振分数欄は振分残（0 以外で赤）。
''' 印刷用シート … 現場記入用。旧「棚卸表作成」と同じ横持ちで [棚番・数量] は画面で選んだ組数（3～5）。
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
    Private Const L_REG As Integer = 8
    Private Const L_SHELF As Integer = 9                        ' 棚番1（数量1 は +1、以降 2 列おき）
    ' 棚番列より右は組数で位置が変わる
    Private ReadOnly L_ALLOC As Integer                         ' 振分数（親行は振分残）
    Private ReadOnly L_TOTAL As Integer                         ' 棚卸総数
    Private ReadOnly L_BOOK As Integer                          ' AX在庫数
    Private ReadOnly L_PLUS As Integer                          ' ＋（売上漏れ）
    Private ReadOnly L_MINUS As Integer                         ' －（仕入漏れ）
    Private ReadOnly L_DIFF As Integer                          ' 差数
    Private ReadOnly L_PRICE As Integer                         ' 単価
    Private ReadOnly L_AMOUNT As Integer                        ' 差額
    Private ReadOnly L_NOTE As Integer                          ' 備考
    Private ReadOnly L_CUST_CODE As Integer                     ' 得意先CD（AXIS内部コード。非表示。棚番の Excel 読み込み用）
    Private ReadOnly L_COLS As Integer
    Private Const L_HEADER_ROW As Integer = 2
    Private Const L_FIRST_ROW As Integer = 3

    ' 色（旧アプリの重複明細と同じ灰色）
    Private Shared ReadOnly COLOR_CHILD As Integer = RGB(217, 217, 217)
    Private Shared ReadOnly COLOR_EXCLUDED_BACK As Integer = RGB(242, 242, 242)
    Private Shared ReadOnly COLOR_EXCLUDED_FONT As Integer = RGB(128, 128, 128)
    Private Shared ReadOnly COLOR_INPUT As Integer = RGB(255, 242, 204)
    Private Shared ReadOnly COLOR_ALERT As Integer = RGB(255, 199, 206)

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
        L_ALLOC = L_SHELF + SLOTS * 2
        L_TOTAL = L_ALLOC + 1
        L_BOOK = L_ALLOC + 2
        L_PLUS = L_ALLOC + 3
        L_MINUS = L_ALLOC + 4
        L_DIFF = L_ALLOC + 5
        L_PRICE = L_ALLOC + 6
        L_AMOUNT = L_ALLOC + 7
        L_NOTE = L_ALLOC + 8
        L_CUST_CODE = L_ALLOC + 9
        L_COLS = L_CUST_CODE
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

    ' ============================================================
    ' 一覧シート
    ' ============================================================
    Private Sub WriteListSheet(xl As ExcelObject, items As List(Of StocktakingItem), opt As StocktakingOptions)
        Dim lines As New List(Of Object())
        Dim childSpans As New List(Of RowSpan)     ' 子行（灰色・振分数入力）
        Dim excludedSpans As New List(Of RowSpan)  ' 棚卸対象外
        Dim groupSpans As New List(Of RowSpan)     ' 行グループ（各型式の2行目以降）
        Dim inputSpans As New List(Of RowSpan)     ' 振分数の入力欄（子行が2つ以上のとき）

        Dim rowNo As Integer = L_FIRST_ROW
        For Each item In items
            Dim itemFirstRow As Integer = rowNo

            If item.IsGroup Then
                Dim parentRow As Integer = rowNo
                Dim lastChildRow As Integer = parentRow + item.Targets.Count
                lines.Add(ListParentLine(item, parentRow, lastChildRow))
                rowNo += 1
                ' 子行が1つなら共用棚の数量はその得意先の分なので振分数は自動（入力欄にしない）
                Dim autoAllocRow As Integer = If(item.Targets.Count = 1, parentRow, 0)
                For Each r In item.Targets
                    lines.Add(ListChildLine(item, r, rowNo, autoAllocRow))
                    rowNo += 1
                Next
                childSpans.Add(New RowSpan(parentRow + 1, lastChildRow))
                If autoAllocRow = 0 Then inputSpans.Add(New RowSpan(parentRow + 1, lastChildRow))

            ElseIf item.Targets.Count = 1 Then
                lines.Add(ListSingleLine(item, rowNo))
                rowNo += 1
            End If

            If item.Excluded.Count > 0 Then
                Dim firstExcluded As Integer = rowNo
                For Each r In item.Excluded
                    lines.Add(ListExcludedLine(item, r, showNo:=(rowNo = itemFirstRow)))
                    rowNo += 1
                Next
                excludedSpans.Add(New RowSpan(firstExcluded, rowNo - 1))
            End If

            If rowNo - 1 > itemFirstRow Then groupSpans.Add(New RowSpan(itemFirstRow + 1, rowNo - 1))
        Next

        Dim firstAppendRow As Integer = rowNo
        For i As Integer = 1 To APPEND_ROWS
            lines.Add(ListAppendLine(rowNo))
            rowNo += 1
        Next
        Dim lastRow As Integer = rowNo - 1

        ' 1行目・2行目 + 明細を配列に詰める
        Dim data(lastRow - 1, L_COLS - 1) As Object
        data(0, L_NO - 1) = $"{opt.OutputAt:yyyy/MM/dd HH:mm} 時点"
        data(0, L_NAME - 1) = opt.OfficeLabel
        data(0, L_ALLOC - 1) = "親行=振分残"
        data(0, L_PLUS - 1) = "売上漏れ"
        data(0, L_MINUS - 1) = "仕入漏れ"
        data(0, L_AMOUNT - 1) = $"=SUBTOTAL(9,{Col(L_AMOUNT)}{L_FIRST_ROW}:{Col(L_AMOUNT)}{lastRow})"

        Dim headers As New List(Of String) From {"No", "倉庫", "得意先", "仕入先", "商品CD", "品名", "ステータス", "登録数"}
        For i As Integer = 1 To SLOTS
            headers.Add("棚番")
            headers.Add("数量")
        Next
        headers.AddRange({"振分数", "棚卸総数", "AX在庫数", "＋", "－", "差数", "単価", "差額", "備考", ListHeaders.CUSTOMER_CODE})
        For c As Integer = 0 To L_COLS - 1
            data(L_HEADER_ROW - 1, c) = headers(c)
        Next

        For i As Integer = 0 To lines.Count - 1
            For c As Integer = 0 To L_COLS - 1
                data(L_FIRST_ROW - 1 + i, c) = lines(i)(c)
            Next
        Next

        ' 表示形式（コードの先頭ゼロを残すため値の設定前に行う）
        xl.SetColumnsNumberFormat(L_WH, L_STATUS, nfString)
        For i As Integer = 0 To SLOTS - 1
            xl.SetColumnNumberFormat(L_SHELF + i * 2, nfString)
            xl.SetColumnNumberFormat(L_SHELF + i * 2 + 1, nfInteger)
        Next
        xl.SetColumnsNumberFormat(L_ALLOC, L_DIFF, nfInteger)
        xl.SetColumnNumberFormat(L_PRICE, If(HasFraction(items), nfDecimal2, nfInteger))
        xl.SetColumnNumberFormat(L_AMOUNT, nfInteger)
        xl.SetColumnNumberFormat(L_CUST_CODE, nfString)

        xl.SetValue(data)

        ' 配置・列幅
        For Each c In {L_NO, L_WH, L_CUST, L_SUP, L_STATUS, L_REG}
            xl.SetColumnAlign(c, xlCenter)
        Next
        xl.SetRowAlign(L_HEADER_ROW, xlCenter)
        xl.SetCellAlign(1, L_PLUS, xlCenter)
        xl.SetCellAlign(1, L_MINUS, xlCenter)

        Dim widths As New List(Of Double) From {5, 5.5, 5.5, 5.5, 12, 30, 8, 5}
        For i As Integer = 1 To SLOTS
            widths.Add(8)
            widths.Add(7)
        Next
        widths.AddRange({8, 8.5, 8.5, 7, 7, 7.5, 10, 12, 24, 10})
        xl.SetColumnsWidth(widths.ToArray())
        xl.SetColumnHidden(L_CUST_CODE)

        ' 子行・振分数入力欄・棚卸対象外
        For Each s In childSpans
            xl.SetCellsBackColor(s.First, 1, s.Last, L_COLS, COLOR_CHILD)
        Next
        For Each s In inputSpans
            xl.SetCellsBackColor(s.First, L_ALLOC, s.Last, L_ALLOC, COLOR_INPUT)
        Next
        For Each s In excludedSpans
            xl.SetCellsBackColor(s.First, 1, s.Last, L_COLS, COLOR_EXCLUDED_BACK)
            xl.SetCellsForeColor(s.First, 1, s.Last, L_COLS, COLOR_EXCLUDED_FONT)
        Next

        Dim sheetName As String = xl.SheetName
        xl.WithSheet(sheetName,
            Sub(sh)
                ' 親行の振分残が 0 以外なら赤（No が数値で得意先が空の行が親行）
                Dim rng = sh.Range(sh.Cells(L_FIRST_ROW, L_ALLOC), sh.Cells(lastRow, L_ALLOC))
                Dim fc = rng.FormatConditions.Add(Type:=xlExpression,
                    Formula1:=$"=AND(ISNUMBER(INDEX(${Col(L_NO)}:${Col(L_NO)},ROW())),INDEX(${Col(L_CUST)}:${Col(L_CUST)},ROW())="""",ROUND(INDEX(${Col(L_ALLOC)}:${Col(L_ALLOC)},ROW()),4)<>0)")
                fc.Interior.Color = COLOR_ALERT

                ' 追記行の枠
                sh.Range(sh.Cells(firstAppendRow, 1), sh.Cells(lastRow, L_NOTE)).Borders.LineStyle = xlContinuous

                ' 行グループのボタンを親行側に出す
                sh.Outline.SummaryRow = 0   ' xlAbove
            End Sub)

        For Each s In groupSpans
            xl.SetRowsGroup(s.First, s.Last)
        Next

        ' 商品CD と ＋/－ は折りたたみ
        xl.SetColumnGroup(L_ITEM)
        xl.SetColumnsGroup(L_PLUS, L_MINUS)
        xl.ColumnLevels = 1

        xl.SetAutoFilter(L_HEADER_ROW, 1, lastRow, L_NOTE)

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

    ''' <summary>型式・在庫行の共通項目</summary>
    Private Sub SetListRowInfo(line As Object(), item As StocktakingItem, r As StockRow)
        line(L_WH - 1) = r.Warehouse
        line(L_CUST - 1) = r.CustomerShortCode
        line(L_CUST_CODE - 1) = r.CustomerCode
        line(L_SUP - 1) = r.SupplierShortCode
        line(L_ITEM - 1) = item.ItemCode
        line(L_NAME - 1) = item.Model
        line(L_STATUS - 1) = StatusLabel(r)
        line(L_BOOK - 1) = CDbl(r.Quantity)
        line(L_PRICE - 1) = CDbl(r.UnitPrice)
    End Sub

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

    ''' <summary>差数・差額の数式</summary>
    Private Sub SetDiffFormulas(line As Object(), row As Integer)
        line(L_DIFF - 1) = $"={Col(L_TOTAL)}{row}-{Col(L_BOOK)}{row}+{Col(L_PLUS)}{row}-{Col(L_MINUS)}{row}"
        line(L_AMOUNT - 1) = $"={Col(L_DIFF)}{row}*{Col(L_PRICE)}{row}"
    End Sub

    ''' <summary>単独行（棚卸対象が1行だけの型式）</summary>
    Private Function ListSingleLine(item As StocktakingItem, row As Integer) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = item.No
        SetListRowInfo(line, item, item.Targets(0))
        line(L_REG - 1) = 1
        SetListShelves(line, item.SingleRowShelves())
        line(L_TOTAL - 1) = $"={QtySum(row)}"
        SetDiffFormulas(line, row)
        Return line
    End Function

    ''' <summary>親行（共用棚の数量を記入。振分数欄は振分残）</summary>
    Private Function ListParentLine(item As StocktakingItem, row As Integer, lastChildRow As Integer) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = item.No
        line(L_WH - 1) = item.CommonValue(Function(r) r.Warehouse)
        line(L_SUP - 1) = item.CommonValue(Function(r) r.SupplierShortCode)
        line(L_ITEM - 1) = item.ItemCode
        line(L_NAME - 1) = item.Model
        line(L_REG - 1) = item.Targets.Count
        SetListShelves(line, item.SharedShelves)
        line(L_ALLOC - 1) = $"={Col(L_TOTAL)}{row}-SUM({Col(L_ALLOC)}{row + 1}:{Col(L_ALLOC)}{lastChildRow})"
        line(L_TOTAL - 1) = $"={QtySum(row)}"
        line(L_BOOK - 1) = $"=SUM({Col(L_BOOK)}{row + 1}:{Col(L_BOOK)}{lastChildRow})"
        If line(L_NOTE - 1) Is Nothing Then line(L_NOTE - 1) = "共用棚の数量を子行の振分数へ"
        Return line
    End Function

    ''' <summary>子行（振分数＋得意先指定棚の数量が棚卸総数）</summary>
    ''' <param name="autoAllocRow">0 以外なら振分数をこの行（親行）の棚卸総数にする（子行が1つのとき）</param>
    Private Function ListChildLine(item As StocktakingItem, r As StockRow, row As Integer, autoAllocRow As Integer) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = ListHeaders.CHILD_NO
        SetListRowInfo(line, item, r)
        SetListShelves(line, item.ShelvesOf(r))
        If autoAllocRow > 0 Then line(L_ALLOC - 1) = $"={Col(L_TOTAL)}{autoAllocRow}"
        line(L_TOTAL - 1) = $"={Col(L_ALLOC)}{row}+{QtySum(row)}"
        SetDiffFormulas(line, row)
        Return line
    End Function

    ''' <summary>棚卸対象外ステータスの行（参考表示のみ、数式なし）</summary>
    Private Function ListExcludedLine(item As StocktakingItem, r As StockRow, showNo As Boolean) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = If(showNo, CObj(item.No), ListHeaders.CHILD_NO)
        SetListRowInfo(line, item, r)
        line(L_NOTE - 1) = ListHeaders.EXCLUDED_NOTE
        Return line
    End Function

    ''' <summary>追記用の空行（数式のみ）</summary>
    Private Function ListAppendLine(row As Integer) As Object()
        Dim line = NewListLine()
        line(L_NO - 1) = APPEND_LABEL
        line(L_TOTAL - 1) = $"={QtySum(row)}"
        SetDiffFormulas(line, row)
        Return line
    End Function

    ''' <summary>数量列の合計式（SUM(J3,L3,...)）</summary>
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
        Dim pCols As Integer = pShelf - 1 + SLOTS * 2

        Dim lines As New List(Of Object())
        Dim childSpans As New List(Of RowSpan)

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
                          For Each kv In shelves.Where(Function(x) x.Key <= SLOTS)
                              line(pShelf - 1 + (kv.Key - 1) * 2) = kv.Value
                          Next
                          Return line
                      End Function

        Dim rowNo As Integer = 2
        For Each item In items
            If item.IsGroup Then
                lines.Add(newLine(item.No, item.CommonValue(Function(x) x.Warehouse), "",
                                  item.CommonValue(Function(x) x.SupplierShortCode),
                                  item.Model, item.Targets.Count, item.TargetQuantity, item.SharedShelves))
                rowNo += 1
                Dim firstChild As Integer = rowNo
                For Each r In item.Targets
                    lines.Add(newLine(item.No, r.Warehouse, r.CustomerShortCode, r.SupplierShortCode,
                                      PrintName(item, r), Nothing, r.Quantity, item.ShelvesOf(r)))
                    rowNo += 1
                Next
                childSpans.Add(New RowSpan(firstChild, rowNo - 1))

            ElseIf item.Targets.Count = 1 Then
                Dim r = item.Targets(0)
                lines.Add(newLine(item.No, r.Warehouse, r.CustomerShortCode, r.SupplierShortCode,
                                  PrintName(item, r), 1, r.Quantity, item.SingleRowShelves()))
                rowNo += 1
            End If
        Next

        For i As Integer = 1 To APPEND_ROWS
            Dim line(pCols - 1) As Object
            line(P_NO - 1) = APPEND_LABEL
            lines.Add(line)
            rowNo += 1
        Next
        Dim lastRow As Integer = rowNo - 1

        Dim data(lastRow - 1, pCols - 1) As Object
        Dim headers As New List(Of String) From {"No", "倉庫", "得意先", "仕入先", "品名", "登録数"}
        If pStock > 0 Then headers.Add("在庫数")
        For i As Integer = 1 To SLOTS
            headers.Add("棚番")
            headers.Add("数量")
        Next
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
        xl.SetColumnsWidth(widths.ToArray())

        For Each c In {P_NO, P_WH, P_CUST, P_SUP, P_REG}
            xl.SetColumnAlign(c, xlCenter)
        Next
        xl.SetRowAlign(1, xlCenter)

        For Each s In childSpans
            xl.SetCellsBackColor(s.First, 1, s.Last, pCols, COLOR_CHILD)
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
        Return items.SelectMany(Function(i) i.Targets.Concat(i.Excluded)) _
                    .Any(Function(r) r.UnitPrice <> Decimal.Truncate(r.UnitPrice))
    End Function

    Private Shared Function Col(columnIndex As Integer) As String
        Return GetAlphabets(columnIndex)
    End Function

End Class
