Imports ChuoUtils

''' <summary>
''' 棚番編集画面。営業所の棚番を 型式 × 得意先指定 の行で横並び（棚番1～5）に表示・編集し、AXIS に保存する。
''' </summary>
Public Class FrmShelf

    Private Const APP_NAME As String = "在庫表"
    Private Const SLOTS As Integer = ShelfGroup.EDIT_SLOTS

    ' DataTable の列名
    Private Const COL_KEY As String = "Key"
    Private Const COL_MODEL As String = "型式"
    Private Const COL_CUST As String = "得意先"
    Private Const COL_CUST_NAME As String = "得意先名"
    Private Const COL_STOCK As String = "在庫数"
    Private Const COL_SHELF As String = "棚番"   ' + 1～5
    Private Const COL_EXTRA As String = "他"
    Private Const COL_CHANGED As String = "変更"
    Private Const COL_HAS_STOCK As String = "在庫あり"

    Private Shared ReadOnly COLOR_CHANGED As Color = Color.FromArgb(255, 242, 204)
    Private Shared ReadOnly COLOR_DUPLICATE As Color = Color.FromArgb(255, 199, 206)
    Private Shared ReadOnly COLOR_READONLY As Color = Color.FromArgb(242, 242, 242)

    Private ReadOnly _userId As String
    Private _officeIndex As Integer = -1
    Private _groups As New Dictionary(Of String, ShelfGroup)
    Private ReadOnly _table As New DataTable
    Private _suppressOfficeChange As Boolean
    ''' <summary>プログラムから行を書き換え中（CellValueChanged を無視する）</summary>
    Private _rendering As Boolean

    ''' <param name="officeIndex">初期表示する営業所（Chuo.OfficeList のインデックス）</param>
    ''' <param name="userId">更新担当者として記録する担当者ID</param>
    Public Sub New(officeIndex As Integer, userId As String)
        InitializeComponent()
        _userId = userId

        BuildTable()
        SetupColumns()
        Dgv.DataSource = _table.DefaultView

        _suppressOfficeChange = True
        CboOffice.Items.AddRange(Chuo.OfficeList)
        CboOffice.SelectedIndex = Math.Max(0, officeIndex)
        _suppressOfficeChange = False
    End Sub

    Private ReadOnly Property OfficeCode As String
        Get
            Return Chuo.OfficeCDList(_officeIndex)
        End Get
    End Property

    Private Sub FrmShelf_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        LoadOffice(CboOffice.SelectedIndex)
    End Sub

    ' ============================================================
    ' 表示
    ' ============================================================
    Private Sub BuildTable()
        _table.Columns.Add(COL_KEY, GetType(String))
        _table.Columns.Add(COL_MODEL, GetType(String))
        _table.Columns.Add(COL_CUST, GetType(String))
        _table.Columns.Add(COL_CUST_NAME, GetType(String))
        _table.Columns.Add(COL_STOCK, GetType(Decimal))
        For i As Integer = 1 To SLOTS
            _table.Columns.Add(COL_SHELF & i, GetType(String))
        Next
        _table.Columns.Add(COL_EXTRA, GetType(String))
        _table.Columns.Add(COL_CHANGED, GetType(Boolean))
        _table.Columns.Add(COL_HAS_STOCK, GetType(Boolean))
        _table.PrimaryKey = {_table.Columns(COL_KEY)}
    End Sub

    ''' <summary>表示列を定義する（Key・変更・在庫あり はフィルタ用で表示しない）</summary>
    Private Sub SetupColumns()
        Dgv.AutoGenerateColumns = False
        Dgv.Columns.Clear()

        Dim add = Function(name As String, width As Integer, isReadOnly As Boolean) As DataGridViewTextBoxColumn
                      Dim c As New DataGridViewTextBoxColumn With {
                          .Name = name, .DataPropertyName = name, .HeaderText = name,
                          .Width = width, .ReadOnly = isReadOnly,
                          .SortMode = DataGridViewColumnSortMode.NotSortable
                      }
                      If isReadOnly Then c.DefaultCellStyle.BackColor = COLOR_READONLY
                      Dgv.Columns.Add(c)
                      Return c
                  End Function

        add(COL_MODEL, 260, True).Frozen = True
        add(COL_CUST, 64, True)
        add(COL_CUST_NAME, 150, True)
        With add(COL_STOCK, 80, True).DefaultCellStyle
            .Format = "#,##0"
            .Alignment = DataGridViewContentAlignment.MiddleRight
        End With
        For i As Integer = 1 To SLOTS
            add(COL_SHELF & i, 90, False).MaxInputLength = 200
        Next
        add(COL_EXTRA, 56, True).ToolTipText = $"棚番{SLOTS + 1}以降に登録されている棚番の件数（保存しても変更されません）"
    End Sub

    ''' <summary>営業所を読み込む（DB から取り直し）</summary>
    Private Sub LoadOffice(index As Integer)
        Me.Cursor = Cursors.WaitCursor
        Try
            _officeIndex = index
            Dim groups = ShelfRepository.LoadGroups(OfficeCode)

            _groups = groups.ToDictionary(Function(g) g.Key)
            _table.BeginLoadData()
            _table.Rows.Clear()
            For Each g In groups
                Dim row = _table.NewRow()
                FillRow(row, g)
                _table.Rows.Add(row)
            Next
            _table.EndLoadData()

            Me.Text = $"棚番編集 - {CboOffice.Text}営業所"
            ApplyFilter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, APP_NAME)
            App.WriteErrLog(APP_NAME, ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FillRow(row As DataRow, g As ShelfGroup)
        Dim map = g.CurrentMap
        row(COL_KEY) = g.Key
        row(COL_MODEL) = g.Model
        row(COL_CUST) = If(g.IsShared, "", g.CustomerShortCode)
        row(COL_CUST_NAME) = If(g.IsShared, "（共用棚）", g.CustomerName)
        row(COL_STOCK) = g.StockQuantity
        For pos As Integer = 1 To SLOTS
            Dim v As String = Nothing
            row(COL_SHELF & pos) = If(map.TryGetValue(pos, v), v, "")
        Next
        row(COL_EXTRA) = If(g.ExtraCount > 0, $"+{g.ExtraCount}", "")
        row(COL_CHANGED) = g.IsChanged OrElse IsNewGroup(g)
        row(COL_HAS_STOCK) = g.StockQuantity <> 0
    End Sub

    ''' <summary>まだ DB に1件も無いグループ（行追加・Excel 読み込みで追加した行）</summary>
    Private Shared Function IsNewGroup(g As ShelfGroup) As Boolean
        Return g.Records.Count = 0 AndAlso g.CurrentMap.Count > 0
    End Function

    Private Sub ApplyFilter()
        Dim conds As New List(Of String)
        Dim text = TxtSearch.Text.Trim()
        If text <> "" Then conds.Add($"{COL_MODEL} LIKE '%{EscapeLike(text)}%'")
        If ChkStockOnly.Checked Then conds.Add($"{COL_HAS_STOCK} = True")
        If ChkChangedOnly.Checked Then conds.Add($"{COL_CHANGED} = True")
        _table.DefaultView.RowFilter = String.Join(" AND ", conds)
        UpdateStatus()
    End Sub

    ''' <summary>DataView の LIKE 用にワイルドカード・引用符をエスケープする</summary>
    Private Shared Function EscapeLike(value As String) As String
        Dim sb As New System.Text.StringBuilder()
        For Each ch In value
            Select Case ch
                Case "*"c, "%"c, "["c, "]"c
                    sb.Append("[").Append(ch).Append("]")
                Case "'"c
                    sb.Append("''")
                Case Else
                    sb.Append(ch)
            End Select
        Next
        Return sb.ToString()
    End Function

    Private Sub UpdateStatus()
        Dim changed = ChangedGroups().Count
        LblStatus.Text = $"{_table.DefaultView.Count:#,0} 件表示 / 全 {_table.Rows.Count:#,0} 件　　未保存の変更 {changed} 件"
        BtnSave.Enabled = changed > 0
    End Sub

    Private Function ChangedGroups() As List(Of ShelfGroup)
        Return _groups.Values.Where(Function(g) g.IsChanged OrElse IsNewGroup(g)).ToList()
    End Function

    Private Sub TxtSearch_TextChanged(sender As Object, e As EventArgs) Handles TxtSearch.TextChanged
        ApplyFilter()
    End Sub

    Private Sub Filter_CheckedChanged(sender As Object, e As EventArgs) Handles ChkStockOnly.CheckedChanged, ChkChangedOnly.CheckedChanged
        ApplyFilter()
    End Sub

    ' ============================================================
    ' 編集
    ' ============================================================
    Private Function GroupOf(rowIndex As Integer) As ShelfGroup
        If rowIndex < 0 OrElse rowIndex >= Dgv.Rows.Count Then Return Nothing
        Dim view = TryCast(Dgv.Rows(rowIndex).DataBoundItem, DataRowView)
        If view Is Nothing Then Return Nothing
        Dim g As ShelfGroup = Nothing
        _groups.TryGetValue(CStr(view(COL_KEY)), g)
        Return g
    End Function

    Private Sub Dgv_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv.CellValueChanged
        If _rendering OrElse e.RowIndex < 0 OrElse Not Dgv.Columns(e.ColumnIndex).Name.StartsWith(COL_SHELF) Then Return
        Dim view = TryCast(Dgv.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If view Is Nothing Then Return
        Dim key = CStr(view(COL_KEY))

        ' 行の表示（変更フラグ・他件数）を書き換えるため、編集の確定後に行う
        BeginInvoke(New Action(Sub() ApplyRowEdit(key)))
    End Sub

    Private Sub ApplyRowEdit(key As String)
        Dim row = _table.Rows.Find(key)
        Dim g As ShelfGroup = Nothing
        If row Is Nothing OrElse Not _groups.TryGetValue(key, g) Then Return

        g.ApplySlots(Enumerable.Range(1, SLOTS).Select(Function(i) TryCast(row(COL_SHELF & i), String)).ToList())
        _rendering = True
        Try
            FillRow(row, g)
        Finally
            _rendering = False
        End Try
        UpdateStatus()
    End Sub

    Private Sub Dgv_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles Dgv.CellFormatting
        If e.RowIndex < 0 OrElse Not Dgv.Columns(e.ColumnIndex).Name.StartsWith(COL_SHELF) Then Return
        Dim g = GroupOf(e.RowIndex)
        If g Is Nothing Then Return

        Dim pos = CInt(Dgv.Columns(e.ColumnIndex).Name.Substring(COL_SHELF.Length))
        Dim value = If(TryCast(e.Value, String), "")
        Dim originalValue As String = Nothing
        If Not g.OriginalMap.TryGetValue(pos, originalValue) Then originalValue = ""

        If value <> "" AndAlso g.CurrentMap.Values.Where(Function(s) s = value).Count() > 1 Then
            e.CellStyle.BackColor = COLOR_DUPLICATE
        ElseIf value <> originalValue Then
            e.CellStyle.BackColor = COLOR_CHANGED
        End If
    End Sub

    Private Sub Dgv_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Dgv.DataError
        e.Cancel = True
    End Sub

    ' ============================================================
    ' 行追加
    ' ============================================================
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        Dim dlg As FrmShelfAdd
        Try
            dlg = New FrmShelfAdd(OfficeCode)   ' 得意先一覧を DB から読み込む
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, APP_NAME)
            App.WriteErrLog(APP_NAME, ex.ToString)
            Return
        End Try

        Using dlg
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return

            Dim g = New ShelfGroup With {
                .ItemCode = dlg.ItemCode,
                .Model = dlg.Model,
                .CustomerCode = If(dlg.SelectedCustomer?.Code, ""),
                .CustomerShortCode = If(dlg.SelectedCustomer?.ShortCode, ""),
                .CustomerName = If(dlg.SelectedCustomer?.Name, "")
            }
            If Not _groups.ContainsKey(g.Key) Then
                AddGroup(g)
            End If

            ' 追加した（または既にある）行を表示して棚番1を編集
            ChkChangedOnly.Checked = False
            ChkStockOnly.Checked = False
            TxtSearch.Text = g.Model
            FocusRow(g.Key)
        End Using
    End Sub

    Private Sub AddGroup(g As ShelfGroup)
        _groups.Add(g.Key, g)
        Dim row = _table.NewRow()
        FillRow(row, g)
        _table.Rows.Add(row)
    End Sub

    Private Sub FocusRow(key As String)
        For Each r As DataGridViewRow In Dgv.Rows
            Dim view = TryCast(r.DataBoundItem, DataRowView)
            If view IsNot Nothing AndAlso CStr(view(COL_KEY)) = key Then
                Dgv.CurrentCell = r.Cells(COL_SHELF & 1)
                Dgv.Focus()
                Return
            End If
        Next
    End Sub

    ' ============================================================
    ' Excel 読み込み
    ' ============================================================
    Private Sub BtnImport_Click(sender As Object, e As EventArgs) Handles BtnImport.Click
        Dim path As String
        Using ofd As New OpenFileDialog With {
            .Title = "棚卸表（本アプリで出力した Excel）を選択",
            .Filter = "Excel ブック (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
        }
            If ofd.ShowDialog(Me) <> DialogResult.OK Then Return
            path = ofd.FileName
        End Using

        Dim changes As List(Of ShelfImportChange)
        Dim planner As ShelfImportPlanner
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim file = ShelfExcelImporter.Read(path)
            If file.OfficeCode <> OfficeCode Then
                Me.Cursor = Cursors.Default
                MsgBox($"選択したファイルは「{file.SheetName}」です。{vbCrLf}営業所を切り替えてから読み込んでください。", MsgBoxStyle.Exclamation, APP_NAME)
                Return
            End If

            planner = New ShelfImportPlanner(OfficeCode, _groups.Values)
            changes = planner.Plan(file)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, APP_NAME)
            App.WriteErrLog(APP_NAME, ex.ToString)
            Return
        Finally
            Me.Cursor = Cursors.Default
        End Try

        If changes.Count = 0 AndAlso planner.Errors.Count = 0 Then
            MsgBox("棚番の変更はありませんでした。", MsgBoxStyle.Information, APP_NAME)
            Return
        End If

        Using dlg As New FrmShelfImport(changes, planner.Errors)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return

            _rendering = True
            Try
                For Each c In dlg.SelectedChanges
                    If Not _groups.ContainsKey(c.Group.Key) Then AddGroup(c.Group)
                    c.Group.PendingMap = c.After
                    FillRow(_table.Rows.Find(c.Group.Key), c.Group)
                Next
            Finally
                _rendering = False
            End Try
        End Using

        TxtSearch.Text = ""
        ChkStockOnly.Checked = False
        ChkChangedOnly.Checked = True
        ApplyFilter()
    End Sub

    ' ============================================================
    ' 保存
    ' ============================================================
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        Dgv.EndEdit()
        Dim changed = ChangedGroups()
        If changed.Count = 0 Then Return

        If String.IsNullOrEmpty(_userId) Then
            MsgBox("担当者が設定されていないため保存できません。メイン画面で担当者を選択してください。", MsgBoxStyle.Exclamation, APP_NAME)
            Return
        End If

        Dim duplicates = changed.Where(Function(g) g.CurrentMap.Values.Distinct().Count() <> g.CurrentMap.Count).ToList()
        Dim msg = $"{CboOffice.Text}営業所の棚番を {changed.Count} 件（型式×得意先）更新します。よろしいですか？"
        If duplicates.Count > 0 Then
            msg = $"同じ棚番が重複している行が {duplicates.Count} 件あります（{String.Join("、", duplicates.Take(5).Select(Function(g) g.Model))}）。{vbCrLf}{vbCrLf}" & msg
        End If
        If MsgBox(msg, MsgBoxStyle.Question Or MsgBoxStyle.OkCancel, APP_NAME) <> MsgBoxResult.Ok Then Return

        Dim saved As Integer = 0
        Dim conflicts As New List(Of String)
        Dim failures As New List(Of String)

        Me.Cursor = Cursors.WaitCursor
        Try
            For Each g In changed
                Try
                    If ShelfRepository.Save(OfficeCode, _userId, g) = ShelfRepository.SaveOutcome.Saved Then
                        saved += 1
                    Else
                        conflicts.Add(Describe(g))
                    End If
                Catch ex As Exception
                    failures.Add($"{Describe(g)}：{ex.Message}")
                    App.WriteErrLog(APP_NAME, $"棚番保存エラー {OfficeCode} {g.Key}{vbCrLf}{ex}")
                End Try
            Next
            App.AddCount(APP_NAME, "棚番編集", $"{OfficeCode} 保存{saved} 競合{conflicts.Count} 失敗{failures.Count}")
        Finally
            Me.Cursor = Cursors.Default
        End Try

        Dim report As String = $"{saved} 件保存しました。"
        If conflicts.Count > 0 Then
            report &= $"{vbCrLf}{vbCrLf}読み込み後に AXIS 側で更新されていたため、次の {conflicts.Count} 件は保存していません。最新の内容を確認して入力し直してください。{vbCrLf}" &
                      String.Join(vbCrLf, conflicts.Take(20))
        End If
        If failures.Count > 0 Then
            report &= $"{vbCrLf}{vbCrLf}次の {failures.Count} 件は保存に失敗しました。{vbCrLf}" & String.Join(vbCrLf, failures.Take(20))
        End If
        MsgBox(report, If(conflicts.Count + failures.Count > 0, MsgBoxStyle.Exclamation, MsgBoxStyle.Information), APP_NAME)

        ' 保存できなかった分は破棄して最新を表示する（競合した行は入力し直してもらう）
        LoadOffice(_officeIndex)
    End Sub

    Private Shared Function Describe(g As ShelfGroup) As String
        Return If(g.IsShared, g.Model, $"{g.Model}（{g.CustomerShortCode}）")
    End Function

    ' ============================================================
    ' 未保存の変更の確認
    ' ============================================================
    Private Function ConfirmDiscard() As Boolean
        Dgv.EndEdit()
        Dim changed = ChangedGroups().Count
        If changed = 0 Then Return True
        Return MsgBox($"未保存の変更が {changed} 件あります。破棄してよろしいですか？",
                      MsgBoxStyle.Question Or MsgBoxStyle.OkCancel, APP_NAME) = MsgBoxResult.Ok
    End Function

    Private Sub CboOffice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboOffice.SelectedIndexChanged
        If _suppressOfficeChange OrElse _officeIndex < 0 OrElse CboOffice.SelectedIndex = _officeIndex Then Return
        If Not ConfirmDiscard() Then
            _suppressOfficeChange = True
            CboOffice.SelectedIndex = _officeIndex
            _suppressOfficeChange = False
            Return
        End If
        LoadOffice(CboOffice.SelectedIndex)
    End Sub

    Private Sub BtnReload_Click(sender As Object, e As EventArgs) Handles BtnReload.Click
        If ConfirmDiscard() Then LoadOffice(_officeIndex)
    End Sub

    Private Sub FrmShelf_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not ConfirmDiscard() Then e.Cancel = True
    End Sub

End Class
