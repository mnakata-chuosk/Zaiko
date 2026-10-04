''' <summary>
''' 棚番編集の行追加。型式（完全一致）で商品を特定し、共用棚か得意先指定かを選ぶ。
''' 得意先は受発注管理と同じく「短縮CD 略称」のドロップダウンで、短縮CDを打つと候補が絞られる。
''' </summary>
Public Class FrmShelfAdd

    Private Const SHARED_LABEL As String = "（共用棚）"
    Private Const COL_CODE As String = "取引先CD"
    Private Const COL_NAME As String = "名称"

    Private ReadOnly _officeCode As String
    Private _items As New Dictionary(Of String, String)
    Private ReadOnly _customers As New Dictionary(Of String, ShelfRepository.Customer)

    Public ReadOnly Property ItemCode As String
    Public ReadOnly Property Model As String
    ''' <summary>選択した得意先（共用棚は Nothing）</summary>
    Public ReadOnly Property SelectedCustomer As ShelfRepository.Customer

    Public Sub New(officeCode As String)
        InitializeComponent()
        _officeCode = officeCode
        LoadCustomers()
        UpdateButtons()
    End Sub

    ''' <summary>営業所の得意先を「短縮CD 略称」で並べる（先頭は共用棚）</summary>
    Private Sub LoadCustomers()
        Dim dt As New DataTable
        dt.Columns.Add(COL_CODE, GetType(String))
        dt.Columns.Add(COL_NAME, GetType(String))
        dt.Rows.Add("", SHARED_LABEL)

        For Each c In ShelfRepository.LoadOfficeCustomers(_officeCode)
            If _customers.ContainsKey(c.Code) Then Continue For
            _customers.Add(c.Code, c)
            dt.Rows.Add(c.Code, $"{c.ShortCode} {c.Name}".Trim())
        Next

        CboCustomer.DisplayMember = COL_NAME
        CboCustomer.ValueMember = COL_CODE
        CboCustomer.DataSource = dt
        CboCustomer.SelectedIndex = 0
    End Sub

    Private Sub BtnFind_Click(sender As Object, e As EventArgs) Handles BtnFind.Click
        Dim model = TxtModel.Text.Trim()
        If model = "" Then Return

        Me.Cursor = Cursors.WaitCursor
        Try
            _items = ShelfRepository.FindItemsByModel(model)
        Finally
            Me.Cursor = Cursors.Default
        End Try

        CboItem.Items.Clear()
        If _items.Count = 0 Then
            MsgBox($"型式「{model}」は商品マスタに見つかりません（完全一致で検索します）。", MsgBoxStyle.Exclamation, Me.Text)
        Else
            For Each kv In _items
                CboItem.Items.Add($"{kv.Key}：{kv.Value}")
            Next
            CboItem.SelectedIndex = 0
            CboCustomer.Focus()
            Me.AcceptButton = BtnOk
        End If
        UpdateButtons()
    End Sub

    Private Sub UpdateButtons()
        BtnOk.Enabled = CboItem.SelectedIndex >= 0
    End Sub

    Private Sub TxtModel_TextChanged(sender As Object, e As EventArgs) Handles TxtModel.TextChanged
        ' 型式を打ち直したら再検索させる
        Me.AcceptButton = BtnFind
    End Sub

    ''' <summary>
    ''' 入力・選択された得意先を特定する。空欄・共用棚は Nothing（共用棚）。
    ''' 一覧から選ばずに打ち込んだ場合は「短縮CD」または「短縮CD 略称」と一致する得意先を探す。
    ''' </summary>
    Private Function ResolveCustomer(ByRef found As Boolean) As ShelfRepository.Customer
        found = True
        Dim code = TryCast(CboCustomer.SelectedValue, String)
        If CboCustomer.SelectedIndex > 0 AndAlso Not String.IsNullOrEmpty(code) Then Return _customers(code)

        Dim text = CboCustomer.Text.Trim()
        If text = "" OrElse text = SHARED_LABEL Then Return Nothing

        Dim matches = _customers.Values.Where(Function(c) c.ShortCode = text OrElse $"{c.ShortCode} {c.Name}".Trim() = text).ToList()
        If matches.Count = 1 Then Return matches(0)

        found = False
        Return Nothing
    End Function

    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        Dim found As Boolean
        Dim cust = ResolveCustomer(found)
        If Not found Then
            MsgBox($"得意先「{CboCustomer.Text}」が見つかりません。一覧から選択してください。", MsgBoxStyle.Exclamation, Me.Text)
            CboCustomer.Focus()
            Return
        End If

        Dim key = _items.Keys.ElementAt(CboItem.SelectedIndex)
        _ItemCode = key
        _Model = _items(key)
        _SelectedCustomer = cust
        DialogResult = DialogResult.OK
    End Sub

    ' 得意先欄で Enter を押したら追加（受発注管理の得意先選択と同じ操作感）
    Private Sub CboCustomer_KeyDown(sender As Object, e As KeyEventArgs) Handles CboCustomer.KeyDown
        If e.KeyCode = Keys.Enter AndAlso BtnOk.Enabled Then
            e.SuppressKeyPress = True
            BtnOk.PerformClick()
        End If
    End Sub

End Class
