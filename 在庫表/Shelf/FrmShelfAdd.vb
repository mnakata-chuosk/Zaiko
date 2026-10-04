''' <summary>
''' 棚番編集の行追加。型式（完全一致）で商品を特定し、共用棚か得意先指定（取扱得意先から選択）かを選ぶ。
''' </summary>
Public Class FrmShelfAdd

    Private Const SHARED_LABEL As String = "（共用棚）"

    Private ReadOnly _officeCode As String
    Private _items As New Dictionary(Of String, String)
    Private _customers As New List(Of ShelfRepository.Customer)

    Public ReadOnly Property ItemCode As String
    Public ReadOnly Property Model As String
    ''' <summary>選択した得意先（共用棚は Nothing）</summary>
    Public ReadOnly Property SelectedCustomer As ShelfRepository.Customer

    Public Sub New(officeCode As String)
        InitializeComponent()
        _officeCode = officeCode
        UpdateButtons()
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
        CboCustomer.Items.Clear()
        If _items.Count = 0 Then
            MsgBox($"型式「{model}」は商品マスタに見つかりません（完全一致で検索します）。", MsgBoxStyle.Exclamation, Me.Text)
        Else
            For Each kv In _items
                CboItem.Items.Add($"{kv.Key}：{kv.Value}")
            Next
            CboItem.SelectedIndex = 0
            Me.AcceptButton = BtnOk
        End If
        UpdateButtons()
    End Sub

    Private Sub CboItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboItem.SelectedIndexChanged
        CboCustomer.Items.Clear()
        If CboItem.SelectedIndex < 0 Then Return

        Me.Cursor = Cursors.WaitCursor
        Try
            _customers = ShelfRepository.LoadCustomers(_officeCode, _items.Keys.ElementAt(CboItem.SelectedIndex))
        Finally
            Me.Cursor = Cursors.Default
        End Try

        CboCustomer.Items.Add(SHARED_LABEL)
        For Each c In _customers
            CboCustomer.Items.Add($"{c.ShortCode} {c.Name}")
        Next
        CboCustomer.SelectedIndex = 0
        UpdateButtons()
    End Sub

    Private Sub UpdateButtons()
        BtnOk.Enabled = CboItem.SelectedIndex >= 0 AndAlso CboCustomer.SelectedIndex >= 0
    End Sub

    Private Sub TxtModel_TextChanged(sender As Object, e As EventArgs) Handles TxtModel.TextChanged
        ' 型式を打ち直したら再検索させる
        Me.AcceptButton = BtnFind
    End Sub

    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        Dim key = _items.Keys.ElementAt(CboItem.SelectedIndex)
        _ItemCode = key
        _Model = _items(key)
        _SelectedCustomer = If(CboCustomer.SelectedIndex <= 0, Nothing, _customers(CboCustomer.SelectedIndex - 1))
        DialogResult = DialogResult.OK
    End Sub

End Class
