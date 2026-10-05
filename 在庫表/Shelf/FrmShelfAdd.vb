''' <summary>
''' 棚番編集の行追加。型式と得意先を入れて「追加」を押すと、型式（完全一致）で商品を特定して行を追加する。
''' ・得意先は受発注管理と同じく「短縮CD 略称」のドロップダウンで、短縮CDを打つと候補が絞られる（（共用棚）は共用棚）
''' ・型式・得意先は営業所の取扱得意先（m商品取扱得意先明細）で相互に絞り込む
'''   型式を入れる → その型式の取扱得意先だけ、得意先を選ぶ → その得意先の型式だけ
''' ・同じ型式の商品が複数あるときだけ「商品」欄を出して選ばせる
''' </summary>
Public Class FrmShelfAdd

    Private Const SHARED_LABEL As String = "（共用棚）"
    Private Const COL_CODE As String = "取引先CD"
    Private Const COL_NAME As String = "名称"

    Private ReadOnly _officeCode As String
    ''' <summary>営業所の得意先（キー: 得意先CD）</summary>
    Private ReadOnly _customers As New Dictionary(Of String, ShelfRepository.Customer)
    ''' <summary>型式 → 取扱得意先CD</summary>
    Private ReadOnly _customersByModel As New Dictionary(Of String, HashSet(Of String))
    ''' <summary>得意先CD → 取扱型式</summary>
    Private ReadOnly _modelsByCustomer As New Dictionary(Of String, SortedSet(Of String))
    Private _allModels As String()

    ''' <summary>得意先の候補を絞り込んでいる型式（Nothing は絞り込みなし）</summary>
    Private _customerFilterModel As String
    ''' <summary>型式の候補を絞り込んでいる得意先CD（Nothing は絞り込みなし）</summary>
    Private _modelFilterCustomer As String
    ''' <summary>プログラムから候補を入れ替え中</summary>
    Private _updating As Boolean

    ''' <summary>検索した型式と結果（キー: 商品CD、値: 型式）</summary>
    Private _searchedModel As String
    Private _items As New Dictionary(Of String, String)

    Public ReadOnly Property ItemCode As String
    Public ReadOnly Property Model As String
    ''' <summary>選択した得意先（共用棚は Nothing）</summary>
    Public ReadOnly Property SelectedCustomer As ShelfRepository.Customer

    Public Sub New(officeCode As String)
        InitializeComponent()
        _officeCode = officeCode
        LoadMasters()
        ShowItemRow(False)
    End Sub

    ' ============================================================
    ' 候補の読み込み・絞り込み
    ' ============================================================

    ''' <summary>営業所の得意先と取扱得意先を読み込み、候補を初期化する</summary>
    Private Sub LoadMasters()
        For Each c In ShelfRepository.LoadOfficeCustomers(_officeCode)
            If Not _customers.ContainsKey(c.Code) Then _customers.Add(c.Code, c)
        Next

        For Each h In ShelfRepository.LoadOfficeHandling(_officeCode)
            If Not _customers.ContainsKey(h.Customer.Code) Then _customers.Add(h.Customer.Code, h.Customer)

            If Not _customersByModel.ContainsKey(h.Model) Then _customersByModel(h.Model) = New HashSet(Of String)
            _customersByModel(h.Model).Add(h.Customer.Code)

            If Not _modelsByCustomer.ContainsKey(h.Customer.Code) Then _modelsByCustomer(h.Customer.Code) = New SortedSet(Of String)(StringComparer.Ordinal)
            _modelsByCustomer(h.Customer.Code).Add(h.Model)
        Next
        _allModels = _customersByModel.Keys.OrderBy(Function(m) m, StringComparer.Ordinal).ToArray()

        _updating = True
        Try
            CboCustomer.DisplayMember = COL_NAME
            CboCustomer.ValueMember = COL_CODE
            SetCustomerCandidates(_customers.Values)
            SetModelCandidates(_allModels)
        Finally
            _updating = False
        End Try
    End Sub

    ''' <summary>得意先の候補を入れ替える（選択中の得意先が候補に残ればそのまま）</summary>
    Private Sub SetCustomerCandidates(customers As IEnumerable(Of ShelfRepository.Customer))
        Dim selected = SelectedCustomerCode()

        Dim dt As New DataTable
        dt.Columns.Add(COL_CODE, GetType(String))
        dt.Columns.Add(COL_NAME, GetType(String))
        dt.Rows.Add("", SHARED_LABEL)
        For Each c In customers.OrderBy(Function(x) x.ShortCode, StringComparer.Ordinal).ThenBy(Function(x) x.Code, StringComparer.Ordinal)
            dt.Rows.Add(c.Code, $"{c.ShortCode} {c.Name}".Trim())
        Next

        CboCustomer.DataSource = dt
        Dim index = If(selected Is Nothing, 0, dt.Rows.IndexOf(dt.Select($"{COL_CODE} = '{selected.Replace("'", "''")}'").FirstOrDefault()))
        CboCustomer.SelectedIndex = Math.Max(0, index)
    End Sub

    ''' <summary>型式の候補を入れ替える（入力中の文字は残す）</summary>
    Private Sub SetModelCandidates(models As IEnumerable(Of String))
        Dim text = CboModel.Text
        Dim caret = CboModel.SelectionStart
        CboModel.BeginUpdate()
        Try
            CboModel.Items.Clear()
            CboModel.Items.AddRange(models.Cast(Of Object)().ToArray())
        Finally
            CboModel.EndUpdate()
        End Try
        CboModel.Text = text
        CboModel.SelectionStart = Math.Min(caret, text.Length)
    End Sub

    ''' <summary>選択中の得意先CD（共用棚・未選択は Nothing）</summary>
    Private Function SelectedCustomerCode() As String
        If CboCustomer.SelectedIndex <= 0 Then Return Nothing
        Dim code = TryCast(CboCustomer.SelectedValue, String)
        Return If(String.IsNullOrEmpty(code), Nothing, code)
    End Function

    ''' <summary>型式が取扱得意先に一致したら、得意先をその型式の取扱得意先に絞る</summary>
    Private Sub CboModel_TextChanged(sender As Object, e As EventArgs) Handles CboModel.TextChanged
        If _updating Then Return

        ' 型式を打ち直したら商品の選択はやり直し
        If CboItem.Visible Then ShowItemRow(False)
        CboItem.Items.Clear()
        _searchedModel = Nothing

        Dim model = CboModel.Text.Trim()
        Dim filterModel = If(_customersByModel.ContainsKey(model), model, Nothing)
        If filterModel = _customerFilterModel Then Return
        _customerFilterModel = filterModel

        Dim before = SelectedCustomerCode()
        _updating = True
        Try
            If filterModel Is Nothing Then
                SetCustomerCandidates(_customers.Values)
            Else
                SetCustomerCandidates(_customersByModel(filterModel).Select(Function(code) _customers(code)))
            End If
        Finally
            _updating = False
        End Try

        ' 選んでいた得意先が候補から外れたら、型式の候補も絞り込みを解く
        If SelectedCustomerCode() <> before Then UpdateModelCandidates()
    End Sub

    ''' <summary>得意先を選んだら、型式をその得意先の取扱型式に絞る</summary>
    Private Sub CboCustomer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboCustomer.SelectedIndexChanged
        If _updating Then Return
        UpdateModelCandidates()
    End Sub

    ''' <summary>一覧から選ばずに短縮CDを打ち込んだ場合も、欄を離れたら得意先を確定して型式を絞る</summary>
    Private Sub CboCustomer_Leave(sender As Object, e As EventArgs) Handles CboCustomer.Leave
        If CboCustomer.SelectedIndex >= 0 Then Return
        Dim found As Boolean
        Dim cust = ResolveCustomer(found)
        If found AndAlso cust IsNot Nothing Then
            CboCustomer.SelectedValue = cust.Code
        End If
    End Sub

    Private Sub UpdateModelCandidates()
        Dim code = SelectedCustomerCode()
        If code = _modelFilterCustomer Then Return
        _modelFilterCustomer = code

        _updating = True
        Try
            Dim models As SortedSet(Of String) = Nothing
            If code IsNot Nothing AndAlso _modelsByCustomer.TryGetValue(code, models) Then
                SetModelCandidates(models)
            ElseIf code IsNot Nothing Then
                SetModelCandidates(Array.Empty(Of String)())
            Else
                SetModelCandidates(_allModels)
            End If
        Finally
            _updating = False
        End Try
    End Sub

    ' ============================================================
    ' 追加
    ' ============================================================

    ''' <summary>「商品」欄（同じ型式の商品が複数あるときだけ）の表示を切り替え、ボタン位置と画面の高さを合わせる</summary>
    Private Sub ShowItemRow(visible As Boolean)
        LblItem.Visible = visible
        CboItem.Visible = visible
        Dim gap = CboCustomer.Top - CboModel.Bottom
        Dim top = If(visible, CboItem.Bottom, CboCustomer.Bottom) + gap
        BtnOk.Top = top
        BtnCancel.Top = top
        Me.ClientSize = New Size(Me.ClientSize.Width, BtnOk.Bottom + gap)
    End Sub

    ''' <summary>
    ''' 型式から商品を特定する。複数あれば「商品」欄を出して選ばせ、選択済みならその商品を返す。
    ''' </summary>
    Private Function ResolveItem() As String
        Dim model = CboModel.Text.Trim()
        If model = "" Then
            MsgBox("型式を入力してください。", MsgBoxStyle.Exclamation, Me.Text)
            CboModel.Focus()
            Return Nothing
        End If

        If model <> _searchedModel Then
            Me.Cursor = Cursors.WaitCursor
            Try
                _items = ShelfRepository.FindItemsByModel(model)
                _searchedModel = model
            Finally
                Me.Cursor = Cursors.Default
            End Try

            If _items.Count = 0 Then
                MsgBox($"型式「{model}」は商品マスタに見つかりません（完全一致で検索します）。", MsgBoxStyle.Exclamation, Me.Text)
                CboModel.Focus()
                Return Nothing
            End If

            If _items.Count > 1 Then
                CboItem.Items.Clear()
                For Each kv In _items
                    CboItem.Items.Add($"{kv.Key}：{kv.Value}")
                Next
                CboItem.SelectedIndex = -1
                ShowItemRow(True)
                MsgBox($"型式「{model}」の商品が {_items.Count} 件あります。「商品」欄で選んでから、もう一度「追加」を押してください。", MsgBoxStyle.Information, Me.Text)
                CboItem.Focus()
                Return Nothing
            End If
        End If

        If _items.Count = 1 Then Return _items.Keys.First()
        If CboItem.SelectedIndex < 0 Then
            MsgBox("「商品」欄で商品を選んでください。", MsgBoxStyle.Exclamation, Me.Text)
            CboItem.Focus()
            Return Nothing
        End If
        Return _items.Keys.ElementAt(CboItem.SelectedIndex)
    End Function

    ''' <summary>
    ''' 入力・選択された得意先を特定する。空欄・共用棚は Nothing（共用棚）。
    ''' 一覧から選ばずに打ち込んだ場合は「短縮CD」または「短縮CD 略称」と一致する得意先を候補の中から探す。
    ''' </summary>
    Private Function ResolveCustomer(ByRef found As Boolean) As ShelfRepository.Customer
        found = True
        Dim code = SelectedCustomerCode()
        If code IsNot Nothing Then Return _customers(code)
        If CboCustomer.SelectedIndex = 0 Then Return Nothing

        Dim text = CboCustomer.Text.Trim()
        If text = "" OrElse text = SHARED_LABEL Then Return Nothing

        Dim candidates = DirectCast(CboCustomer.DataSource, DataTable).Rows.Cast(Of DataRow)() _
            .Select(Function(r) CStr(r(COL_CODE))).Where(Function(c) c <> "").Select(Function(c) _customers(c))
        Dim matches = candidates.Where(Function(c) c.ShortCode = text OrElse $"{c.ShortCode} {c.Name}".Trim() = text).ToList()
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

        Dim key = ResolveItem()
        If key Is Nothing Then Return

        _ItemCode = key
        _Model = _items(key)
        _SelectedCustomer = cust
        DialogResult = DialogResult.OK
    End Sub

    ' 得意先欄で Enter を押したら追加（受発注管理の得意先選択と同じ操作感）
    Private Sub CboCustomer_KeyDown(sender As Object, e As KeyEventArgs) Handles CboCustomer.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            BtnOk.PerformClick()
        End If
    End Sub

End Class
