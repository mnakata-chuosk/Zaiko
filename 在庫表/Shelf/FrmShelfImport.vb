''' <summary>
''' Excel 読み込みの変更確認。変更がある行だけを 変更前 → 変更後 で表示し、反映する行を選ばせる。
''' </summary>
Public Class FrmShelfImport

    Private Const COL_APPLY As String = "適用"
    Private Shared ReadOnly COLOR_WARNING As Color = Color.FromArgb(255, 235, 156)

    ''' <summary>チェックされた変更</summary>
    Public ReadOnly Property SelectedChanges As New List(Of ShelfImportChange)

    Public Sub New(changes As List(Of ShelfImportChange), errors As List(Of String))
        InitializeComponent()

        SetupColumns()
        For Each c In changes
            Dim i = Dgv.Rows.Add(
                c.Warnings.Count = 0,                                       ' 確認事項のある行は既定で外す
                String.Join(",", c.ExcelRows),
                c.Group.Model,
                If(c.Group.IsShared, "（共用棚）", $"{c.Group.CustomerShortCode} {c.Group.CustomerName}".Trim()),
                Join(c.Before),
                Join(c.After),
                String.Join(" / ", c.Warnings) & If(c.IsNewGroup, If(c.Warnings.Count > 0, " / ", "") & "新しい行", ""))
            Dgv.Rows(i).Tag = c
            If c.Warnings.Count > 0 Then Dgv.Rows(i).DefaultCellStyle.BackColor = COLOR_WARNING
        Next

        Me.Text = $"棚番の変更確認（Excel読込）　変更 {changes.Count} 件"
        If errors.Count > 0 Then
            TxtErrors.Text = $"反映できなかった行（{errors.Count} 件）" & vbCrLf & String.Join(vbCrLf, errors)
        Else
            TxtErrors.Visible = False
        End If
        BtnApply.Enabled = changes.Count > 0
    End Sub

    Private Sub SetupColumns()
        Dgv.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = COL_APPLY, .HeaderText = COL_APPLY, .Width = 48})
        Dim add = Sub(name As String, width As Integer)
                      Dgv.Columns.Add(New DataGridViewTextBoxColumn With {
                          .Name = name, .HeaderText = name, .Width = width, .ReadOnly = True,
                          .SortMode = DataGridViewColumnSortMode.NotSortable
                      })
                  End Sub
        add("Excel行", 70)
        add("型式", 220)
        add("得意先", 150)
        add("変更前", 230)
        add("変更後", 230)
        add("確認事項", 260)
        Dgv.Columns("変更前").DefaultCellStyle.ForeColor = Color.DimGray
        Dgv.Columns("確認事項").DefaultCellStyle.ForeColor = Color.DarkRed
    End Sub

    Private Shared Function Join(shelves As List(Of String)) As String
        Return If(shelves.Count = 0, "（なし）", String.Join(" / ", shelves))
    End Function

    Private Sub SetAll(value As Boolean)
        Dgv.EndEdit()
        For Each r As DataGridViewRow In Dgv.Rows
            r.Cells(COL_APPLY).Value = value
        Next
    End Sub

    Private Sub BtnCheckAll_Click(sender As Object, e As EventArgs) Handles BtnCheckAll.Click
        SetAll(True)
    End Sub

    Private Sub BtnUncheckAll_Click(sender As Object, e As EventArgs) Handles BtnUncheckAll.Click
        SetAll(False)
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As EventArgs) Handles BtnApply.Click
        Dgv.EndEdit()
        SelectedChanges.Clear()
        For Each r As DataGridViewRow In Dgv.Rows
            If CBool(r.Cells(COL_APPLY).Value) Then SelectedChanges.Add(DirectCast(r.Tag, ShelfImportChange))
        Next
        If SelectedChanges.Count = 0 Then
            MsgBox("反映する行にチェックを付けてください。", MsgBoxStyle.Exclamation, Me.Text)
            Return
        End If
        DialogResult = DialogResult.OK
    End Sub

    ' チェックボックスはクリックですぐ確定させる
    Private Sub Dgv_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles Dgv.CurrentCellDirtyStateChanged
        If Dgv.IsCurrentCellDirty AndAlso TypeOf Dgv.CurrentCell Is DataGridViewCheckBoxCell Then
            Dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

End Class
