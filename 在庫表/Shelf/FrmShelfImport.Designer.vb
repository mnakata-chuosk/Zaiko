<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmShelfImport
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.LblGuide = New System.Windows.Forms.Label()
        Me.Dgv = New System.Windows.Forms.DataGridView()
        Me.TxtErrors = New System.Windows.Forms.TextBox()
        Me.PnlBottom = New System.Windows.Forms.Panel()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.BtnApply = New System.Windows.Forms.Button()
        Me.BtnUncheckAll = New System.Windows.Forms.Button()
        Me.BtnCheckAll = New System.Windows.Forms.Button()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'LblGuide
        '
        Me.LblGuide.Dock = System.Windows.Forms.DockStyle.Top
        Me.LblGuide.Location = New System.Drawing.Point(0, 0)
        Me.LblGuide.Name = "LblGuide"
        Me.LblGuide.Padding = New System.Windows.Forms.Padding(8, 8, 8, 0)
        Me.LblGuide.Size = New System.Drawing.Size(1084, 36)
        Me.LblGuide.TabIndex = 0
        Me.LblGuide.Text = "変更がある行だけ表示しています。反映する行にチェックを付けて「編集画面に反映」を押してください（保存は編集画面で行います）。"
        '
        'Dgv
        '
        Me.Dgv.AllowUserToAddRows = False
        Me.Dgv.AllowUserToDeleteRows = False
        Me.Dgv.AllowUserToResizeRows = False
        Me.Dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Dgv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Dgv.Location = New System.Drawing.Point(0, 36)
        Me.Dgv.Name = "Dgv"
        Me.Dgv.RowHeadersVisible = False
        Me.Dgv.Size = New System.Drawing.Size(1084, 433)
        Me.Dgv.TabIndex = 1
        '
        'TxtErrors
        '
        Me.TxtErrors.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TxtErrors.ForeColor = System.Drawing.Color.DarkRed
        Me.TxtErrors.Location = New System.Drawing.Point(0, 469)
        Me.TxtErrors.Multiline = True
        Me.TxtErrors.Name = "TxtErrors"
        Me.TxtErrors.ReadOnly = True
        Me.TxtErrors.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtErrors.Size = New System.Drawing.Size(1084, 110)
        Me.TxtErrors.TabIndex = 2
        '
        'PnlBottom
        '
        Me.PnlBottom.Controls.Add(Me.BtnCancel)
        Me.PnlBottom.Controls.Add(Me.BtnApply)
        Me.PnlBottom.Controls.Add(Me.BtnUncheckAll)
        Me.PnlBottom.Controls.Add(Me.BtnCheckAll)
        Me.PnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PnlBottom.Location = New System.Drawing.Point(0, 579)
        Me.PnlBottom.Name = "PnlBottom"
        Me.PnlBottom.Size = New System.Drawing.Size(1084, 52)
        Me.PnlBottom.TabIndex = 3
        '
        'BtnCheckAll
        '
        Me.BtnCheckAll.Location = New System.Drawing.Point(12, 8)
        Me.BtnCheckAll.Name = "BtnCheckAll"
        Me.BtnCheckAll.Size = New System.Drawing.Size(100, 36)
        Me.BtnCheckAll.TabIndex = 0
        Me.BtnCheckAll.Text = "全選択"
        Me.BtnCheckAll.UseVisualStyleBackColor = True
        '
        'BtnUncheckAll
        '
        Me.BtnUncheckAll.Location = New System.Drawing.Point(120, 8)
        Me.BtnUncheckAll.Name = "BtnUncheckAll"
        Me.BtnUncheckAll.Size = New System.Drawing.Size(100, 36)
        Me.BtnUncheckAll.TabIndex = 1
        Me.BtnUncheckAll.Text = "全解除"
        Me.BtnUncheckAll.UseVisualStyleBackColor = True
        '
        'BtnApply
        '
        Me.BtnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnApply.Location = New System.Drawing.Point(828, 8)
        Me.BtnApply.Name = "BtnApply"
        Me.BtnApply.Size = New System.Drawing.Size(140, 36)
        Me.BtnApply.TabIndex = 2
        Me.BtnApply.Text = "編集画面に反映"
        Me.BtnApply.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(976, 8)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(100, 36)
        Me.BtnCancel.TabIndex = 3
        Me.BtnCancel.Text = "キャンセル"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'FrmShelfImport
        '
        Me.AcceptButton = Me.BtnApply
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(1084, 631)
        Me.Controls.Add(Me.Dgv)
        Me.Controls.Add(Me.LblGuide)
        Me.Controls.Add(Me.TxtErrors)
        Me.Controls.Add(Me.PnlBottom)
        Me.Font = New System.Drawing.Font("Yu Gothic UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(800, 400)
        Me.Name = "FrmShelfImport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "棚番の変更確認（Excel読込）"
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlBottom.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblGuide As Label
    Friend WithEvents Dgv As DataGridView
    Friend WithEvents TxtErrors As TextBox
    Friend WithEvents PnlBottom As Panel
    Friend WithEvents BtnCheckAll As Button
    Friend WithEvents BtnUncheckAll As Button
    Friend WithEvents BtnApply As Button
    Friend WithEvents BtnCancel As Button
End Class
