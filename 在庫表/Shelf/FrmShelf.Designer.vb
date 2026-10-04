<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmShelf
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
        Me.PnlTop = New System.Windows.Forms.Panel()
        Me.BtnSave = New System.Windows.Forms.Button()
        Me.BtnImport = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnReload = New System.Windows.Forms.Button()
        Me.ChkChangedOnly = New System.Windows.Forms.CheckBox()
        Me.ChkStockOnly = New System.Windows.Forms.CheckBox()
        Me.TxtSearch = New System.Windows.Forms.TextBox()
        Me.LblSearch = New System.Windows.Forms.Label()
        Me.CboOffice = New System.Windows.Forms.ComboBox()
        Me.LblOffice = New System.Windows.Forms.Label()
        Me.Dgv = New System.Windows.Forms.DataGridView()
        Me.StsBar = New System.Windows.Forms.StatusStrip()
        Me.LblStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.PnlTop.SuspendLayout()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StsBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'PnlTop
        '
        Me.PnlTop.Controls.Add(Me.BtnSave)
        Me.PnlTop.Controls.Add(Me.BtnImport)
        Me.PnlTop.Controls.Add(Me.BtnAdd)
        Me.PnlTop.Controls.Add(Me.BtnReload)
        Me.PnlTop.Controls.Add(Me.ChkChangedOnly)
        Me.PnlTop.Controls.Add(Me.ChkStockOnly)
        Me.PnlTop.Controls.Add(Me.TxtSearch)
        Me.PnlTop.Controls.Add(Me.LblSearch)
        Me.PnlTop.Controls.Add(Me.CboOffice)
        Me.PnlTop.Controls.Add(Me.LblOffice)
        Me.PnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.PnlTop.Location = New System.Drawing.Point(0, 0)
        Me.PnlTop.Name = "PnlTop"
        Me.PnlTop.Size = New System.Drawing.Size(1084, 84)
        Me.PnlTop.TabIndex = 0
        '
        'LblOffice
        '
        Me.LblOffice.AutoSize = True
        Me.LblOffice.Location = New System.Drawing.Point(12, 14)
        Me.LblOffice.Name = "LblOffice"
        Me.LblOffice.Size = New System.Drawing.Size(51, 20)
        Me.LblOffice.TabIndex = 0
        Me.LblOffice.Text = "営業所"
        '
        'CboOffice
        '
        Me.CboOffice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboOffice.FormattingEnabled = True
        Me.CboOffice.Location = New System.Drawing.Point(72, 10)
        Me.CboOffice.Name = "CboOffice"
        Me.CboOffice.Size = New System.Drawing.Size(140, 28)
        Me.CboOffice.TabIndex = 1
        '
        'LblSearch
        '
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(12, 52)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(37, 20)
        Me.LblSearch.TabIndex = 2
        Me.LblSearch.Text = "型式"
        '
        'TxtSearch
        '
        Me.TxtSearch.Location = New System.Drawing.Point(72, 48)
        Me.TxtSearch.Name = "TxtSearch"
        Me.TxtSearch.Size = New System.Drawing.Size(240, 27)
        Me.TxtSearch.TabIndex = 3
        '
        'ChkStockOnly
        '
        Me.ChkStockOnly.AutoSize = True
        Me.ChkStockOnly.Location = New System.Drawing.Point(328, 50)
        Me.ChkStockOnly.Name = "ChkStockOnly"
        Me.ChkStockOnly.Size = New System.Drawing.Size(103, 24)
        Me.ChkStockOnly.TabIndex = 4
        Me.ChkStockOnly.Text = "在庫ありのみ"
        Me.ChkStockOnly.UseVisualStyleBackColor = True
        '
        'ChkChangedOnly
        '
        Me.ChkChangedOnly.AutoSize = True
        Me.ChkChangedOnly.Location = New System.Drawing.Point(448, 50)
        Me.ChkChangedOnly.Name = "ChkChangedOnly"
        Me.ChkChangedOnly.Size = New System.Drawing.Size(103, 24)
        Me.ChkChangedOnly.TabIndex = 5
        Me.ChkChangedOnly.Text = "変更行のみ"
        Me.ChkChangedOnly.UseVisualStyleBackColor = True
        '
        'BtnReload
        '
        Me.BtnReload.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnReload.Location = New System.Drawing.Point(632, 42)
        Me.BtnReload.Name = "BtnReload"
        Me.BtnReload.Size = New System.Drawing.Size(100, 36)
        Me.BtnReload.TabIndex = 6
        Me.BtnReload.Text = "再読込"
        Me.BtnReload.UseVisualStyleBackColor = True
        '
        'BtnAdd
        '
        Me.BtnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnAdd.Location = New System.Drawing.Point(740, 42)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(100, 36)
        Me.BtnAdd.TabIndex = 7
        Me.BtnAdd.Text = "行追加"
        Me.BtnAdd.UseVisualStyleBackColor = True
        '
        'BtnImport
        '
        Me.BtnImport.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnImport.Location = New System.Drawing.Point(848, 42)
        Me.BtnImport.Name = "BtnImport"
        Me.BtnImport.Size = New System.Drawing.Size(116, 36)
        Me.BtnImport.TabIndex = 8
        Me.BtnImport.Text = "Excel読込"
        Me.BtnImport.UseVisualStyleBackColor = True
        '
        'BtnSave
        '
        Me.BtnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnSave.Location = New System.Drawing.Point(972, 42)
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(100, 36)
        Me.BtnSave.TabIndex = 9
        Me.BtnSave.Text = "保存"
        Me.BtnSave.UseVisualStyleBackColor = True
        '
        'Dgv
        '
        Me.Dgv.AllowUserToAddRows = False
        Me.Dgv.AllowUserToDeleteRows = False
        Me.Dgv.AllowUserToResizeRows = False
        Me.Dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Dgv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Dgv.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystrokeOrF2
        Me.Dgv.Location = New System.Drawing.Point(0, 84)
        Me.Dgv.Name = "Dgv"
        Me.Dgv.RowHeadersWidth = 24
        Me.Dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.Dgv.Size = New System.Drawing.Size(1084, 545)
        Me.Dgv.TabIndex = 1
        '
        'StsBar
        '
        Me.StsBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LblStatus})
        Me.StsBar.Location = New System.Drawing.Point(0, 629)
        Me.StsBar.Name = "StsBar"
        Me.StsBar.Size = New System.Drawing.Size(1084, 22)
        Me.StsBar.TabIndex = 2
        '
        'LblStatus
        '
        Me.LblStatus.Name = "LblStatus"
        Me.LblStatus.Size = New System.Drawing.Size(0, 17)
        '
        'FrmShelf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(1084, 651)
        Me.Controls.Add(Me.Dgv)
        Me.Controls.Add(Me.StsBar)
        Me.Controls.Add(Me.PnlTop)
        Me.Font = New System.Drawing.Font("Yu Gothic UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(900, 400)
        Me.Name = "FrmShelf"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "棚番編集"
        Me.PnlTop.ResumeLayout(False)
        Me.PnlTop.PerformLayout()
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StsBar.ResumeLayout(False)
        Me.StsBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PnlTop As Panel
    Friend WithEvents LblOffice As Label
    Friend WithEvents CboOffice As ComboBox
    Friend WithEvents LblSearch As Label
    Friend WithEvents TxtSearch As TextBox
    Friend WithEvents ChkStockOnly As CheckBox
    Friend WithEvents ChkChangedOnly As CheckBox
    Friend WithEvents BtnReload As Button
    Friend WithEvents BtnAdd As Button
    Friend WithEvents BtnImport As Button
    Friend WithEvents BtnSave As Button
    Friend WithEvents Dgv As DataGridView
    Friend WithEvents StsBar As StatusStrip
    Friend WithEvents LblStatus As ToolStripStatusLabel
End Class
