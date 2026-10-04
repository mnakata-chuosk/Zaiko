<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
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
        Me.Mbar = New ChuoUtils.MenuBar()
        Me.USelector = New ChuoUtils.UserSelector()
        Me.LblOffice = New System.Windows.Forms.Label()
        Me.CboOffice = New System.Windows.Forms.ComboBox()
        Me.GbxStocktaking = New System.Windows.Forms.GroupBox()
        Me.ChkPrintStock = New System.Windows.Forms.CheckBox()
        Me.BtnStocktaking = New System.Windows.Forms.Button()
        Me.LblShelfSlots = New System.Windows.Forms.Label()
        Me.CboShelfSlots = New System.Windows.Forms.ComboBox()
        Me.GbxShelf = New System.Windows.Forms.GroupBox()
        Me.LblShelfGuide = New System.Windows.Forms.Label()
        Me.BtnShelf = New System.Windows.Forms.Button()
        Me.GbxStocktaking.SuspendLayout()
        Me.GbxShelf.SuspendLayout()
        Me.SuspendLayout()
        '
        'Mbar
        '
        Me.Mbar.Dock = System.Windows.Forms.DockStyle.Top
        Me.Mbar.Location = New System.Drawing.Point(0, 0)
        Me.Mbar.Margin = New System.Windows.Forms.Padding(4, 2, 4, 2)
        Me.Mbar.Name = "Mbar"
        Me.Mbar.Size = New System.Drawing.Size(360, 26)
        Me.Mbar.TabIndex = 0
        '
        'USelector
        '
        Me.USelector.Location = New System.Drawing.Point(16, 34)
        Me.USelector.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.USelector.MaximumSize = New System.Drawing.Size(500, 29)
        Me.USelector.MinimumSize = New System.Drawing.Size(300, 29)
        Me.USelector.Name = "USelector"
        Me.USelector.Size = New System.Drawing.Size(328, 29)
        Me.USelector.TabIndex = 1
        '
        'LblOffice
        '
        Me.LblOffice.AutoSize = True
        Me.LblOffice.Location = New System.Drawing.Point(16, 79)
        Me.LblOffice.Name = "LblOffice"
        Me.LblOffice.Size = New System.Drawing.Size(51, 20)
        Me.LblOffice.TabIndex = 2
        Me.LblOffice.Text = "営業所"
        '
        'CboOffice
        '
        Me.CboOffice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboOffice.FormattingEnabled = True
        Me.CboOffice.Location = New System.Drawing.Point(76, 75)
        Me.CboOffice.Name = "CboOffice"
        Me.CboOffice.Size = New System.Drawing.Size(140, 28)
        Me.CboOffice.TabIndex = 3
        '
        'GbxStocktaking
        '
        Me.GbxStocktaking.Controls.Add(Me.CboShelfSlots)
        Me.GbxStocktaking.Controls.Add(Me.LblShelfSlots)
        Me.GbxStocktaking.Controls.Add(Me.ChkPrintStock)
        Me.GbxStocktaking.Controls.Add(Me.BtnStocktaking)
        Me.GbxStocktaking.Location = New System.Drawing.Point(16, 116)
        Me.GbxStocktaking.Name = "GbxStocktaking"
        Me.GbxStocktaking.Size = New System.Drawing.Size(328, 104)
        Me.GbxStocktaking.TabIndex = 4
        Me.GbxStocktaking.TabStop = False
        Me.GbxStocktaking.Text = "棚卸表"
        '
        'ChkPrintStock
        '
        Me.ChkPrintStock.AutoSize = True
        Me.ChkPrintStock.Location = New System.Drawing.Point(16, 66)
        Me.ChkPrintStock.Name = "ChkPrintStock"
        Me.ChkPrintStock.Size = New System.Drawing.Size(151, 24)
        Me.ChkPrintStock.TabIndex = 2
        Me.ChkPrintStock.Text = "印刷用に在庫数を表示"
        Me.ChkPrintStock.UseVisualStyleBackColor = True
        '
        'BtnStocktaking
        '
        Me.BtnStocktaking.Location = New System.Drawing.Point(216, 58)
        Me.BtnStocktaking.Name = "BtnStocktaking"
        Me.BtnStocktaking.Size = New System.Drawing.Size(96, 36)
        Me.BtnStocktaking.TabIndex = 3
        Me.BtnStocktaking.Text = "出力"
        Me.BtnStocktaking.UseVisualStyleBackColor = True
        '
        'LblShelfSlots
        '
        Me.LblShelfSlots.AutoSize = True
        Me.LblShelfSlots.Location = New System.Drawing.Point(16, 32)
        Me.LblShelfSlots.Name = "LblShelfSlots"
        Me.LblShelfSlots.Size = New System.Drawing.Size(79, 20)
        Me.LblShelfSlots.TabIndex = 0
        Me.LblShelfSlots.Text = "棚番表示数"
        '
        'CboShelfSlots
        '
        Me.CboShelfSlots.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboShelfSlots.FormattingEnabled = True
        Me.CboShelfSlots.Location = New System.Drawing.Point(104, 28)
        Me.CboShelfSlots.Name = "CboShelfSlots"
        Me.CboShelfSlots.Size = New System.Drawing.Size(56, 28)
        Me.CboShelfSlots.TabIndex = 1
        '
        'GbxShelf
        '
        Me.GbxShelf.Controls.Add(Me.LblShelfGuide)
        Me.GbxShelf.Controls.Add(Me.BtnShelf)
        Me.GbxShelf.Location = New System.Drawing.Point(16, 228)
        Me.GbxShelf.Name = "GbxShelf"
        Me.GbxShelf.Size = New System.Drawing.Size(328, 68)
        Me.GbxShelf.TabIndex = 5
        Me.GbxShelf.TabStop = False
        Me.GbxShelf.Text = "棚番"
        '
        'LblShelfGuide
        '
        Me.LblShelfGuide.AutoSize = True
        Me.LblShelfGuide.Location = New System.Drawing.Point(16, 32)
        Me.LblShelfGuide.Name = "LblShelfGuide"
        Me.LblShelfGuide.Size = New System.Drawing.Size(163, 20)
        Me.LblShelfGuide.TabIndex = 0
        Me.LblShelfGuide.Text = "編集・Excelから反映"
        '
        'BtnShelf
        '
        Me.BtnShelf.Location = New System.Drawing.Point(216, 22)
        Me.BtnShelf.Name = "BtnShelf"
        Me.BtnShelf.Size = New System.Drawing.Size(96, 36)
        Me.BtnShelf.TabIndex = 1
        Me.BtnShelf.Text = "棚番編集"
        Me.BtnShelf.UseVisualStyleBackColor = True
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(360, 310)
        Me.Controls.Add(Me.GbxShelf)
        Me.Controls.Add(Me.GbxStocktaking)
        Me.Controls.Add(Me.CboOffice)
        Me.Controls.Add(Me.LblOffice)
        Me.Controls.Add(Me.USelector)
        Me.Controls.Add(Me.Mbar)
        Me.Font = New System.Drawing.Font("Yu Gothic UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "FrmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "在庫表"
        Me.GbxStocktaking.ResumeLayout(False)
        Me.GbxStocktaking.PerformLayout()
        Me.GbxShelf.ResumeLayout(False)
        Me.GbxShelf.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Mbar As ChuoUtils.MenuBar
    Friend WithEvents USelector As ChuoUtils.UserSelector
    Friend WithEvents LblOffice As Label
    Friend WithEvents CboOffice As ComboBox
    Friend WithEvents GbxStocktaking As GroupBox
    Friend WithEvents ChkPrintStock As CheckBox
    Friend WithEvents BtnStocktaking As Button
    Friend WithEvents LblShelfSlots As Label
    Friend WithEvents CboShelfSlots As ComboBox
    Friend WithEvents GbxShelf As GroupBox
    Friend WithEvents LblShelfGuide As Label
    Friend WithEvents BtnShelf As Button
End Class
