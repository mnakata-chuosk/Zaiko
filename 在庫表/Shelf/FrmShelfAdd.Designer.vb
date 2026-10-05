<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmShelfAdd
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
        Me.LblModel = New System.Windows.Forms.Label()
        Me.CboModel = New System.Windows.Forms.ComboBox()
        Me.LblCustomer = New System.Windows.Forms.Label()
        Me.CboCustomer = New System.Windows.Forms.ComboBox()
        Me.LblItem = New System.Windows.Forms.Label()
        Me.CboItem = New System.Windows.Forms.ComboBox()
        Me.BtnOk = New System.Windows.Forms.Button()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LblModel
        '
        Me.LblModel.AutoSize = True
        Me.LblModel.Location = New System.Drawing.Point(16, 20)
        Me.LblModel.Name = "LblModel"
        Me.LblModel.Size = New System.Drawing.Size(37, 20)
        Me.LblModel.TabIndex = 0
        Me.LblModel.Text = "型式"
        '
        'CboModel
        '
        Me.CboModel.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CboModel.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CboModel.FormattingEnabled = True
        Me.CboModel.IntegralHeight = False
        Me.CboModel.MaxDropDownItems = 15
        Me.CboModel.Location = New System.Drawing.Point(84, 16)
        Me.CboModel.Name = "CboModel"
        Me.CboModel.Size = New System.Drawing.Size(368, 28)
        Me.CboModel.TabIndex = 1
        '
        'LblCustomer
        '
        Me.LblCustomer.AutoSize = True
        Me.LblCustomer.Location = New System.Drawing.Point(16, 60)
        Me.LblCustomer.Name = "LblCustomer"
        Me.LblCustomer.Size = New System.Drawing.Size(51, 20)
        Me.LblCustomer.TabIndex = 2
        Me.LblCustomer.Text = "得意先"
        '
        'CboCustomer
        '
        Me.CboCustomer.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CboCustomer.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CboCustomer.FormattingEnabled = True
        Me.CboCustomer.IntegralHeight = False
        Me.CboCustomer.MaxDropDownItems = 15
        Me.CboCustomer.Location = New System.Drawing.Point(84, 56)
        Me.CboCustomer.Name = "CboCustomer"
        Me.CboCustomer.Size = New System.Drawing.Size(368, 28)
        Me.CboCustomer.TabIndex = 3
        '
        'LblItem
        '
        Me.LblItem.AutoSize = True
        Me.LblItem.Location = New System.Drawing.Point(16, 100)
        Me.LblItem.Name = "LblItem"
        Me.LblItem.Size = New System.Drawing.Size(37, 20)
        Me.LblItem.TabIndex = 4
        Me.LblItem.Text = "商品"
        '
        'CboItem
        '
        Me.CboItem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboItem.FormattingEnabled = True
        Me.CboItem.Location = New System.Drawing.Point(84, 96)
        Me.CboItem.Name = "CboItem"
        Me.CboItem.Size = New System.Drawing.Size(368, 28)
        Me.CboItem.TabIndex = 5
        '
        'BtnOk
        '
        Me.BtnOk.Location = New System.Drawing.Point(264, 140)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(92, 36)
        Me.BtnOk.TabIndex = 6
        Me.BtnOk.Text = "追加"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(360, 140)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(92, 36)
        Me.BtnCancel.TabIndex = 7
        Me.BtnCancel.Text = "キャンセル"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'FrmShelfAdd
        '
        Me.AcceptButton = Me.BtnOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(468, 190)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOk)
        Me.Controls.Add(Me.CboItem)
        Me.Controls.Add(Me.LblItem)
        Me.Controls.Add(Me.CboCustomer)
        Me.Controls.Add(Me.LblCustomer)
        Me.Controls.Add(Me.CboModel)
        Me.Controls.Add(Me.LblModel)
        Me.Font = New System.Drawing.Font("Yu Gothic UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmShelfAdd"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "棚番の行追加"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblModel As Label
    Friend WithEvents CboModel As ComboBox
    Friend WithEvents LblCustomer As Label
    Friend WithEvents CboCustomer As ComboBox
    Friend WithEvents LblItem As Label
    Friend WithEvents CboItem As ComboBox
    Friend WithEvents BtnOk As Button
    Friend WithEvents BtnCancel As Button
End Class
