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
        Me.TxtModel = New System.Windows.Forms.TextBox()
        Me.BtnFind = New System.Windows.Forms.Button()
        Me.LblItem = New System.Windows.Forms.Label()
        Me.CboItem = New System.Windows.Forms.ComboBox()
        Me.LblCustomer = New System.Windows.Forms.Label()
        Me.CboCustomer = New System.Windows.Forms.ComboBox()
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
        'TxtModel
        '
        Me.TxtModel.Location = New System.Drawing.Point(84, 16)
        Me.TxtModel.Name = "TxtModel"
        Me.TxtModel.Size = New System.Drawing.Size(280, 27)
        Me.TxtModel.TabIndex = 1
        '
        'BtnFind
        '
        Me.BtnFind.Location = New System.Drawing.Point(372, 12)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(80, 34)
        Me.BtnFind.TabIndex = 2
        Me.BtnFind.Text = "検索"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'LblItem
        '
        Me.LblItem.AutoSize = True
        Me.LblItem.Location = New System.Drawing.Point(16, 60)
        Me.LblItem.Name = "LblItem"
        Me.LblItem.Size = New System.Drawing.Size(51, 20)
        Me.LblItem.TabIndex = 3
        Me.LblItem.Text = "商品"
        '
        'CboItem
        '
        Me.CboItem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboItem.FormattingEnabled = True
        Me.CboItem.Location = New System.Drawing.Point(84, 56)
        Me.CboItem.Name = "CboItem"
        Me.CboItem.Size = New System.Drawing.Size(368, 28)
        Me.CboItem.TabIndex = 4
        '
        'LblCustomer
        '
        Me.LblCustomer.AutoSize = True
        Me.LblCustomer.Location = New System.Drawing.Point(16, 100)
        Me.LblCustomer.Name = "LblCustomer"
        Me.LblCustomer.Size = New System.Drawing.Size(51, 20)
        Me.LblCustomer.TabIndex = 5
        Me.LblCustomer.Text = "得意先"
        '
        'CboCustomer
        '
        Me.CboCustomer.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CboCustomer.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CboCustomer.FormattingEnabled = True
        Me.CboCustomer.IntegralHeight = False
        Me.CboCustomer.MaxDropDownItems = 15
        Me.CboCustomer.Location = New System.Drawing.Point(84, 96)
        Me.CboCustomer.Name = "CboCustomer"
        Me.CboCustomer.Size = New System.Drawing.Size(368, 28)
        Me.CboCustomer.TabIndex = 6
        '
        'BtnOk
        '
        Me.BtnOk.Location = New System.Drawing.Point(264, 140)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(92, 36)
        Me.BtnOk.TabIndex = 7
        Me.BtnOk.Text = "追加"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(360, 140)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(92, 36)
        Me.BtnCancel.TabIndex = 8
        Me.BtnCancel.Text = "キャンセル"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'FrmShelfAdd
        '
        Me.AcceptButton = Me.BtnFind
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(468, 190)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOk)
        Me.Controls.Add(Me.CboCustomer)
        Me.Controls.Add(Me.LblCustomer)
        Me.Controls.Add(Me.CboItem)
        Me.Controls.Add(Me.LblItem)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.TxtModel)
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
    Friend WithEvents TxtModel As TextBox
    Friend WithEvents BtnFind As Button
    Friend WithEvents LblItem As Label
    Friend WithEvents CboItem As ComboBox
    Friend WithEvents LblCustomer As Label
    Friend WithEvents CboCustomer As ComboBox
    Friend WithEvents BtnOk As Button
    Friend WithEvents BtnCancel As Button
End Class
