Imports ChuoUtils

Public Class FrmMain

    Private Const APP_NAME As String = "在庫表"

    Private _userSettings As UserSettings

    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.AddVersion
        Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)

        If Not USelector.EnsureSignedIn(Me) Then
            Me.Close()
            Return
        End If

        Using ps As New PcSettings()
            _userSettings = New UserSettings(ps.UserID)
        End Using

        InitOfficeList()
        InitShelfSlots()
    End Sub

    ''' <summary>棚番表示数（[棚番・数量] の組数）を初期化する</summary>
    Private Sub InitShelfSlots()
        CboShelfSlots.Items.Clear()
        For n As Integer = StocktakingOptions.MIN_SHELF_SLOTS To StocktakingOptions.MAX_SHELF_SLOTS
            CboShelfSlots.Items.Add(n)
        Next
        CboShelfSlots.SelectedItem = StocktakingOptions.DEFAULT_SHELF_SLOTS
    End Sub

    ''' <summary>営業所コンボを初期化し、担当者の初期値営業所を選択する</summary>
    Private Sub InitOfficeList()
        CboOffice.Items.Clear()
        CboOffice.Items.AddRange(Chuo.OfficeList)

        Dim idx As Integer = Array.IndexOf(Chuo.OfficeCDList, _userSettings?.CurrentOffice)
        CboOffice.SelectedIndex = If(idx >= 0, idx, 0)
    End Sub

    ''' <summary>選択中の営業所コード（"01" 等）</summary>
    Private ReadOnly Property SelectedOfficeCode As String
        Get
            Return Chuo.OfficeCDList(CboOffice.SelectedIndex)
        End Get
    End Property

    Private Sub BtnShelf_Click(sender As Object, e As EventArgs) Handles BtnShelf.Click
        Using frm As New FrmShelf(CboOffice.SelectedIndex, _userSettings?.UserID)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub BtnStocktaking_Click(sender As Object, e As EventArgs) Handles BtnStocktaking.Click
        Dim officeCode As String = SelectedOfficeCode
        Dim officeIndex As Integer = CboOffice.SelectedIndex

        Me.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim rows As List(Of StockRow) = StocktakingLoader.LoadStock(officeCode)
            If rows.Count = 0 Then
                MsgBox($"{CboOffice.Text} に出力対象の在庫がありません。", MsgBoxStyle.Information, APP_NAME)
                Exit Sub
            End If

            Dim items As List(Of StocktakingItem) = StocktakingBuilder.Build(rows, StocktakingLoader.LoadShelves(officeCode))

            Dim opt As New StocktakingOptions With {
                .OfficeCode = officeCode,
                .OfficeLabel = Chuo.OfficeList(officeIndex) & "営業所",
                .OfficeHeader = $"{officeCode}：{Chuo.OfficeNMList(officeIndex)}営業所",
                .ShelfSlots = CInt(CboShelfSlots.SelectedItem),
                .ShowStockOnPrintSheet = ChkPrintStock.Checked,
                .OutputAt = DateTime.Now
            }
            StocktakingWriter.Write(items, opt)

            App.AddCount(APP_NAME, "棚卸表", $"{officeCode} 棚番：{opt.ShelfSlots} 印刷用在庫数：{If(ChkPrintStock.Checked, "○", "☓")}")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, APP_NAME)
            App.WriteErrLog(APP_NAME, ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
            Me.Enabled = True
        End Try
    End Sub

End Class
