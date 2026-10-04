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

    Private Sub BtnStocktaking_Click(sender As Object, e As EventArgs) Handles BtnStocktaking.Click
        ' TODO: 棚卸表出力（データ取得 → 一覧シート・印刷用シート作成）を実装する
        MsgBox($"{CboOffice.Text} の棚卸表出力は未実装です。", MsgBoxStyle.Information, APP_NAME)
    End Sub

End Class
