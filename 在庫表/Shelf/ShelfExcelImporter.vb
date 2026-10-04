Option Strict Off  ' Excel COM の遅延バインディングに必要

Imports System.IO
Imports ChuoUtils

''' <summary>
''' 棚卸表（本アプリで出力した一覧シート）の1行分の棚番
''' </summary>
Public Class ShelfImportRow
    Public Enum RowKind
        ''' <summary>単独行（共用棚＋得意先指定棚が混在）</summary>
        SingleRow
        ''' <summary>親行（共用棚）</summary>
        Parent
        ''' <summary>子行（得意先指定棚）</summary>
        Child
        ''' <summary>追記行（型式・得意先は手入力）</summary>
        Append
    End Enum

    Public Property Kind As RowKind
    Public Property ExcelRow As Integer
    Public Property ItemCode As String = ""
    Public Property Model As String = ""
    ''' <summary>得意先CD（非表示列。古いファイルや追記行は空）</summary>
    Public Property CustomerCode As String = ""
    Public Property CustomerShortCode As String = ""
    ''' <summary>棚番（空欄を除いて左から順）</summary>
    Public ReadOnly Property Shelves As New List(Of String)
End Class

''' <summary>
''' 棚卸表ファイルの読み込み結果
''' </summary>
Public Class ShelfImportFile
    Public Property OfficeCode As String
    Public Property SheetName As String
    ''' <summary>ファイル上の棚番の列数（これより後ろの登録済み棚番は変更しない）</summary>
    Public Property Slots As Integer
    Public ReadOnly Property Rows As New List(Of ShelfImportRow)
End Class

''' <summary>
''' 本アプリで出力した棚卸表（一覧シート）から棚番を読み取る
''' </summary>
Public NotInheritable Class ShelfExcelImporter

    Private Sub New()
    End Sub

    Public Shared Function Read(path As String) As ShelfImportFile
        If Not File.Exists(path) Then Throw New FileNotFoundException("ファイルが見つかりません。", path)

        Using xl As New ExcelObject(path, isReadOnly:=True)
            Dim sheetName As String = xl.SheetNames.FirstOrDefault(Function(n) n.EndsWith(ListHeaders.SHEET_SUFFIX))
            If sheetName Is Nothing Then
                Throw New InvalidDataException($"棚卸表の一覧シート（「{ListHeaders.SHEET_SUFFIX.Trim()}」）が見つかりません。本アプリで出力したファイルを指定してください。")
            End If

            Dim values As Object(,) = Nothing
            xl.WithSheet(sheetName,
                Sub(sh)
                    Dim ur = sh.UsedRange
                    Dim lastRow As Integer = ur.Row + ur.Rows.Count - 1
                    Dim lastCol As Integer = ur.Column + ur.Columns.Count - 1
                    values = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, lastCol)).Value2
                End Sub)

            Return Parse(sheetName, values)
        End Using
    End Function

    ''' <summary>シートの値（1始まりの2次元配列）を解析する</summary>
    Private Shared Function Parse(sheetName As String, values As Object(,)) As ShelfImportFile
        Dim result As New ShelfImportFile With {.SheetName = sheetName, .OfficeCode = sheetName.Substring(0, 2)}

        Dim lastRow As Integer = values.GetUpperBound(0)
        Dim lastCol As Integer = values.GetUpperBound(1)
        Const hr As Integer = ListHeaders.HEADER_ROW
        If lastRow < hr Then Throw New InvalidDataException("見出し行がありません。")

        ' 見出しから列位置を求める
        Dim cols As New Dictionary(Of String, Integer)
        Dim shelfCols As New List(Of Integer)
        For c As Integer = 1 To lastCol
            Dim h As String = Text(values(hr, c))
            If h = ListHeaders.SHELF Then
                shelfCols.Add(c)
            ElseIf h <> "" AndAlso Not cols.ContainsKey(h) Then
                cols.Add(h, c)
            End If
        Next
        For Each required In {ListHeaders.NO, ListHeaders.ITEM_CODE, ListHeaders.MODEL, ListHeaders.CUSTOMER, ListHeaders.REGISTERED}
            If Not cols.ContainsKey(required) Then Throw New InvalidDataException($"見出し「{required}」が見つかりません。")
        Next
        If shelfCols.Count = 0 Then Throw New InvalidDataException("見出し「棚番」が見つかりません。")
        result.Slots = shelfCols.Count

        Dim cell = Function(r As Integer, name As String) As String
                       Dim c As Integer
                       Return If(cols.TryGetValue(name, c), Text(values(r, c)), "")
                   End Function

        For r As Integer = hr + 1 To lastRow
            Dim noText As String = cell(r, ListHeaders.NO)
            If cell(r, ListHeaders.NOTE) = ListHeaders.EXCLUDED_NOTE Then Continue For

            Dim row As New ShelfImportRow With {
                .ExcelRow = r,
                .ItemCode = cell(r, ListHeaders.ITEM_CODE),
                .Model = cell(r, ListHeaders.MODEL),
                .CustomerCode = cell(r, ListHeaders.CUSTOMER_CODE),
                .CustomerShortCode = cell(r, ListHeaders.CUSTOMER)
            }
            For Each c In shelfCols
                Dim s As String = Text(values(r, c))
                If s <> "" Then row.Shelves.Add(s)
            Next

            Dim registered As Integer = 0
            Integer.TryParse(cell(r, ListHeaders.REGISTERED), registered)

            If noText = ListHeaders.APPEND_NO Then
                If row.Model = "" AndAlso row.Shelves.Count = 0 Then Continue For
                row.Kind = ShelfImportRow.RowKind.Append
            ElseIf row.ItemCode = "" Then
                Continue For
            ElseIf registered >= 2 Then
                row.Kind = ShelfImportRow.RowKind.Parent
                row.CustomerCode = ""
                row.CustomerShortCode = ""
            ElseIf noText = ListHeaders.CHILD_NO Then
                row.Kind = ShelfImportRow.RowKind.Child
            Else
                row.Kind = ShelfImportRow.RowKind.SingleRow
            End If

            result.Rows.Add(row)
        Next

        Return result
    End Function

    ''' <summary>セル値を文字列に（数値で入力された棚番も "301" のように扱う）</summary>
    Private Shared Function Text(value As Object) As String
        If value Is Nothing OrElse TypeOf value Is DBNull Then Return ""
        If TypeOf value Is Double Then Return CDbl(value).ToString("0.##########")
        Return value.ToString().Trim()
    End Function

End Class
