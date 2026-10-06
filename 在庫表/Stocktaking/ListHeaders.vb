''' <summary>
''' 一覧シートの見出し・目印の文字列。
''' 出力（StocktakingWriter）と棚番の Excel 読み込み（ShelfExcelImporter）で共通に使う。
''' </summary>
Public NotInheritable Class ListHeaders
    Private Sub New()
    End Sub

    ''' <summary>一覧シート名の末尾（"08.小山営業所 在庫一覧"）</summary>
    Public Const SHEET_SUFFIX As String = " 在庫一覧"
    Public Const HEADER_ROW As Integer = 2

    Public Const NO As String = "No"
    Public Const CUSTOMER As String = "得意先"
    Public Const ITEM_CODE As String = "商品CD"
    Public Const MODEL As String = "品名"
    Public Const REGISTERED As String = "登録数"
    Public Const SHELF As String = "棚番"
    Public Const NOTE As String = "備考"
    ''' <summary>非表示列。AXIS 内部の得意先CD（短縮CDは一意でないため）</summary>
    Public Const CUSTOMER_CODE As String = "得意先CD"
    ''' <summary>非表示列。出力時の並び（並べ替え後に元へ戻す用）</summary>
    Public Const ORDER As String = "並び順"

    ''' <summary>棚区分：この行の棚番が共用棚か得意先専用棚か。単独行を「専用」にするとその得意先の専用棚として読み込む</summary>
    Public Const KIND As String = "棚区分"
    Public Const KIND_SHARED As String = "共用"
    Public Const KIND_DEDICATED As String = "専用"

    ''' <summary>No 列：子行・2行目以降の対象外行</summary>
    Public Const CHILD_NO As String = "-"
    ''' <summary>No 列：追記用の空行</summary>
    Public Const APPEND_NO As String = "追記"
    ''' <summary>備考列：棚卸対象外ステータスの行</summary>
    Public Const EXCLUDED_NOTE As String = "棚卸対象外"
End Class
