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

    ''' <summary>
    ''' 専用欄：専用棚の行に○を出す。読み込みでは空欄以外（○・〇・◯・O・1 など何でも）を専用とみなす。
    ''' 内訳行に入れて棚番を書けば専用棚に昇格、専用棚の行で消せば共用棚に戻す
    ''' </summary>
    Public Const MARK As String = "専用"
    Public Const MARK_VALUE As String = "○"

    ''' <summary>非表示列。行の種類（並べ替え後も判別できるように持つ）</summary>
    Public Const ROW_TYPE As String = "行種別"
    ''' <summary>集計対象：共用棚の行</summary>
    Public Const ROW_SHARED As String = "共用"
    ''' <summary>集計対象：専用棚の行</summary>
    Public Const ROW_DEDICATED As String = "専用"
    ''' <summary>集計対象外：得意先別の内訳行</summary>
    Public Const ROW_DETAIL As String = "内訳"
    ''' <summary>集計対象外：棚卸対象外ステータスの行</summary>
    Public Const ROW_EXCLUDED As String = "対象外"

    ''' <summary>No 列：旧形式の子行（読み込みの互換用）</summary>
    Public Const CHILD_NO As String = "-"
    ''' <summary>No 列：追記用の空行</summary>
    Public Const APPEND_NO As String = "追記"
    ''' <summary>備考列：棚卸対象外ステータスの行</summary>
    Public Const EXCLUDED_NOTE As String = "棚卸対象外"
End Class
