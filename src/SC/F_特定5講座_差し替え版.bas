' 貼り付け先: F_特定5講座（メインフォーム）のフォームモジュール【全文差し替え】
'   既存のコードを Ctrl+A で全選択して削除してから、このファイル全文を貼り付ける。
'
' 2026-07-24 の変更点（それ以外は客先の現行コードのまま）:
'   ・RequerySubForms に「一覧が0件ならチェック状況サブフォームを非表示にする」
'     処理を追加。0件の一覧に対してリンクが評価されると、下のキー項目が
'     ####・#エラー 表示になるため。フォームを開いたとき（Form_Load）も
'     日付を切り替えたときも必ずここを通るので、開いた直後の0件にも効く
'   ・非表示にする前に cbo対象日 へフォーカスを退避
'     （フォーカスのあるコントロールを含むサブフォームは非表示にできないため）
'   ・btn並び順リセット_Click を追加（2026-07-24）: ヘッダーで並べ替えた後に
'     既定の並び（未チェック上・No 順・チェック済み下）へ戻すボタン
'
' ★★★ 客先での準備（並び順リセットボタン）★★★
'   1. F_特定5講座 のヘッダー（cbo対象日 の近くなど）にコマンドボタンを1つ追加
'      （ウィザードが出たらキャンセル）
'   2. ボタンのプロパティを設定:
'        名前:     btn並び順リセット
'        標題:     並び順リセット
'   3. ボタンの「クリック時」イベントに [イベント プロシージャ] を選択
Option Compare Database
Option Explicit

' ★サブフォームコントロールの名前（フォーム名ではなく、メインフォーム上の
'   サブフォームコントロールの「名前」プロパティを確認して合わせる）
Private Const SUB_LIST As String = "F_SUB_特定5講座"
Private Const SUB_CHECK As String = "F_SUB_チェック状況"

' ★登録日フィールド名（両サブフォームのレコードソースにあるフィールド）
Private Const DATE_FIELD As String = "登録日"

' ★ドロップダウンの日付一覧を作る元（登録日を持つテーブルまたはクエリ）
Private Const DATE_SOURCE As String = "T_WORCS"

Private Sub cbo対象日_AfterUpdate()
    RequerySubForms
End Sub


'ヘッダーで並べ替えた後、既定の並び（未チェック上・No順・チェック済み下）に戻す
Private Sub btn並び順リセット_Click()
    On Error GoTo ErrHandler
    Me(SUB_LIST).Form.ResetOrder
    Exit Sub

ErrHandler:
    MsgBox "並び順のリセットに失敗しました。" & vbCrLf & _
        "明細サブフォームに ResetOrder（差し替え版のコード）が" & vbCrLf & _
        "貼り付けられているか確認してください。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Load()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strAllData As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT 商品コード, 商品名称 FROM M_特定講座 ORDER BY 商品コード")
    strAllData = ""
    Do Until rs.EOF
        strAllData = strAllData & rs![商品名称] & "（" & rs![商品コード] & "）　"
        rs.MoveNext
    Loop
    Me.lbl対象講座.Caption = "【対象講座】" & strAllData

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    'ドロップダウンに「データがある登録日」を新しい順で一覧表示
    Me!cbo対象日.RowSource = _
        "SELECT DISTINCT [登録日] FROM [" & DATE_SOURCE & "] " & _
        "ORDER BY [登録日] DESC"

    '初期表示は今日
    Me!cbo対象日.Value = Date
    RequerySubForms
End Sub


'一覧サブフォームを再読込し、0件ならチェック状況サブフォームを隠す
Private Sub RequerySubForms()
    On Error GoTo ErrHandler

    Dim hasData As Boolean

    Me(SUB_LIST).Form.Requery
    hasData = (Me(SUB_LIST).Form.RecordsetClone.RecordCount > 0)

    'フォーカスがチェック状況側にあると非表示にできないため先に退避
    If Not hasData Then Me!cbo対象日.SetFocus
    Me(SUB_CHECK).Visible = hasData
    Exit Sub

ErrHandler:
    MsgBox "サブフォームの再読込に失敗しました。" & vbCrLf & _
        "定数のサブフォームコントロール名を確認してください。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
