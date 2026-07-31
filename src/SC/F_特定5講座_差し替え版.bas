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
    '選んだ日付を記憶（ナビゲーションのタブを離れてもセッション内は保持される）
    TempVars("特定5講座_対象日") = Me!cbo対象日.Value
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

    '初期表示: このセッションで選んだ日付があればそれ、なければ今日
    '（ナビゲーションフォームはタブを離れるとフォームを閉じるため、
    '  タブを戻ったときに Form_Load が再実行される。TempVars に覚えて
    '  おいた日付を復元することで、選択日がリセットされないようにする）
    Dim lastDate As Variant
    lastDate = Null
    On Error Resume Next
    lastDate = TempVars("特定5講座_対象日").Value
    On Error GoTo 0

    If IsDate(lastDate) Then
        Me!cbo対象日.Value = lastDate
    Else
        Me!cbo対象日.Value = Date
    End If
    RequerySubForms
End Sub


'一覧サブフォームを対象日で絞り込んで再読込し、0件ならチェック状況サブフォームを隠す
'※絞り込みはクエリの抽出条件ではなくここでフィルタとして適用する。
'  クエリに [Forms]![F_特定5講座]![cbo対象日] と書く方式は、ナビゲーション
'  フォームに埋め込むと参照が解決できず「パラメーターの入力」が出るため廃止
'  （2026-07-31。Q_特定5講座 の登録日の抽出条件は削除すること）
Private Sub RequerySubForms()
    On Error GoTo ErrHandler

    Dim hasData As Boolean

    If IsDate(Me!cbo対象日) Then
        'フィールド名はテーブル名で修飾する（Q_特定5講座 は T_WORCS と
        ' T_WORCS_Check を結合しており、両方に登録日があるため、修飾しないと
        ' エラー3079「複数のテーブルを参照しました」になる）
        Me(SUB_LIST).Form.Filter = "[" & DATE_SOURCE & "].[" & DATE_FIELD & "] = " & _
            Format$(Me!cbo対象日, "\#yyyy\/mm\/dd\#")
    Else
        '対象日が未指定のときは何も表示しない
        Me(SUB_LIST).Form.Filter = "1 = 0"
    End If
    Me(SUB_LIST).Form.FilterOn = True

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
