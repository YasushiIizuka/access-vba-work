' 貼り付け先: F_特定5講座 のコードビハインド（フォームモジュール）に追記
'
' =============================================================
' 登録日によるサブフォームの絞り込み
'
' 動き:
'   ・フォームを開くと「今日」の登録日のデータだけを2つのサブフォームに表示
'   ・ヘッダーのドロップダウン（データがある登録日の一覧）で別日付に切替可能
'   ・ドロップダウンを空にすると全件表示に戻る
'
' ★★★ 客先での準備（デザインビューでの作業） ★★★
'   1. F_特定5講座 のヘッダーにコンボボックスを1つ追加
'      （ウィザードが出たらキャンセル）
'   2. コンボボックスのプロパティを設定:
'        名前:             cbo対象日
'        書式:             日付 (S)
'        入力チェック:     いいえ（LimitToList=No。一覧に無い日付も入力可に）
'      ※値集合ソースはコードが自動設定するので空欄でよい
'   3. コンボボックスの「更新後処理」イベントに [イベント プロシージャ] を選択
'   4. フォームの「読み込み時」イベントに [イベント プロシージャ] を選択
'      （既に Form_Load がある場合は、その末尾に本コードの Form_Load の
'        中身を移して1つにまとめる）
'   5. 下の定数4つを実環境に合わせて修正
' =============================================================

' ★サブフォームコントロールの名前（フォーム名ではなく、メインフォーム上の
'   サブフォームコントロールの「名前」プロパティを確認して合わせる）
Private Const SUB_LIST As String = "F_SUB_特定5講座"
Private Const SUB_CHECK As String = "F_SUB_特定5講座_チェック状況"

' ★登録日フィールド名（両サブフォームのレコードソースにあるフィールド）
Private Const DATE_FIELD As String = "登録日"

' ★ドロップダウンの日付一覧を作る元（登録日を持つテーブルまたはクエリ）
Private Const DATE_SOURCE As String = "<<T_WORCSまたは絞り込みクエリ名>>"

Private Sub Form_Load()
    'ドロップダウンに「データがある登録日」を新しい順で一覧表示
    Me!cbo対象日.RowSource = _
        "SELECT DISTINCT DateValue([" & DATE_FIELD & "]) AS 対象日 " & _
        "FROM [" & DATE_SOURCE & "] " & _
        "ORDER BY DateValue([" & DATE_FIELD & "]) DESC"

    '初期表示は今日
    Me!cbo対象日.Value = Date
    ApplyDateFilter
End Sub

Private Sub cbo対象日_AfterUpdate()
    ApplyDateFilter
End Sub

'選択された日付で2つのサブフォームを絞り込む（空なら全件表示）
Private Sub ApplyDateFilter()
    If IsNull(Me!cbo対象日.Value) Or Me!cbo対象日.Value = "" Then
        '全件表示
        SetSubFilter SUB_LIST, ""
        SetSubFilter SUB_CHECK, ""
    Else
        '登録日は Now()（時刻付き）で入っているため、= ではなく範囲で比較する
        Dim d As Date
        d = CDate(Me!cbo対象日.Value)

        Dim f As String
        f = "[" & DATE_FIELD & "] >= #" & Format(d, "mm\/dd\/yyyy") & "#" & _
            " AND [" & DATE_FIELD & "] < #" & Format(d + 1, "mm\/dd\/yyyy") & "#"

        SetSubFilter SUB_LIST, f
        SetSubFilter SUB_CHECK, f
    End If
End Sub

'サブフォームにフィルタを適用する（空文字なら解除）
Private Sub SetSubFilter(ByVal ctlName As String, ByVal filterStr As String)
    On Error GoTo ErrHandler

    With Me(ctlName).Form
        If Len(filterStr) = 0 Then
            .FilterOn = False
            .Filter = ""
        Else
            .Filter = filterStr
            .FilterOn = True
        End If
    End With
    Exit Sub

ErrHandler:
    MsgBox "サブフォーム「" & ctlName & "」の絞り込みに失敗しました。" & vbCrLf & _
        "コントロール名・フィールド名の定数を確認してください。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
