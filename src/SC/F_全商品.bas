' 貼り付け先: F_全商品（メインフォーム）のフォームモジュール【全文差し替え】
'   既存のコードを Ctrl+A で全選択して削除してから、このファイル全文を貼り付ける。
'   貼り付け後にデバッグ→コンパイルでエラーがないことを確認する。
'
' F_特定5講座 の 2026-07-31 版（動作確認済み）を全商品向けに調整したもの（2026-08-04）:
'   ・サブフォームコントロール名の定数を F_SUB_全商品 に変更
'   ・対象日の記憶キーを TempVars("全商品_対象日") に変更（特定5講座と別管理）
'   ・Form_Load の対象講座ラベル（M_特定講座 の一覧表示）を廃止。
'     ラベル lbl対象講座 を残している場合のみ「【対象】全商品」に文言を更新する
'
' ★★★ 貼り付け前の確認（コピーで作ったフォームは参照が旧のまま残りやすい）★★★
'   1. F_全商品 上の一覧側サブフォームコントロール:
'        名前 = F_SUB_全商品 ／ ソースオブジェクト = F_SUB_全商品
'   2. チェック状況側サブフォームコントロール:
'        名前 = F_SUB_チェック状況 ／ ソースオブジェクト = F_SUB_全商品_チェック状況
'   3. F_SUB_全商品 のレコードソース = Q_全商品
'      F_SUB_全商品_チェック状況 のレコードソースも全商品用になっているか
'   4. 実際の名前が上と違う場合は、下の定数を実際の名前に合わせる
Option Compare Database
Option Explicit

' ★サブフォームコントロールの名前（フォーム名ではなく、メインフォーム上の
'   サブフォームコントロールの「名前」プロパティを確認して合わせる）
Private Const SUB_LIST As String = "F_SUB_全商品"
Private Const SUB_CHECK As String = "F_SUB_チェック状況"

' ★登録日フィールド名（両サブフォームのレコードソースにあるフィールド）
Private Const DATE_FIELD As String = "登録日"

' ★ドロップダウンの日付一覧を作る元（登録日を持つテーブルまたはクエリ）
Private Const DATE_SOURCE As String = "T_WORCS"

Private Sub cbo対象日_AfterUpdate()
    '選んだ日付を記憶（ナビゲーションのタブを離れてもセッション内は保持される）
    TempVars("全商品_対象日") = Me!cbo対象日.Value
    RequerySubForms
End Sub


'ヘッダーで並べ替えた後、既定の並び（未チェック上・No順・チェック済み下）に戻す
Private Sub btn並び順リセット_Click()
    On Error GoTo ErrHandler
    Me(SUB_LIST).Form.ResetOrder
    Exit Sub

ErrHandler:
    MsgBox "並び順のリセットに失敗しました。" & vbCrLf & _
        "明細サブフォームに ResetOrder（全商品版のコード）が" & vbCrLf & _
        "貼り付けられているか確認してください。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Load()
    '対象講座ラベルは使わない（全商品が対象）。
    'ラベル lbl対象講座 をフォームに残している場合のみ文言を更新する
    '（ラベルを削除済みならこの2行は何もしない）
    On Error Resume Next
    Me.lbl対象講座.Caption = "【対象】全商品"
    On Error GoTo 0

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
    lastDate = TempVars("全商品_対象日").Value
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
'  クエリに [Forms]![F_全商品]![cbo対象日] と書く方式は、ナビゲーション
'  フォームに埋め込むと参照が解決できず「パラメーターの入力」が出るためNG
'  （Q_全商品 の登録日の抽出条件は空にしておくこと）
Private Sub RequerySubForms()
    On Error GoTo ErrHandler

    Dim hasData As Boolean

    If IsDate(Me!cbo対象日) Then
        'フィールド名はテーブル名で修飾する（Q_全商品 は T_WORCS と
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
