' 貼り付け先: F_SUB_特定5講座（明細一覧サブフォーム）のフォームモジュール【全文差し替え】
' 内容（2026-07-24 作り直し版）:
'   1. Form_Load: 並べ替えを設定（未チェックが上・チェック済みが下、同グループ内は No 順）
'      ※Yes/No 型は True=-1・False=0 のため、チェック済みを下にするには降順にする
'   2. Form_Load: 条件付き書式をコードで設定（選択＝True の行を黒っぽいグレーに）
'      ※各テキストボックスの既存の条件付き書式は毎回消してから設定し直す。
'        デザイン時に手で付けた条件付き書式は残らないので、書式は全てこのコードで管理する。
'        チェックボックス自体は条件付き書式が使えないため、選択列のセルだけは色が付かない
'   3. 選択_AfterUpdate: チェックを保存して再読込 → チェックした行が一番下に落ちる
'      ※選択チェックボックスの「更新後処理」が [イベント プロシージャ] に
'        なっていることを確認すること
'   4. Form_Current: チェック状況サブフォームへのリンク張り直し
'      （リンクの自動追従が働かないため。0件時は何もしない。
'        0件時の表示制御はメインフォーム側の RequerySubForms が行う）
Option Compare Database
Option Explicit

'チェック済み行の色（&HBBGGRR 形式。グレーは RGB の3値を同じにする）
'  もっと明るく: &HC0C0C0（その場合は文字を黒 &H0 に）／もっと暗く: &H404040
Private Const CHECKED_BACK_COLOR As Long = &H808080   '中間のグレー RGB(128,128,128)
Private Const CHECKED_FORE_COLOR As Long = &HFFFFFF   '文字は白

Private Sub Form_Load()
    '未チェック→チェック済みの順（チェック済みが下）、同グループ内は No 順
    Me.OrderBy = "[選択] DESC, [No]"
    Me.OrderByOn = True

    'チェック済み行の色つけ（条件付き書式をコードで設定）
    SetupRowFormat
End Sub

'選択＝True の行に色を付ける条件付き書式を全テキストボックスに設定する
Private Sub SetupRowFormat()
    Dim ctl As Control
    Dim fc As FormatCondition

    'フォーム上の全コントロールから対象を型で絞る
    '（Me.Detail はセクション名が「詳細」の日本語環境では使えないため）
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            '重複防止のため既存の条件付き書式を消してから設定
            ctl.FormatConditions.Delete
            Set fc = ctl.FormatConditions.Add(acExpression, , "[選択]=True")
            fc.BackColor = CHECKED_BACK_COLOR
            fc.ForeColor = CHECKED_FORE_COLOR
        End If
    Next ctl
End Sub

Private Sub 選択_AfterUpdate()
    'チェックを保存してから再読込 → 並べ替えが効いて行が下に落ちる
    If Me.Dirty Then Me.Dirty = False
    Me.Requery
End Sub

Private Sub Form_Current()
    ' 単体で開いたときなど Parent が無い場合は何もしない
    On Error Resume Next

    ' 0件なら張り直さない（表示のON/OFFはメインフォーム側で制御）
    If Me.RecordsetClone.RecordCount = 0 Then Exit Sub

    ' 詳細原因不明だが、再設定しないとサブフォームのほうが追従しない
    Parent![F_SUB_チェック状況].LinkMasterFields = _
        Parent![F_SUB_チェック状況].LinkMasterFields
End Sub
