' 貼り付け先: F_SUB_特定5講座（明細一覧サブフォーム）のフォームモジュール【全文差し替え】
' 内容（2026-07-24 作り直し版）:
'   1. Form_Load: 並べ替えを設定（未チェックが上・チェック済みが下、同グループ内は No 順）
'      ※Yes/No 型は True=-1・False=0 のため、チェック済みを下にするには降順にする
'   2. 選択_AfterUpdate: チェックを保存して再読込 → チェックした行が一番下に落ちる
'      ※選択チェックボックスの「更新後処理」が [イベント プロシージャ] に
'        なっていることを確認すること
'   3. Form_Current: チェック状況サブフォームへのリンク張り直し
'      （リンクの自動追従が働かないため。0件時は何もしない。
'        0件時の表示制御はメインフォーム側の RequerySubForms が行う）
Option Compare Database
Option Explicit

Private Sub Form_Load()
    '未チェック→チェック済みの順（チェック済みが下）、同グループ内は No 順
    Me.OrderBy = "[選択] DESC, [No]"
    Me.OrderByOn = True
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
