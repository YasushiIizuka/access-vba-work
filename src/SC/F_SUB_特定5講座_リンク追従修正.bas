' 貼り付け先: F_SUB_特定5講座（明細一覧サブフォーム）のフォームモジュール
' 内容: 明細の行を移動するたびに、チェック状況サブフォームへのリンクを張り直す
'   （サブフォーム間リンクの自動追従が働かなくなったため。イミディエイト
'     ウィンドウでの張り直しで同期することを 2026-07-24 に実機確認済み）
' 注意:
'   ・このモジュールに既に Form_Current がある場合は、Sub を丸ごと貼らず、
'     既存の Form_Current の末尾に中の2行（On Error〜Parent…の行）だけ追記する
'   ・先頭の Option 2行が既にある場合は貼らない
'   ・貼り付け後、フォームの「レコード移動時」イベントが
'     [イベント プロシージャ] になっていることを確認する
Option Compare Database
Option Explicit

Private Sub Form_Current()
    ' 単体で開いたときなど Parent が無い場合は何もしない
    On Error Resume Next
    Parent![F_SUB_チェック状況].LinkMasterFields = _
        Parent![F_SUB_チェック状況].LinkMasterFields
End Sub
