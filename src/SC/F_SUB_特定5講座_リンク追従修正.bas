' 貼り付け先: F_SUB_特定5講座（明細一覧サブフォーム）のフォームモジュール【全文差し替え】
' 内容:
'   ・明細の行を移動するたびに、チェック状況サブフォームへのリンクを張り直す
'     （リンクの自動追従が働かないため。2026-07-24 実機確認済み）
'   ・一覧が0件のとき（対象日にデータが無い日）はチェック状況サブフォームを
'     非表示にする。0件の一覧に対してリンクを評価すると、参照先の現在行が
'     無いため下のキー項目が ####・#エラー 表示になるのを防ぐ（2026-07-24 追加）
Option Compare Database
Option Explicit

Private Sub Form_Current()
    ' 単体で開いたときなど Parent が無い場合は何もしない
    On Error Resume Next

    If Me.RecordsetClone.RecordCount = 0 Then
        ' 対象日にデータが無い: 下のチェック状況を隠す（エラー表示防止）
        Parent![F_SUB_チェック状況].Visible = False
    Else
        Parent![F_SUB_チェック状況].Visible = True
        ' 詳細原因不明だが、再設定しないとサブフォームのほうが追従しない
        Parent![F_SUB_チェック状況].LinkMasterFields = _
            Parent![F_SUB_チェック状況].LinkMasterFields
    End If
End Sub
