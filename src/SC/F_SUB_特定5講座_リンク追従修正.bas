' 貼り付け先: F_SUB_特定5講座（明細一覧サブフォーム）のフォームモジュール【全文差し替え】
' 内容: 明細の行を移動するたびに、チェック状況サブフォームへのリンクを張り直す
'   （リンクの自動追従が働かないため。2026-07-24 実機確認済み）
'   0件のときは何もしない（現在行が無い状態でリンクを評価するとエラーになるため。
'   0件時の表示制御はメインフォーム側の RequerySubForms が行う）
Option Compare Database
Option Explicit

Private Sub Form_Current()
    ' 単体で開いたときなど Parent が無い場合は何もしない
    On Error Resume Next

    ' 0件なら張り直さない（表示のON/OFFはメインフォーム側で制御）
    If Me.RecordsetClone.RecordCount = 0 Then Exit Sub

    ' 詳細原因不明だが、再設定しないとサブフォームのほうが追従しない
    Parent![F_SUB_チェック状況].LinkMasterFields = _
        Parent![F_SUB_チェック状況].LinkMasterFields
End Sub
