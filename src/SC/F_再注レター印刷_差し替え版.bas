' 貼り付け先: F_再注レター印刷 のフォームモジュール【全文差し替え】
'
' ★★★ 貼り付け前にデザインビューで1点確認（重要）★★★
'   サブフォーム（F_SUB_再注レター印刷）のレコードソースが
'   「Q_再注レター印刷」（クエリ名そのもの）になっているか確認すること。
'   旧コードのバグ（下記）により「select * from Q_代引きレター印刷」が
'   保存されてしまっている可能性が高い。その場合はクエリ名に直す
'
' 2026-07-31 の変更点（それ以外は客先の現行コードのまま）:
'   ・【バグ修正】chk印刷済みFLG の切替が Q_代引きレター印刷 を SELECT していた
'     （コピー元の直し忘れ。再注の画面に代引きのデータが表示される）
'   ・印刷済みの表示切替を「レコードソースの差し替え」から「Filter の ON/OFF」に変更。
'     レコードソースを実行中に書き換えると閉じるときに保存され、次回開いたとき
'     チェックボックス（オフに戻る）と表示（全件のまま）がズレるため
'   ・Form_Load で毎回「チェックオフ＋未印刷のみ」に初期化（開いた状態を常に一定に）
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
    '2件目の本文を設定する
    Me.Recordset.FindFirst "[ID] = 2"
End Sub

Private Sub Form_Load()
    '開いたときは必ず「未印刷のみ表示」に統一
    Me!chk印刷済みFLG = False
    ApplyPrintedFilter
End Sub

Private Sub chk印刷済みFLG_AfterUpdate()
    ApplyPrintedFilter
End Sub

'チェックボックスの状態に合わせてサブフォームの絞り込みを適用する
Private Sub ApplyPrintedFilter()
    With Me![F_SUB_再注レター印刷].Form
        If Me!chk印刷済みFLG = True Then
            '印刷済みも表示（絞り込み解除）
            .FilterOn = False
        Else
            '未印刷のみ表示
            .Filter = "[印刷済みFLG] = False"
            .FilterOn = True
        End If
    End With
End Sub

Private Sub btn再注レター印刷_Click()
    Dim db As DAO.Database
    Dim rsSub As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim str本文元 As String
    Dim str本文 As String

    Set db = CurrentDb

    CurrentDb.Execute "delete from T_レター出力", dbFailOnError

    str本文元 = Nz(Me.txt本文元.Value, "")

    If str本文元 = "" Then
        MsgBox "本文が空です。", vbExclamation
        Exit Sub
    End If

    'サブフォームのレコードを取得
    Set rsSub = Me!F_SUB_再注レター印刷.Form.RecordsetClone

    If rsSub.RecordCount = 0 Then
        MsgBox "印刷対象のデータがありません", vbExclamation
        Exit Sub
    End If

    Set rsOut = db.OpenRecordset("T_レター出力", dbOpenDynaset)

    rsSub.MoveFirst

    Do Until rsSub.EOF
        If Nz(rsSub!印刷済みFLG, False) = False Then
            str本文 = str本文元
            rsOut.AddNew
            rsOut![No] = rsSub![No]
            rsOut![本文] = str本文
            rsOut![住所1] = Nz(rsSub![都道府県], "") & Nz(rsSub![市区郡], "")
            rsOut![住所2] = Nz(rsSub![町村番地], "") & Nz(rsSub![建物名], "")
            rsOut![氏名] = Nz(rsSub![顧客姓_漢字], "") & " " & Nz(rsSub![顧客名_漢字], "")
            rsOut![作成日時] = Now()
            rsOut.Update
        End If
        rsSub.MoveNext
    Loop
    DoCmd.OpenReport "R_再注レター_特定5講座", acViewPreview
End Sub
