' 貼り付け先: F_代引きレター印刷 のフォームモジュール【全文差し替え】
'
' ★★★ 貼り付け前にデザインビューで1点確認 ★★★
'   サブフォーム（F_SUB_代引きレター印刷）のレコードソースが
'   「Q_代引きレター印刷」（クエリ名そのもの）になっているか確認し、
'   SELECT 文が保存されてしまっていたらクエリ名に戻すこと
'
' 2026-07-31 の変更点（それ以外は客先の現行コードのまま）:
'   ・印刷済みの表示切替を「レコードソースの差し替え」から「Filter の ON/OFF」に変更。
'     レコードソースを実行中に書き換えると閉じるときに保存され、次回開いたとき
'     チェックボックス（オフに戻る）と表示（全件のまま）がズレるため
'   ・Form_Load で毎回「チェックオフ＋未印刷のみ」に初期化（開いた状態を常に一定に）
Option Compare Database
Option Explicit

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
    With Me![F_SUB_代引きレター印刷].Form
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

Private Sub btn代引きレター印刷_Click()
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
    Set rsSub = Me!F_SUB_代引きレター印刷.Form.RecordsetClone

    If rsSub.RecordCount = 0 Then
        MsgBox "印刷対象のデータがありません", vbExclamation
        Exit Sub
    End If

    Set rsOut = db.OpenRecordset("T_レター出力", dbOpenDynaset)

    rsSub.MoveFirst

    Do Until rsSub.EOF
        If Nz(rsSub!印刷済みFLG, False) = False Then
            str本文 = str本文元
            str本文 = Replace(str本文, "{商品名称}", Nz(rsSub!商品名称, ""))
            str本文 = Replace(str本文, "{返送期限}", Nz(rsSub!代引きレター返送期限, ""))
            rsOut.AddNew
            rsOut![No] = rsSub![No]
            rsOut![本文] = str本文
            rsOut![住所1] = Nz(rsSub![都道府県], "") & Nz(rsSub![市区郡], "")
            rsOut![住所2] = Nz(rsSub![町村番地], "") & Nz(rsSub![建物名], "")
            rsOut![氏名] = Nz(rsSub![顧客姓_漢字], "") & " " & Nz(rsSub![顧客名_漢字], "")
            rsOut![お申込商品] = "「" & Nz(rsSub![商品名称], "") & "」を代金引換払いで申し込みます。"
            rsOut![作成日時] = Now()
            rsOut.Update
        End If
        rsSub.MoveNext
    Loop
    DoCmd.OpenReport "R_代引きレター_特定5講座", acViewPreview
End Sub
