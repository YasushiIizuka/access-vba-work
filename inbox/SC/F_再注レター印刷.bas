Option Compare Database
Option Explicit


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

Private Sub chk印刷済みFLG_AfterUpdate()
    Dim sql As String
    If Me!chk印刷済みFLG = True Then
        sql = "select * from Q_代引きレター印刷"
    Else
        sql = "select * from Q_代引きレター印刷 where 印刷済みFLG = False"
    End If
    Me![F_SUB_再注レター印刷].Form.FilterOn = False
    Me![F_SUB_再注レター印刷].Form.RecordSource = sql
    Me![F_SUB_再注レター印刷].Form.Requery
End Sub

Private Sub Form_Open(Cancel As Integer)
    '2件目の本文を設定する
    Me.Recordset.FindFirst "[ID] = 2"
End Sub
