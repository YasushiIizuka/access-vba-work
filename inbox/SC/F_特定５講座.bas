Option Compare Database
Option Explicit

Private Function PassCreate() As String
    Dim allowedChars As String
    Dim i As Integer
    Dim pos As Integer
    Dim newPassword As String
    
    '間違いやすい文字は除く
    allowedChars = "ABDEFGHJMNQRTUYabdefghjmnrtuy"
    
    '乱数初期化
    Randomize
    
    '6桁のパスワード生成
    newPassword = ""
    For i = 0 To 6
        pos = Int(Rnd * Len(allowedChars)) + 1
        newPassword = newPassword & Mid(allowedChars, pos, 1)
    Next i
    PassCreate = newPassword

End Function

Private Sub btnPassCreate_Click()
    'テキストボックスに代入
    Me.パスワード生成.Value = PassCreate
    
End Sub

Private Sub btn再注レター作成_Click()
    Dim msg As String
    Dim 再注レター返送期限 As Date
    Dim 事業部 As String
    Dim 事業部TEL As String
    Dim sql As String
    Dim 未成年FLG As Boolean
    Dim 重複件数 As Long
    Dim パスワード As String
    
    重複件数 = DCount("*", "T_再注レター", "[No] = " & Me!No & " AND [印刷済みFLG] = False")
    
    If 重複件数 > 0 Then
        MsgBox "このデータはすでに印刷待ちに登録されています。", vbExclamation, "登録重複"
        Exit Sub
    End If
     
    再注レター返送期限 = Date + 10 '今日日付＋10日
    
    If Parent![F_SUB_特定5講座].Form![ショップ区分] = "通教" Then
        事業部 = "U-CAN"
        事業部TEL = "03-5388-6111"
    Else
        事業部 = "ライフ＆カルチャー"
        事業部TEL = "0120-552-476"
    End If
    
    If IsNull(Parent![F_SUB_特定5講座].Form![年齢]) Then
        未成年FLG = False
    ElseIf Parent![F_SUB_特定5講座].Form![年齢] < 18 Then
        未成年FLG = True
    Else
        未成年FLG = False
    End If
    パスワード = PassCreate
    
    msg = "【確認】次の情報で再注レター印刷情報を登録します" & vbCrLf & vbCrLf
    msg = msg & "No：" & Parent![F_SUB_特定5講座].Form![No] & vbCrLf
    msg = msg & "郵便番号：" & Parent![F_SUB_特定5講座].Form![郵便番号] & vbCrLf
    msg = msg & "住所：" & Parent![F_SUB_特定5講座].Form![都道府県] & Parent![F_SUB_特定5講座].Form![市区郡] & Parent![F_SUB_特定5講座].Form![町村番地] & Parent![F_SUB_特定5講座].Form![建物名] & vbCrLf
    msg = msg & "顧客名：" & Parent![F_SUB_特定5講座].Form![顧客姓_漢字] & " " & Parent![F_SUB_特定5講座].Form![顧客名_漢字] & vbCrLf
    msg = msg & "顧客番号：" & Parent![F_SUB_特定5講座].Form![顧客番号] & vbCrLf
    msg = msg & "事業部：" & 事業部 & vbCrLf
    msg = msg & "事業部TEL：" & 事業部TEL & vbCrLf
    msg = msg & "商品名：" & Parent![F_SUB_特定5講座].Form![商品名称] & vbCrLf
    msg = msg & "再注レター返送期限：" & 再注レター返送期限 & vbCrLf
    msg = msg & "パスワード：" & パスワード & vbCrLf
    
    
    If MsgBox(msg, vbYesNo + vbQuestion, "代引きレター印刷情報の登録") = vbYes Then
        Me!再注レター返送期限 = 再注レター返送期限
        Me!事業部 = 事業部
        Me!事業部TEL = 事業部TEL
        Me!パスワード生成 = パスワード
        sql = "INSERT INTO [T_再注レター] ([No], [未成年FLG]) VALUES (" & Me!No & ", " & 未成年FLG & " );"
        Parent![F_SUB_特定5講座].Form![再注レター] = True
        CurrentDb.Execute sql, dbFailOnError
        If 未成年FLG Then
            MsgBox "印刷情報を未成年（" & Parent![F_SUB_特定5講座].Form![年齢] & "歳）として登録しました"
        Else
            MsgBox "印刷情報を登録しました"
        End If
    Else
        '何もしないでキャンセル
    End If

End Sub

Private Sub btn代引きレター作成_Click()
    Dim msg As String
    Dim 代引きレター返送期限 As Date
    Dim 事業部 As String
    Dim 事業部TEL As String
    Dim sql As String
    Dim 未成年FLG As Boolean
    Dim 重複件数 As Long
    
    重複件数 = DCount("*", "T_代引きレター", "[No] = " & Me!No & " AND [印刷済みFLG] = False")
    
    If 重複件数 > 0 Then
        MsgBox "このデータはすでに印刷待ちに登録されています。", vbExclamation, "登録重複"
        Exit Sub
    End If
     
    代引きレター返送期限 = Date + 20 '今日日付＋20日
    
    If Parent![F_SUB_特定5講座].Form![ショップ区分] = "通教" Then
        事業部 = "U-CAN"
        事業部TEL = "03-5388-6111"
    Else
        事業部 = "ライフ＆カルチャー"
        事業部TEL = "0120-552-476"
    End If
    
    If IsNull(Parent![F_SUB_特定5講座].Form![年齢]) Then
        未成年FLG = False
    ElseIf Parent![F_SUB_特定5講座].Form![年齢] < 18 Then
        未成年FLG = True
    Else
        未成年FLG = False
    End If
    
    msg = "【確認】次の情報で代引きレター印刷情報を登録します" & vbCrLf & vbCrLf
    msg = msg & "No：" & Parent![F_SUB_特定5講座].Form![No] & vbCrLf
    msg = msg & "郵便番号：" & Parent![F_SUB_特定5講座].Form![郵便番号] & vbCrLf
    msg = msg & "住所：" & Parent![F_SUB_特定5講座].Form![都道府県] & Parent![F_SUB_特定5講座].Form![市区郡] & Parent![F_SUB_特定5講座].Form![町村番地] & Parent![F_SUB_特定5講座].Form![建物名] & vbCrLf
    msg = msg & "顧客名：" & Parent![F_SUB_特定5講座].Form![顧客姓_漢字] & " " & Parent![F_SUB_特定5講座].Form![顧客名_漢字] & vbCrLf
    msg = msg & "顧客番号：" & Parent![F_SUB_特定5講座].Form![顧客番号] & vbCrLf
    msg = msg & "事業部：" & 事業部 & vbCrLf
    msg = msg & "事業部TEL：" & 事業部TEL & vbCrLf
    msg = msg & "商品名：" & Parent![F_SUB_特定5講座].Form![商品名称] & vbCrLf
    msg = msg & "代引きレター返送期限：" & 代引きレター返送期限 & vbCrLf
    
    
    If MsgBox(msg, vbYesNo + vbQuestion, "代引きレター印刷情報の登録") = vbYes Then
        Me!代引きレター返送期限 = 代引きレター返送期限
        Me!事業部 = 事業部
        Me!事業部TEL = 事業部TEL
        sql = "INSERT INTO [T_代引きレター] ([No], [未成年FLG]) VALUES (" & Me!No & ", " & 未成年FLG & " );"
        Parent![F_SUB_特定5講座].Form![代引きレター] = True
        CurrentDb.Execute sql, dbFailOnError
        If 未成年FLG Then
            MsgBox "印刷情報を未成年（" & Parent![F_SUB_特定5講座].Form![年齢] & "歳）として登録しました"
        Else
            MsgBox "印刷情報を登録しました"
        End If
    Else
        '何もしないでキャンセル
    End If
    

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me![受注番号] = Parent![F_SUB_特定5講座].Form![受注番号]
    Me![商品コード] = Parent![F_SUB_特定5講座].Form![商品コード]
    Me![登録日] = Now()
    Me![更新日] = Now()
End Sub

' 2つのコンボの状態を現在値に合わせて更新する（共通処理）
Private Sub UpdateLetterComboState()
    Dim hasDaibiki As Boolean
    Dim hasSaichu As Boolean

    ' Null・空文字のどちらも「空」とみなす
    hasDaibiki = (Nz(Me.代引きレター返送状況.Value, "") <> "")
    hasSaichu = (Nz(Me.再注レターの返送状況.Value, "") <> "")

    ' フォーカスが乗っているコントロールは無効化できない（エラー2164）ため、
    ' これから無効化する側にフォーカスがある場合だけ相手側へ退避する
    On Error Resume Next
    If hasDaibiki And Me.ActiveControl.Name = "再注レターの返送状況" Then
        Me.代引きレター返送状況.SetFocus
    End If
    If hasSaichu And Me.ActiveControl.Name = "代引きレター返送状況" Then
        Me.再注レターの返送状況.SetFocus
    End If
    On Error GoTo 0

    Me.再注レターの返送状況.Enabled = Not hasDaibiki
    Me.代引きレター返送状況.Enabled = Not hasSaichu
End Sub
Private Sub 再注レターの返送状況_AfterUpdate()
    UpdateLetterComboState
End Sub

Private Sub 代引きレター返送状況_AfterUpdate()
    UpdateLetterComboState
End Sub

' 連結フォームの場合: レコード移動のたびに、その行の値に合わせて状態を揃える
Private Sub Form_Current()
    UpdateLetterComboState
End Sub

