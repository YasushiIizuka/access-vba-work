Option Compare Database

Private Sub btnテキスト追加編集_Click()
    DoCmd.OpenTable "メッセージ"
End Sub

Private Sub btnリフレッシュ_Click()
    Me.Refresh
End Sub

Private Sub btn配信済み_Click()
    On Error GoTo ErrHandler
    
    Dim fd As Object
    Dim filePathCSV As Variant
    Dim fileNo As Integer
    Dim lineData As String
    Dim csvData() As String
    Dim targetClumn As Integer
    Dim totalRows As Long
    Dim currentRow As Long
    Dim returnValue As Variant
    Dim mSMS配信日 As String
    Dim startTime, endTime, processTime As Variant
    Dim m携帯電話番号 As String
    Dim m送信ステータス As String
    Dim strSQL As String
    Dim db As DAO.Database
    
    'エクセル関係
    Dim filePathXLSX As Variant
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim foundCell As Object
    Dim m顧客番号 As String
    
    Set db = CurrentDb
    
    mSMS配信日 = InputBox("配信日を入力してください", "配信日", Format(Date, "yyyy/mm/dd"))
    If mSMS配信日 = "" Or IsDate(mSMS配信日) = False Then
        MsgBox "配信日の設定が正しくありません"
        Exit Sub
    End If
    
    '対象列番号
    targetClumn = 3
    
    'ファイル選択ダイアログを表示
    Set fd = Application.FileDialog(3) '3はダイアログピッカー
    With fd
        .Title = "CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        If .Show = True Then
            filePathCSV = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    'エクセルファイル選択ダイアログを表示
    With fd
        .Title = "エクセルファイルを選択してください"
        .Filters.Clear
        .Filters.Add "XLSXファイル", "*.xlsx"
        .AllowMultiSelect = False
        If .Show = True Then
            filePathXLSX = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    'エクセルを起動して対象ファイルを開く
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(filePathXLSX, ReadOnly:=True)
    Set xlSheet = xlBook.Worksheets(1)
          
    startTime = Timer
    
    '全行を数える
    fileNo = FreeFile
    Open filePathCSV For Input As #fileNo
    totalRow = 0
    Do Until EOF(fileNo)
        Line Input #fileNo, lineData
        If Trim(lineData) <> "" Then totalRows = totalRows + 1
    Loop
    Close #fileNo
    
    'ヘッダーを引く（データが空ならば削除）
    If totalRows <= 1 Then
        MsgBox "データが見つかりません"
        Exit Sub
    End If
    totalRows = totalRows - 1 'ヘッダーを除いた純粋なデータ数
    
    '進捗バーの初期化
    returnValue = SysCmd(acSysCmdInitMeter, "配信日を設定中…", totalRows)
    
    'データの読み込み
    fileNo = FreeFile
    Open filePathCSV For Input As #fileNo
    
    'ヘッダー行を1行読み飛ばす
    If Not EOF(fileNo) Then
        Line Input #fileNo, lineData
    End If
    
    currentRow = 0
    Do Until EOF(fileNo)
        Line Input #fileNo, lineData
        If Trim(lineData) <> "" Then
            currentRow = currentRow + 1
            csvData = Split(lineData, ",")
            '携帯電話番号の取得
            If UBound(csvData) >= targetClumn - 1 Then
                m携帯電話番号 = Trim(Replace(csvData(targetClumn - 1), """", ""))
                'エクセルから携帯電話番号をキーに顧客番号を取得する（-4163はlValues, 1はxlWhole完全一致）
                Set foundCell = xlSheet.Columns("D").Find(What:=m携帯電話番号, LookIn:=-4163, LookAt:=1)
                If Not foundCell Is Nothing Then
                    m顧客番号 = xlSheet.Cells(foundCell.Row, "A").Value
                    strSQL = "UPDATE 顧客 " & _
                            " SET SMS配信日 = #" & mSMS配信日 & "# " & _
                            "WHERE 顧客番号 = '" & m顧客番号 & "';"
                    Debug.Print strSQL
                    db.Execute strSQL, dbFailOnError
                Else
                    MsgBox m携帯電話番号 & "に対応する顧客番号がエクセルファイルに見つかりませんでした"
                End If
            End If
            
            '進捗バーの更新
            returnValue = SysCmd(acSysCmdUpdateMeter, currentRow)
        End If
    Loop
    Close #fileNo
    endTime = Timer
    processTime = endTime - startTime
    
    returnValue = SysCmd(acSysCmdRemoveMeter)
    
    MsgBox totalRows & "件のSMS配信日を設定しました。" & vbNewLine & "所要時間:" & TimeSerial(0, 0, processTime)
    
    '後片付け
    xlBook.Close SaveChanges:=False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Exit Sub
    
ErrHandler:
    MsgBox "更新中にエラーが発生しました。ロールバックしました。エラー番号：" & Err.Number & " 内容：" & Err.Description
End Sub

Private Sub Form_Load()
    If Me!cmb送信テキスト.ListCount > 0 Then
        Me!cmb送信テキスト = Me!cmb送信テキスト.ItemData(0)
    End If
End Sub

Private Sub 抽出ボタン_Click()
    '宣言
    Dim sql As String
    Dim cSql As String 'チェック用のSQL
    Dim inSql As String
    Dim rs As DAO.Recordset
    Dim cRs As DAO.Recordset 'チェック用のレコードセット
    Dim lineNum As Long
    Dim loopCounter As Long
    Dim startTime, endTime, processTime As Variant
    Dim dataNum As Long
    Dim whereStr As String
    Dim upSql As String
    Dim dCnt2 As Integer
    Dim dCnt As Integer
    
    
    'TMPテーブルを削除する
    sql = "delete from TMP_SMS抽出用"
    CurrentDb.Execute sql
    Me!F_SUB_TMP_SMS抽出用.Requery
        
    If IsNull(Me.m残高From) Or Me.m残高From = "" Then
        MsgBox "残高Fromを設定してください"
        Exit Sub
    End If
    If IsNull(Me.m残高To) Or Me.m残高To = "" Then
        MsgBox "残高Toを設定してください"
        Exit Sub
    End If
    If (IsNull(Me.m最終取次日日数) Or Me.m最終取次日日数 = "") And chk最終取次日_未入力 = False Then
        MsgBox "最終取次日日数を設定してください"
        Exit Sub
    End If
    If (IsNull(Me.m入金約束日日数) Or Me.m入金約束日日数 = "") And chk入金約束日_未入力 = False Then
        MsgBox "入金約束日日数を設定してください"
        Exit Sub
    End If
    If IsNull(Me.m受託日From) Or Me.m受託日From = "" Then
        MsgBox "受託日Fromを設定してください"
        Exit Sub
    End If
    If IsNull(Me.m受託日To) Or Me.m受託日To = "" Then
        MsgBox "受託日Toを設定してください"
        Exit Sub
    End If
    
    If vbNo = MsgBox("SMS送信用のデータ抽出を行います。よろしいですか？", vbYesNo + vbInformation, "抽出実行") Then
        Exit Sub
    End If
    
    'where句を生成する
    whereStr = "[残高の合計] >= " & Me.m残高From & " And [残高の合計] <= " & Me.m残高To
    
    If chk不到達 Then
        whereStr = whereStr & " And [変更] = '不送達'"
    Else
        whereStr = whereStr & " And [変更] <> '不送達'"
    End If
    
    If chk最終取次日_未入力 Then
        whereStr = whereStr & " And [最終取次日] Is Null"
    Else
        whereStr = whereStr & " And [最終取次日] <= #" & Date - Val(Me.m最終取次日日数) & "#"
    End If
    
    If chk入金約束日_未入力 Then
        whereStr = whereStr & " And [入金約束日] Is Null"
    Else
        whereStr = whereStr & " And [入金約束日] <= #" & Date - Val(Me.m入金約束日日数) & "#"
    End If
    
    whereStr = whereStr & " And [受託日] >= #" & Me.m受託日From & "# And [受託日] <= #" & Me.m受託日To & "#"
        
    '件数を数える
    sql = "select count(*) as 件数 from Q_SMS抽出用 where " & whereStr
    Debug.Print sql
    Set rs = CurrentDb.OpenRecordset(sql)
    lineNum = rs![件数]
    
    'データを取得する
    sql = "select * from Q_SMS抽出用 where " & whereStr
    Set rs = CurrentDb.OpenRecordset(sql)
    
    loopCounter = 1
    dataNum = 0
    startTime = Timer
    SysCmd acSysCmdInitMeter, "TMP_SMS抽出用へ展開中…", lineNum
    
     Do Until rs.EOF
        SysCmd acSysCmdUpdateMeter, loopCounter
        '既存データに★を付ける
        upSql = "update TMP_SMS抽出用 set [氏名] = '★' & [氏名] where [顧客番号] = '" & rs![顧客番号] & "'"
        CurrentDb.Execute upSql, dbFailOnError
        
        '応対記録の対応依頼（対応依頼マスタ）が07（SMS送信）の場合は応対記録の完了にチェックが入っているデータを抽出（完了のデータがない）します
        cSql = "SELECT COUNT(*) as count from [応対記録] WHERE [顧客番号] = '" & rs![顧客番号] & "' AND [対応依頼] = '07' AND [完了] = FALSE"
        Set cRs = CurrentDb.OpenRecordset(cSql)
        
        '電話番号と氏名が一致する場合は◆を付ける
        upSql = "update TMP_SMS抽出用 set [氏名] = '◆' & [氏名] where [氏名] = '" & rs![氏名] & "' AND [電話１] = '" & rs![電話1] & "'"
        CurrentDb.Execute upSql, dbFailOnError
        
        'TMP_SMS抽出用に存在しないかを確認する→存在しない場合は投入する（顧客番号）
        dCnt = DCount("*", "TMP_SMS抽出用", "[顧客番号] = '" & rs![顧客番号] & "'")
        
        'TMP_SMS抽出用に存在しないかを確認する→存在しない場合は投入する（顧客番号）
        dCnt2 = DCount("*", "TMP_SMS抽出用", "[電話１] = '" & rs![電話1] & "'")
        
        If cRs![Count] = 0 Then
            If dCnt = 0 And dCnt2 = 0 Then
                dataNum = dataNum + 1
                inSql = "insert into TMP_SMS抽出用([顧客番号],[氏名],[残高の合計],[電話１],[電話２],[受託番号],[受託日],[最終取次日],[入金約束日],[変更],[外部委託ステータス])"
                inSql = inSql & " values( "
                inSql = inSql & "'" & rs![顧客番号] & "',"
                inSql = inSql & "'" & rs![氏名] & "',"
                inSql = inSql & "'" & rs![残高の合計] & "',"
                inSql = inSql & "'" & rs![電話1] & "',"
                If rs![電話２] = "" Then
                   inSql = inSql & "NULL,"
                Else
                    inSql = inSql & "'" & rs![電話２] & "',"
                End If
                inSql = inSql & "'" & rs![受託番号] & "',"
                inSql = inSql & "#" & Format(rs![受託日], "YYYY/MM/DD") & "#, "
                If rs![最終取次日] = "" Or IsNull(rs![最終取次日]) Then
                    inSql = inSql & "NULL,"
                Else
                    inSql = inSql & "#" & Format(rs![最終取次日], "YYYY/MM/DD") & "#, "
                End If
                If rs![入金約束日] = "" Or IsNull(rs![入金約束日]) Then
                    inSql = inSql & "NULL,"
                Else
                    inSql = inSql & "#" & Format(rs![入金約束日], "YYYY/MM/DD") & "#, "
                End If
                If rs![変更] = "" Then
                    inSql = inSql & "NULL,"
                Else
                    inSql = inSql & "'" & rs![変更] & "',"
                End If
                If rs![外部委託ステータス] = "" Then
                    inSql = inSql & "NULL,"
                Else
                    inSql = inSql & "'" & rs![外部委託ステータス] & "')"
                End If
                
                'Debug.Print inSql
                CurrentDb.Execute inSql
            End If
        Else
            MsgBox rs![顧客番号] & "のお客様は「応対依頼」が「SMS送信（07）」で「完了」が未完了のデータがあります"
        End If
        loopCounter = loopCounter + 1
        rs.MoveNext
    Loop
    Me!F_SUB_TMP_SMS抽出用.Requery
    endTime = Timer
    processTime = endTime - startTime
    SysCmd acSysCmdClearStatus 'ステータスバーの削除
    MsgBox dataNum & "件のデータをTMP_SMS抽出用データへ展開しました。" & vbNewLine & "所要時間:" & TimeSerial(0, 0, processTime)
    Set rs = Nothing
    Set cRs = Nothing
End Sub

Private Sub CSV出力_Click()
    Dim fileName As String
    Dim dlg As Object
    Dim boolResult As Boolean
    Dim strFiles As String
    Dim i As Long
    Dim folderPath As String
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    '保存先ダイアログの表示
    MsgBox "保存用のエクセルファイルの出力を行います"
    With dlg
        .Title = "SMS配信追跡用エクセルファイル出力"
        .ButtonName = "出力"
        .InitialFileName = "C:\" & "SMS一括配信_" & Format(Now, "yyyymmdd") & ".xlsx"
    End With
    boolResult = dlg.Show
    If boolResult Then
        For i = 1 To dlg.SelectedItems.Count
            strFiles = strFiles & dlg.SelectedItems(i)
        Next i
        fileName = strFiles
        folderPath = Left(fileName, InStrRev(fileName, "\"))
    Else
        'キャンセル
        Exit Sub
    End If
    'エクセルファイルの保存
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Q_TMP_SMS抽出用_顧客番号順", fileName, True
    
    MsgBox "エクセルの出力が完了しました｡"
    
    '配信用CSVファイル出力
    fileName = folderPath & "SMS一括配信_" & Format(Now, "yyyymmdd") & ".csv"
    If MsgBox("配信用CSVファイルの出力を行います。" & fileName, vbYesNo + vbQuestion) = vbYes Then
    Else
        'キャンセル
        Exit Sub
    End If
    Dim fileNo As Long
    Dim rs As Recordset
    Dim m携帯電話番号 As String
    Dim mメッセージ As String
    Dim mRc As VbMsgBoxResult
    fileNo = FreeFile()
    Me!F_SUB_TMP_SMS抽出用.Form.Recordset.MoveFirst
    Open fileName For Output As #fileNo
    Do Until Me!F_SUB_TMP_SMS抽出用.Form.Recordset.EOF
        Debug.Print Me!F_SUB_TMP_SMS抽出用.Form.Recordset![電話1]
        m携帯電話番号 = Me!F_SUB_TMP_SMS抽出用.Form.Recordset![電話1]
        Me!cmb送信テキスト.SetFocus
        mメッセージ = Me!cmb送信テキスト.Text
        Print #fileNo, Replace(m携帯電話番号, "-", "") & ",,0," & mメッセージ & ",,,"
        Me!F_SUB_TMP_SMS抽出用.Form.Recordset.MoveNext
    Loop
    Close fileNo
    MsgBox "CSVの出力が完了しました｡"

End Sub


