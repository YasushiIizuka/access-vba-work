' 貼り付け先: 標準モジュール mod一括SMS抽出
'
' ■ 役割
'   SMS配信画面の「抽出」「CSV出力」のロジックを画面から切り離した共通版。
'   一括処理画面からも既存画面からも同じ関数を呼べる。
'   - 抽出条件は画面ではなく T_SMS配信抽出条件（ID=1 の1レコード）から読む
'   - メッセージは gBatchMode = True のとき MsgBox ではなく実行ログへ流す
'
' ■ 一括処理画面からの呼び出し例（ボタンクリック内）
'     gBatchMode = True
'     Set g実行ログ = Me!txt実行ログ            ' 実行ログ用テキストボックス
'     If SMS抽出実行() Then
'         SMSファイル出力 "<<出力先フォルダ>>\", Me!cmb送信テキスト
'     End If
'     Me!F_SUB_TMP_SMS抽出用.Requery            ' サブフォームがある画面なら
'     gBatchMode = False
'
' ■ 元コード（SMS配信画面）からの変更点
'   - MsgBox をすべて LogMsg に置換（一括時はログ、単体時は従来どおりダイアログ）
'   - ループ内の「応対依頼07未完了」ダイアログはログ1行＋末尾サマリーに変更
'   - 抽出条件を画面コントロールから T_SMS配信抽出条件 に変更
'   - CSV出力のサブフォーム依存をやめ、Q_TMP_SMS抽出用_顧客番号順 から直接出力
'   - Option Explicit を追加（元コードの totalRow/totalRows の書き間違いバグ対策）
'   - 外部委託ステータスが空のとき INSERT 文が閉じないバグを修正
'   - 電話２・変更・外部委託ステータスは「Null または空文字 → NULL」に統一
'   - 電話1／電話１ の全角半角は元コードのまま（フィールド名に合わせてある）
'
' ■ 注意
'   - 進捗バー（SysCmd）は一括時もステータスバーに表示される（ダイアログではないのでそのまま）
'   - 出力ファイルは同名があると上書きされる（元コードと同じ動き）

Option Compare Database
Option Explicit

' ===== 一括処理用の状態 =====
Public gBatchMode As Boolean    ' True: MsgBox を出さず実行ログへ
Public g実行ログ As Object      ' 実行ログ用テキストボックス（一括処理画面がセットする）

' メッセージ出力の共通口
Public Sub LogMsg(ByVal msg As String)
    If gBatchMode Then
        If Not g実行ログ Is Nothing Then
            g実行ログ.Value = g実行ログ.Value & Format$(Now, "hh:nn:ss") & " " & msg & vbCrLf
        End If
        Debug.Print Format$(Now, "hh:nn:ss") & " " & msg
    Else
        MsgBox msg
    End If
End Sub

' =========================================================
' 抽出: T_SMS配信抽出条件(ID=1) の条件で Q_SMS抽出用 → TMP_SMS抽出用 へ展開
' 戻り値: True=正常終了
' =========================================================
Public Function SMS抽出実行() As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim cond As DAO.Recordset
    Dim rs As DAO.Recordset
    Dim cRs As DAO.Recordset
    Dim sql As String
    Dim cSql As String
    Dim inSql As String
    Dim upSql As String
    Dim whereStr As String
    Dim lineNum As Long
    Dim loopCounter As Long
    Dim dataNum As Long
    Dim skipNum As Long
    Dim dCnt As Long
    Dim dCnt2 As Long
    Dim startTime As Variant, endTime As Variant, processTime As Variant

    Set db = CurrentDb

    ' --- 抽出条件の読み込み ---
    Set cond = db.OpenRecordset("SELECT * FROM T_SMS配信抽出条件 WHERE ID = 1")
    If cond.EOF Then
        LogMsg "抽出条件（T_SMS配信抽出条件 ID=1）が見つかりません"
        Exit Function
    End If

    ' --- 入力チェック（元画面と同じ必須条件） ---
    If IsNull(cond![残高From]) Then
        LogMsg "残高Fromを設定してください"
        Exit Function
    End If
    If IsNull(cond![残高To]) Then
        LogMsg "残高Toを設定してください"
        Exit Function
    End If
    If IsNull(cond![最終取次日からの日数]) And cond![最終取次日が未入力] = False Then
        LogMsg "最終取次日日数を設定してください"
        Exit Function
    End If
    If IsNull(cond![入金約束日からの日数]) And cond![入金約束日が未入力] = False Then
        LogMsg "入金約束日日数を設定してください"
        Exit Function
    End If
    If IsNull(cond![受託日From]) Then
        LogMsg "受託日Fromを設定してください"
        Exit Function
    End If
    If IsNull(cond![受託日To]) Then
        LogMsg "受託日Toを設定してください"
        Exit Function
    End If

    ' --- 確認（画面単体のときのみ。一括時は確認なしで実行） ---
    If Not gBatchMode Then
        If vbNo = MsgBox("SMS送信用のデータ抽出を行います。よろしいですか？", _
                         vbYesNo + vbInformation, "抽出実行") Then
            Exit Function
        End If
    End If

    ' --- TMPテーブルを空にする ---
    db.Execute "delete from TMP_SMS抽出用", dbFailOnError

    ' --- where句を生成する ---
    whereStr = "[残高の合計] >= " & cond![残高From] & " And [残高の合計] <= " & cond![残高To]

    If cond![不到達] Then
        whereStr = whereStr & " And [変更] = '不送達'"
    Else
        whereStr = whereStr & " And [変更] <> '不送達'"
    End If

    If cond![最終取次日が未入力] Then
        whereStr = whereStr & " And [最終取次日] Is Null"
    Else
        whereStr = whereStr & " And [最終取次日] <= #" & _
                   Format$(Date - cond![最終取次日からの日数], "yyyy/mm/dd") & "#"
    End If

    If cond![入金約束日が未入力] Then
        whereStr = whereStr & " And [入金約束日] Is Null"
    Else
        whereStr = whereStr & " And [入金約束日] <= #" & _
                   Format$(Date - cond![入金約束日からの日数], "yyyy/mm/dd") & "#"
    End If

    whereStr = whereStr & " And [受託日] >= #" & Format$(cond![受託日From], "yyyy/mm/dd") & "#" & _
               " And [受託日] <= #" & Format$(cond![受託日To], "yyyy/mm/dd") & "#"

    ' --- 件数を数える ---
    sql = "select count(*) as 件数 from Q_SMS抽出用 where " & whereStr
    Debug.Print sql
    Set rs = db.OpenRecordset(sql)
    lineNum = rs![件数]

    ' --- データを取得して TMP へ展開 ---
    sql = "select * from Q_SMS抽出用 where " & whereStr
    Set rs = db.OpenRecordset(sql)

    loopCounter = 1
    dataNum = 0
    skipNum = 0
    startTime = Timer
    SysCmd acSysCmdInitMeter, "TMP_SMS抽出用へ展開中…", lineNum

    Do Until rs.EOF
        SysCmd acSysCmdUpdateMeter, loopCounter

        ' 既存データ（同一顧客番号）に★を付ける
        upSql = "update TMP_SMS抽出用 set [氏名] = '★' & [氏名] " & _
                "where [顧客番号] = '" & rs![顧客番号] & "'"
        db.Execute upSql, dbFailOnError

        ' 応対記録の対応依頼が07（SMS送信）で完了が未完了のデータがないか確認
        cSql = "SELECT COUNT(*) as count from [応対記録] WHERE [顧客番号] = '" & rs![顧客番号] & "'" & _
               " AND [対応依頼] = '07' AND [完了] = FALSE"
        Set cRs = db.OpenRecordset(cSql)

        ' 電話番号と氏名が一致する場合は◆を付ける
        upSql = "update TMP_SMS抽出用 set [氏名] = '◆' & [氏名] " & _
                "where [氏名] = '" & rs![氏名] & "' AND [電話１] = '" & rs![電話1] & "'"
        db.Execute upSql, dbFailOnError

        ' TMP_SMS抽出用に存在しない場合のみ投入する（顧客番号・電話番号で重複確認）
        dCnt = DCount("*", "TMP_SMS抽出用", "[顧客番号] = '" & rs![顧客番号] & "'")
        dCnt2 = DCount("*", "TMP_SMS抽出用", "[電話１] = '" & rs![電話1] & "'")

        If cRs![Count] = 0 Then
            If dCnt = 0 And dCnt2 = 0 Then
                dataNum = dataNum + 1
                inSql = "insert into TMP_SMS抽出用([顧客番号],[氏名],[残高の合計],[電話１],[電話２]," & _
                        "[受託番号],[受託日],[最終取次日],[入金約束日],[変更],[外部委託ステータス])"
                inSql = inSql & " values( "
                inSql = inSql & "'" & rs![顧客番号] & "',"
                inSql = inSql & "'" & rs![氏名] & "',"
                inSql = inSql & "'" & rs![残高の合計] & "',"
                inSql = inSql & "'" & rs![電話1] & "',"
                inSql = inSql & NullOrText(rs![電話２]) & ","
                inSql = inSql & "'" & rs![受託番号] & "',"
                inSql = inSql & "#" & Format$(rs![受託日], "yyyy/mm/dd") & "#,"
                inSql = inSql & NullOrDate(rs![最終取次日]) & ","
                inSql = inSql & NullOrDate(rs![入金約束日]) & ","
                inSql = inSql & NullOrText(rs![変更]) & ","
                inSql = inSql & NullOrText(rs![外部委託ステータス]) & ")"
                db.Execute inSql, dbFailOnError
            End If
        Else
            ' 元コードはここで1件ずつ MsgBox → ログ1行に変更（一括時にダイアログ連発しない）
            skipNum = skipNum + 1
            LogMsg rs![顧客番号] & "のお客様は「応対依頼」が「SMS送信（07）」で「完了」が未完了のデータがあります（スキップ）"
        End If

        loopCounter = loopCounter + 1
        rs.MoveNext
    Loop

    endTime = Timer
    processTime = endTime - startTime
    SysCmd acSysCmdClearStatus

    LogMsg dataNum & "件のデータをTMP_SMS抽出用データへ展開しました。" & _
           IIf(skipNum > 0, "（応対記録未完了によるスキップ " & skipNum & "件）", "") & _
           " 所要時間:" & TimeSerial(0, 0, processTime)

    Set rs = Nothing
    Set cRs = Nothing
    Set cond = Nothing
    SMS抽出実行 = True
    Exit Function

ErrHandler:
    SysCmd acSysCmdClearStatus
    LogMsg "抽出中にエラーが発生しました。エラー番号：" & Err.Number & " 内容：" & Err.Description
End Function

' =========================================================
' ファイル出力: TMP_SMS抽出用 の内容をエクセル＋配信用CSVに出力
'   folderPath  : 出力先フォルダ（末尾 \ あり・なしどちらでも可）
'   messageText : 配信メッセージ（CSVの4列目に入る送信テキスト）
' 戻り値: True=正常終了
' =========================================================
Public Function SMSファイル出力(ByVal folderPath As String, ByVal messageText As String) As Boolean
    On Error GoTo ErrHandler

    Dim fileNameXlsx As String
    Dim fileNameCsv As String
    Dim fileNo As Integer
    Dim rs As DAO.Recordset
    Dim m携帯電話番号 As String
    Dim outCnt As Long

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    If Dir(folderPath, vbDirectory) = "" Then
        LogMsg "出力先フォルダが見つかりません: " & folderPath
        Exit Function
    End If
    If Trim$(messageText) = "" Then
        LogMsg "送信テキストが設定されていません"
        Exit Function
    End If
    If DCount("*", "TMP_SMS抽出用") = 0 Then
        LogMsg "TMP_SMS抽出用にデータがありません。先に抽出を実行してください"
        Exit Function
    End If

    fileNameXlsx = folderPath & "SMS一括配信_" & Format$(Date, "yyyymmdd") & ".xlsx"
    fileNameCsv = folderPath & "SMS一括配信_" & Format$(Date, "yyyymmdd") & ".csv"

    ' --- 追跡用エクセルの出力 ---
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
        "Q_TMP_SMS抽出用_顧客番号順", fileNameXlsx, True
    LogMsg "エクセルの出力が完了しました: " & fileNameXlsx

    ' --- 配信用CSVの出力（元コードと同じ行形式） ---
    fileNo = FreeFile
    Open fileNameCsv For Output As #fileNo
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Q_TMP_SMS抽出用_顧客番号順")
    outCnt = 0
    Do Until rs.EOF
        m携帯電話番号 = Nz(rs![電話1], "")
        If m携帯電話番号 <> "" Then
            Print #fileNo, Replace(m携帯電話番号, "-", "") & ",,0," & messageText & ",,,"
            outCnt = outCnt + 1
        End If
        rs.MoveNext
    Loop
    Close #fileNo
    Set rs = Nothing

    LogMsg "CSVの出力が完了しました（" & outCnt & "件）: " & fileNameCsv
    SMSファイル出力 = True
    Exit Function

ErrHandler:
    If fileNo <> 0 Then Close #fileNo
    LogMsg "ファイル出力中にエラーが発生しました。エラー番号：" & Err.Number & " 内容：" & Err.Description
End Function

' ===== SQL組み立て用の小道具 =====
' Null/空文字 → NULL、それ以外 → '値'
Private Function NullOrText(ByVal v As Variant) As String
    If IsNull(v) Then
        NullOrText = "NULL"
    ElseIf Trim$(v & "") = "" Then
        NullOrText = "NULL"
    Else
        NullOrText = "'" & v & "'"
    End If
End Function

' Null/空文字 → NULL、それ以外 → #日付#
Private Function NullOrDate(ByVal v As Variant) As String
    If IsNull(v) Then
        NullOrDate = "NULL"
    ElseIf Trim$(v & "") = "" Then
        NullOrDate = "NULL"
    Else
        NullOrDate = "#" & Format$(v, "yyyy/mm/dd") & "#"
    End If
End Function
