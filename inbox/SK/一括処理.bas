Option Compare Database


Private Sub btn一括処理開始_Click()
    Dim startVal As Single
    Dim endVal As Single
    Dim rap As Single
    Dim result As Single
    Dim i As Long
    
    startVal = Timer()
    
    txtLog.Value = ""
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　処理を開始します" & vbNewLine
    If chkユーキャン = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ユーキャン債権督促データ抽出を開始します" & vbNewLine
        Call ユーキャン債権督促データ抽出
    End If
    If chkSpot = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　spot督促抽出実行を開始します" & vbNewLine
        Call spot督促抽出実行
    End If
    If chkSpot金額 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　spot督促(金額範囲指定)抽出実行を開始します" & vbNewLine
        Call spot督促抽出実行_金額範囲指定
    End If
    If chkQSLSpot = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　QSL Spot督促抽出実行を開始します" & vbNewLine
        Call qslspot督促抽出実行
    End If
    
    If chkココチモ共通前処理 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理を開始します" & vbNewLine
        Call ココチモ共通前処理
    End If
    
    If chkココチモHY2 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ通常債権一括抽出HY-2を開始します" & vbNewLine
        Call ココチモ通常債権一括抽出HY2
    End If
    If chkココチモHY3 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ通常債権一括抽出HY-3を開始します" & vbNewLine
        Call ココチモ通常債権一括抽出HY3
    End If
    If chkココチモHY4 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ通常債権一括抽出HY-4を開始します" & vbNewLine
        Call ココチモ通常債権一括抽出HY4
    End If
    If chkココチモSpot = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ通常債権一括抽出Spotを開始します" & vbNewLine
        Call ココチモ通常債権一括抽出Spot
    End If
    If chkココチモQSL = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモQSLを開始します" & vbNewLine
        Call ココチモQSL
    End If
    If chkココチモQSL3 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモQSL3を開始します" & vbNewLine
        Call ココチモQSL3
    End If
    If chkココチモQSL4 = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモQSL4を開始します" & vbNewLine
        Call ココチモQSL4
    End If
    If chkココチモQSLSpot = True Then
        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモQSL Spotを開始します" & vbNewLine
        Call ココチモQSLSpot
    End If
    
    endVal = Timer()
    result = endVal - startVal
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　全体の処理時間は" & CDate(Format(result / 3600 / 24, "hh:nn:ss")) & "でした" & vbNewLine
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　処理が終了しました" & vbNewLine
    
End Sub
Private Sub ユーキャン債権督促データ抽出()

 Dim ret As Integer
 Dim stOutputDir As String

 ' マスタ情報の更新
 DoCmd.SetWarnings True
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業を行います" & vbNewLine
    DoCmd.SetWarnings False
    
    ' 変更情報の統一
    DoCmd.OpenQuery "☆更新☆変更情報の統一", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　変更情報の統一" & vbNewLine
    
    ' 進行状況テーブルの追加
    DoCmd.OpenQuery "進行状況テーブルの追加", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　進行状況テーブルの追加" & vbNewLine
    
    ' テーブル 入金約束データの削除
    DoCmd.OpenQuery "テーブル 入金約束データの削除", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　テーブル 入金約束データの削除" & vbNewLine
    
    ' 進行状況テーブルの入金約束日の削除
    DoCmd.OpenQuery "進行状況テーブルの入金約束日の削除", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　進行状況テーブルの入金約束日の削除" & vbNewLine
    
    ' テーブル入金約束データの追加
    DoCmd.OpenQuery "テーブル入金約束データの追加", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　テーブル入金約束データの追加" & vbNewLine
    
    ' 進行状況テーブルの入金約束日の追加
    DoCmd.OpenQuery "進行状況テーブルの入金約束日の追加", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　進行状況テーブルの入金約束日の追加" & vbNewLine
    
    ' 進行状況テーブル受託番号追加追加
    DoCmd.OpenQuery "進行状況テーブル受託番号追加", acViewNormal, acEdit
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　進行状況テーブル受託番号追加" & vbNewLine
    
    ' はい
    DoCmd.SetWarnings True
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業が完了しました" & vbNewLine
 
    ' ファイル出力
    stOutputDir = Application.CurrentProject.Path
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　督促データ抽出を行います" & vbNewLine
    DoCmd.SetWarnings False
    
    ' 前回実績の削除
    DoCmd.OpenQuery "①前回実績削除－hy2", acViewNormal, acEdit
    DoCmd.OpenQuery "②前回実績削除－hy2m", acViewNormal, acEdit
    DoCmd.OpenQuery "③前回実績削除－hy3", acViewNormal, acEdit
    DoCmd.OpenQuery "④前回実績削除－hy4", acViewNormal, acEdit
    DoCmd.OpenQuery "⑤前回実績削除－hy28", acViewNormal, acEdit
    DoCmd.OpenQuery "⑥前回実績削除－qsl", acViewNormal, acEdit
    DoCmd.OpenQuery "⑦前回実績削除－qsl3", acViewNormal, acEdit
    DoCmd.OpenQuery "⑧前回実績削除－qsl4", acViewNormal, acEdit
    
    ' 督促対象者抽出
    DoCmd.OpenQuery "①督促対象抽出－hy2", acViewNormal, acEdit
    DoCmd.OpenQuery "②督促対象抽出－hy2m", acViewNormal, acEdit
    DoCmd.OpenQuery "③督促対象抽出－hy3", acViewNormal, acEdit
    DoCmd.OpenQuery "④督促対象抽出－hy4", acViewNormal, acEdit
    DoCmd.OpenQuery "⑤督促対象抽出－hy28", acViewNormal, acEdit
    DoCmd.OpenQuery "⑥督促対象週出－qsl", acViewNormal, acEdit
    DoCmd.OpenQuery "⑦督促対象抽出－qsl3", acViewNormal, acEdit
    DoCmd.OpenQuery "⑧督促対象抽出－qsl4", acViewNormal, acEdit
    
    ' CSVファイルの出力・PDFファイルの出力
    cnt = DCount("顧客番号", "①福島印刷連携用－hy2")
    DoCmd.TransferText acExportDelim, , "①福島印刷連携用－hy2", stOutputDir & "\DATA\hy2_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "①前日日付用－hy2", acFormatPDF, stOutputDir & "\DATA\hy2_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "②福島印刷連携用－hy2m")
    DoCmd.TransferText acExportDelim, , "②福島印刷連携用－hy2m", stOutputDir & "\DATA\hy2m_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "②前日日付用－hy2m", acFormatPDF, stOutputDir & "\DATA\hy2m_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "③福島印刷連携用－hy3")
    DoCmd.TransferText acExportDelim, , "③福島印刷連携用－hy3", stOutputDir & "\DATA\hy3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "③前日日付用－hy3", acFormatPDF, stOutputDir & "\DATA\hy3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "④福島印刷連携用－hy4")
    DoCmd.TransferText acExportDelim, , "④福島印刷連携用－hy4", stOutputDir & "\DATA\hy4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "④前日日付用－hy4", acFormatPDF, stOutputDir & "\DATA\hy4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "⑤福島印刷連携用－hy28")
    DoCmd.TransferText acExportDelim, , "⑤福島印刷連携用－hy28", stOutputDir & "\DATA\hy28_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "⑤前日日付用－hy28", acFormatPDF, stOutputDir & "\DATA\hy28_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "⑥福島印刷連携用－qsl")
    DoCmd.TransferText acExportDelim, , "⑥福島印刷連携用－qsl", stOutputDir & "\DATA\qsl_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "⑥前日日付用－qsl", acFormatPDF, stOutputDir & "\DATA\qsl_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "⑦福島印刷連携用－qsl3")
    DoCmd.TransferText acExportDelim, , "⑦福島印刷連携用－qsl3", stOutputDir & "\DATA\qsl3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "⑦前日日付用－qsl3", acFormatPDF, stOutputDir & "\DATA\qsl3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    cnt = DCount("顧客番号", "⑧福島印刷連携用－qsl4")
    DoCmd.TransferText acExportDelim, , "⑧福島印刷連携用－qsl4", stOutputDir & "\DATA\qsl4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    DoCmd.OutputTo acOutputReport, "⑧前日日付用－qsl4", acFormatPDF, stOutputDir & "\DATA\qsl4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"

    ' 後処理
    DoCmd.SetWarnings True
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　督促データ抽出が完了しました" & vbNewLine
 
End Sub

Private Sub spot督促抽出実行()
' spot2抽出用

  Dim ret As Integer
  Dim stOutputDir As String

  stOutputDir = Application.CurrentProject.Path

  DoCmd.SetWarnings False

    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　SPOT2督促データ抽出を行います" & vbNewLine
    
    DoCmd.SetWarnings False
    
    ' 前回実績の削除
    DoCmd.OpenQuery "⑦前回実績削除－spot2", acViewNormal, acEdit
    
    ' 督促対象者抽出
    DoCmd.OpenQuery "⑦督促対象抽出－spot2", acViewNormal, acEdit
    
    ' CSVファイルの出力
    cnt = DCount("顧客番号", "⑦福島印刷連携用－spot2")
    DoCmd.TransferText acExportDelim, , "⑦福島印刷連携用－spot2", stOutputDir & "\DATA\spot2_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    ' PDFファイルの出力
    DoCmd.OutputTo acOutputReport, "⑦前日日付用－spot2", acFormatPDF, stOutputDir & "\DATA\spot2_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    ' 後処理
    DoCmd.SetWarnings True
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　SPOT2督促データ抽出が完了しました" & vbNewLine

End Sub
Private Sub spot督促抽出実行_金額範囲指定()
' SPOT3抽出用

  Dim ret As Integer
  Dim stOutputDir As String

  stOutputDir = Application.CurrentProject.Path

  DoCmd.SetWarnings False

    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　SPOT3督促データ抽出を行います" & vbNewLine
    
    DoCmd.SetWarnings False
    
    ' 前回実績の削除
    DoCmd.OpenQuery "⑧前回実績削除－spot3", acViewNormal, acEdit
    
    ' 督促対象者抽出
    DoCmd.OpenQuery "⑧督促対象抽出－spot3", acViewNormal, acEdit
    
    ' CSVファイルの出力
    cnt = DCount("顧客番号", "⑧福島印刷連携用－spot3")
    DoCmd.TransferText acExportDelim, , "⑧福島印刷連携用－spot3", stOutputDir & "\DATA\spot3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    
    ' PDFファイルの出力
    DoCmd.OutputTo acOutputReport, "⑧前日日付用－spot3", acFormatPDF, stOutputDir & "\DATA\spot3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    ' 後処理
    DoCmd.SetWarnings True
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　SPOT3督促データ抽出が完了しました" & vbNewLine
 
End Sub

Sub ココチモ共通前処理()
    DoCmd.SetWarnings False
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理を開始しました" & vbNewLine
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(☆更新☆変更情報の統一)" & vbNewLine
    CurrentDb.Execute "☆更新☆変更情報の統一"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(☆削除☆ココチモ顧客(注文番号))" & vbNewLine
    CurrentDb.Execute "☆削除☆ココチモ顧客(注文番号)"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(☆削除☆ココチモ変更(注文番号))" & vbNewLine
    CurrentDb.Execute "☆削除☆ココチモ変更(注文番号)"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(☆追加☆ココチモ変更(注文番号)テーブル)" & vbNewLine
    CurrentDb.Execute "☆追加☆ココチモ変更(注文番号)テーブル"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(☆追加☆ココチモ顧客(注文番号)テーブル)" & vbNewLine
    CurrentDb.Execute "☆追加☆ココチモ顧客(注文番号)テーブル"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(オーダー内訳_ココチモ_TMP展開)" & vbNewLine
    Call ココチモ_内訳対応
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(オーダー内訳_ココチモ_QSL_TMP展開)" & vbNewLine
    Call ココチモ_QSL_内訳対応
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理(レター発送用約束の更新作業を開始します)" & vbNewLine
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業(進行状況テーブルの追加)" & vbNewLine
    CurrentDb.Execute "進行状況テーブルの追加"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業(テーブル 入金約束データの削除)" & vbNewLine
    CurrentDb.Execute "テーブル 入金約束データの削除"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業(進行状況テーブルの入金約束日の削除)" & vbNewLine
    CurrentDb.Execute "進行状況テーブルの入金約束日の削除"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業(テーブル入金約束データの追加)" & vbNewLine
    CurrentDb.Execute "テーブル入金約束データの追加"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業(進行状況テーブルの入金約束日の追加)" & vbNewLine
    CurrentDb.Execute "進行状況テーブルの入金約束日の追加"
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　レター発送用約束の更新作業(進行状況テーブル受託番号追加)" & vbNewLine
    CurrentDb.Execute "進行状況テーブル受託番号追加"
            
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ共通前処理が終了しました" & vbNewLine
    
    DoCmd.SetWarnings True
End Sub
Sub ココチモ_QSL_内訳対応() '新通販対応2025-03-25
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsTMP As DAO.Recordset
    Dim strSQL As String
    Dim m商品名 As String
    Dim mご注文番号 As String
    Dim m商品名配列() As String
    Dim m配列初期化sts As Boolean
    Dim i As Integer
    Dim counter As Integer
    Dim rCount As Integer
    m配列初期化sts = False
    counter = 0
    
    'データベースの参照を設定
    Set db = CurrentDb()
    
    'TMPテーブルをクリアする
    db.Execute "DELETE * FROM [オーダー内訳_ココチモ_QSL_TMP]", dbFailOnError
    
    'SQLクエリの作成
'    strSQL = "SELECT * FROM [オーダー内訳 ｺｺﾁﾓ☆] WHERE [商品名] LIKE '*、他'"
    strSQL = "SELECT * FROM [オーダー内訳 ｺｺﾁﾓ☆_QSL]"
    
    'レコードセットの開始
    Set rs = db.OpenRecordset(strSQL)
    Set rsTMP = db.OpenRecordset("オーダー内訳_ココチモ_QSL_TMP", dbOpenDynaset)
    
    'レコードカウントをセットする
    rs.MoveLast
    rCount = rs.recordCount
    SysCmd acSysCmdInitMeter, "ココチモQSL内訳対応", rCount
    rs.MoveFirst
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　オーダー内訳_ココチモ_QSL_TMP展開(展開処理を開始します)" & vbNewLine

    'レコードセットをループして各レコードを出力
    Do Until rs.EOF
        '「、他」を削除
        If InStr(rs!商品名, "、他") > 0 Then
            m商品名 = Left(rs!商品名, Len(rs!商品名) - 2)
        Else
            m商品名 = rs!商品名
        End If
        'クエリで対応しているのでそのまま注文番号とする
        'mご注文番号 = Left(rs!ご注文番号, 9)
        mご注文番号 = rs!ご注文番号
    
        rsTMP.AddNew
        rsTMP!顧客番号 = rs!顧客番号
        rsTMP!ご注文番号 = mご注文番号
        rsTMP!商品名 = Trim(m商品名)
        rsTMP!メモ = rs!メモ
        rsTMP!ご契約日 = rs!ご契約日
        rsTMP!契約金額 = rs!契約金額
        rsTMP!受託番号 = rs!受託番号
        rsTMP!氏名 = rs!氏名
        rsTMP!顧客 = rs!顧客
        rsTMP!フリガナ = rs!フリガナ
        rsTMP!残高 = rs!残高
        rsTMP!クライアント = rs!クライアント
        rsTMP!完済 = rs!完済
        rsTMP!変更 = rs!変更
        rsTMP!レター = rs!レター
        rsTMP!入金額 = rs!入金額
        rsTMP.Update
        
        If Not IsNull(rs!メモ) Then
            m商品名配列 = メモ欄から商品名取得(rs!メモ)
        End If
        
        On Error Resume Next
            m配列初期化sts = (UBound(m商品名配列) >= 0)
        On Error GoTo 0
         
        Debug.Print m配列初期化sts
        
       If m配列初期化sts = True Then
            For i = LBound(m商品名配列) To UBound(m商品名配列)
                rsTMP.AddNew
                rsTMP!顧客番号 = rs!顧客番号
                rsTMP!ご注文番号 = mご注文番号
                rsTMP!商品名 = Trim(m商品名配列(i))
                rsTMP!メモ = rs!メモ
                rsTMP!ご契約日 = rs!ご契約日
                rsTMP!契約金額 = 0 '親の明細で合算されているので展開する場合はゼロ
                rsTMP!受託番号 = rs!受託番号
                rsTMP!氏名 = rs!氏名
                rsTMP!顧客 = rs!顧客
                rsTMP!フリガナ = rs!フリガナ
                rsTMP!残高 = 0
                rsTMP!クライアント = rs!クライアント
                rsTMP!完済 = rs!完済
                rsTMP!変更 = rs!変更
                rsTMP!レター = rs!レター
                rsTMP!入金額 = 0
                rsTMP.Update
            Next i
            m配列初期化sts = False
            Erase m商品名配列
        End If
        counter = counter + 1
        SysCmd acSysCmdUpdateMeter, counter
        rs.MoveNext
    Loop
    SysCmd acSysCmdClearStatus
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　オーダー内訳_ココチモ_QSL_TMP展開(展開処理が終了しました)" & vbNewLine
    'レコードセットとデータベースの参照を閉じる
    rs.Close
    Set rs = Nothing
    Set db = Nothing

End Sub
Sub ココチモ_内訳対応() '新通販対応2025-03-25
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsTMP As DAO.Recordset
    Dim strSQL As String
    Dim m商品名 As String
    Dim mご注文番号 As String
    Dim m商品名配列() As String
    Dim m配列初期化sts As Boolean
    Dim i As Integer
    Dim counter As Integer
    Dim rCount As Integer
    m配列初期化sts = False
    counter = 0
    
    'データベースの参照を設定
    Set db = CurrentDb()
    
    'TMPテーブルをクリアする
    db.Execute "DELETE * FROM [オーダー内訳_ココチモ_TMP]", dbFailOnError
    
    'SQLクエリの作成
'    strSQL = "SELECT * FROM [オーダー内訳 ｺｺﾁﾓ☆] WHERE [商品名] LIKE '*、他'"
    strSQL = "SELECT * FROM [オーダー内訳 ｺｺﾁﾓ☆]"
    
    'レコードセットの開始
    Set rs = db.OpenRecordset(strSQL)
    Set rsTMP = db.OpenRecordset("オーダー内訳_ココチモ_TMP", dbOpenDynaset)
    
    'レコードカウントをセットする
    rs.MoveLast
    rCount = rs.recordCount
    SysCmd acSysCmdInitMeter, "ココチモ内訳対応", rCount
    rs.MoveFirst
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　オーダー内訳_ココチモ_TMP展開(展開処理を開始します)" & vbNewLine
    
    
    'レコードセットをループして各レコードを出力
    Do Until rs.EOF
        '「、他」を削除
        If InStr(rs!商品名, "、他") > 0 Then
            m商品名 = Left(rs!商品名, Len(rs!商品名) - 2)
        Else
            m商品名 = rs!商品名
        End If
        'クエリで対応しているのでそのまま注文番号とする
        'mご注文番号 = Left(rs!ご注文番号, 9)
        mご注文番号 = rs!ご注文番号
    
        rsTMP.AddNew
        rsTMP!顧客番号 = rs!顧客番号
        rsTMP!ご注文番号 = mご注文番号
        rsTMP!商品名 = Trim(m商品名)
        rsTMP!メモ = rs!メモ
        rsTMP!ご契約日 = rs!ご契約日
        rsTMP!契約金額 = rs!契約金額
        rsTMP!受託番号 = rs!受託番号
        rsTMP!氏名 = rs!氏名
        rsTMP!顧客 = rs!顧客
        rsTMP!フリガナ = rs!フリガナ
        rsTMP!残高 = rs!残高
        rsTMP!クライアント = rs!クライアント
        rsTMP!完済 = rs!完済
        rsTMP!変更 = rs!変更
        rsTMP!レター = rs!レター
        rsTMP!入金額 = rs!入金額
        rsTMP.Update
        
        If Not IsNull(rs!メモ) Then
            m商品名配列 = メモ欄から商品名取得(rs!メモ)
        End If
        
            On Error Resume Next
            m配列初期化sts = (UBound(m商品名配列) >= 0)
        On Error GoTo 0
         
        Debug.Print m配列初期化sts
        
       If m配列初期化sts = True Then
            For i = LBound(m商品名配列) To UBound(m商品名配列)
                rsTMP.AddNew
                rsTMP!顧客番号 = rs!顧客番号
                rsTMP!ご注文番号 = mご注文番号
                rsTMP!商品名 = Trim(m商品名配列(i))
                rsTMP!メモ = rs!メモ
                rsTMP!ご契約日 = rs!ご契約日
                rsTMP!契約金額 = 0 '親の明細で合算されているので展開する場合はゼロ
                rsTMP!受託番号 = rs!受託番号
                rsTMP!氏名 = rs!氏名
                rsTMP!フリガナ = rs!フリガナ
                rsTMP!残高 = 0
                rsTMP!クライアント = rs!クライアント
                rsTMP!完済 = rs!完済
                rsTMP!変更 = rs!変更
                rsTMP!レター = rs!レター
                rsTMP!入金額 = 0
                rsTMP.Update
            Next i
            m配列初期化sts = False
            Erase m商品名配列
        End If
        counter = counter + 1
        SysCmd acSysCmdUpdateMeter, counter
        rs.MoveNext
    Loop
    SysCmd acSysCmdClearStatus
    
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　オーダー内訳_ココチモ_TMP展開(展開処理が終了しました)" & vbNewLine
    'レコードセットとデータベースの参照を閉じる
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
End Sub


Function メモ欄から商品名取得(str As String) As Variant
    Dim m開始Idx As Integer
    Dim m終了Idx As Integer
    Dim m商品名 As String
    Dim m商品名配列() As String

    '商品名のリストの開始と終了の位置を見つける
    m開始Idx = InStr(str, "{商品名 : [") + Len("{商品名 : [")
    m終了Idx = InStr(str, "] }")
    
    '商品名のリストを抽出
    If m開始Idx > 0 And m終了Idx > 0 Then
        m商品名 = Mid(str, m開始Idx, m終了Idx - m開始Idx)
        'カンマで分割して配列を作成
        m商品名配列 = Split(m商品名, ",")
    End If
    
    メモ欄から商品名取得 = m商品名配列

End Function
Private Sub ココチモ通常債権一括抽出HY2()

 Dim stOutputDir As String
 Dim ret As Integer

    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正HY_2作業を行います" & vbNewLine
    DoCmd.SetWarnings False
        
    stOutputDir = Application.CurrentProject.Path

    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリHY-2－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリHY-2－福島印刷連携用", stOutputDir & "\DATA\cohy2_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True

    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ HY-2", acFormatPDF, stOutputDir & "\DATA\cohy2_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
           
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正HY_2作業が完了しました" & vbNewLine
    DoCmd.SetWarnings True
    
End Sub
Private Sub ココチモ通常債権一括抽出HY3()

 Dim stOutputDir As String
 Dim ret As Integer

    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正HY_3作業を行います" & vbNewLine
    
    DoCmd.SetWarnings False
    stOutputDir = Application.CurrentProject.Path

    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリHY-3－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリHY-3－福島印刷連携用", stOutputDir & "\DATA\cohy3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True

    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ HY-3", acFormatPDF, stOutputDir & "\DATA\cohy3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
       
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正HY_3作業が完了しました" & vbNewLine
    DoCmd.SetWarnings True
End Sub

Private Sub ココチモ通常債権一括抽出HY4()

 Dim stOutputDir As String
 Dim ret As Integer
 
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正HY_4作業を行います" & vbNewLine
    
    DoCmd.SetWarnings False
    stOutputDir = Application.CurrentProject.Path

    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリHY-4－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリHY-4－福島印刷連携用", stOutputDir & "\DATA\cohy4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True

    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ HY-4", acFormatPDF, stOutputDir & "\DATA\cohy4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
       
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正HY_4作業が完了しました" & vbNewLine
    DoCmd.SetWarnings True

End Sub
Private Sub ココチモ通常債権一括抽出Spot()

 Dim stOutputDir As String
 Dim ret As Integer

    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正Spot作業を行います" & vbNewLine
    
    DoCmd.SetWarnings False
    
    stOutputDir = Application.CurrentProject.Path

    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリSpot－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリSpot－福島印刷連携用", stOutputDir & "\DATA\cospot_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True

    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ Spot", acFormatPDF, stOutputDir & "\DATA\cospot_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
       
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正Spot作業が完了しました" & vbNewLine
    DoCmd.SetWarnings True

End Sub
Private Sub ココチモQSL()
 Dim stOutputDir As String
 Dim ret As Integer
 Dim cnt As Integer
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正qsl作業を行います。QSL" & vbNewLine
    DoCmd.SetWarnings False
    stOutputDir = Application.CurrentProject.Path
    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリqsl－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリqsl－福島印刷連携用", stOutputDir & "\DATA\coqsl_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ qsl", acFormatPDF, stOutputDir & "\DATA\coqsl_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    Beep
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正sql作業が完了しました。QSL" & vbNewLine
    DoCmd.SetWarnings True
End Sub

Private Sub ココチモQSL3()
 Dim stOutputDir As String
 Dim ret As Integer
 Dim cnt As Integer
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正qsl3作業を行います。QSL3" & vbNewLine
    DoCmd.SetWarnings False
    stOutputDir = Application.CurrentProject.Path
    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリqsl3－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリqsl3－福島印刷連携用", stOutputDir & "\DATA\coqsl3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ qsl3", acFormatPDF, stOutputDir & "\DATA\coqsl3_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    Beep
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正qsl3作業が完了しました。QSL3" & vbNewLine
    DoCmd.SetWarnings True

End Sub

Private Sub ココチモQSL4()
 Dim stOutputDir As String
 Dim ret As Integer
 Dim cnt As Integer
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正qsl4作業を行います。QSL4" & vbNewLine
    DoCmd.SetWarnings False
    stOutputDir = Application.CurrentProject.Path
    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリqsl4－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリqsl4－福島印刷連携用", stOutputDir & "\DATA\coqsl4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ qsl4", acFormatPDF, stOutputDir & "\DATA\coqsl4_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    Beep
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正sql4作業が完了しました。QSL4" & vbNewLine
    DoCmd.SetWarnings True

End Sub
Private Sub ココチモQSLSpot()
 Dim stOutputDir As String
 Dim ret As Integer
 Dim cnt As Integer
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正qslspot作業を行います。QSLSpot" & vbNewLine
    DoCmd.SetWarnings False
    stOutputDir = Application.CurrentProject.Path
    'csvファイル出力
    cnt = DCount("顧客番号", "☆☆ココチモ督促顧客クエリqslspot－福島印刷連携用")
    DoCmd.TransferText acExportDelim, , "☆☆ココチモ督促顧客クエリqslspot－福島印刷連携用", stOutputDir & "\DATA\coqslspot_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    'PDFファイル出力
    DoCmd.OutputTo acOutputReport, "☆ココチモ qslspot", acFormatPDF, stOutputDir & "\DATA\coqslspot_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    Beep
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　ココチモ督促状発送修正sqlspot作業が完了しました。QSLSpot" & vbNewLine
    DoCmd.SetWarnings True
End Sub

Private Sub qslspot督促抽出実行()
' sqlspot抽出用

  Dim ret As Integer
  Dim stOutputDir As String
  Dim cnt As Integer

  stOutputDir = Application.CurrentProject.Path

  DoCmd.SetWarnings False

    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　qsl spot督促データ抽出を行います。QSLSpot" & vbNewLine
    
    DoCmd.SetWarnings False
    
    ' 前回実績の削除
    DoCmd.OpenQuery "⑨前回実績削除－qslspot", acViewNormal, acEdit
    ' 督促対象者抽出
    DoCmd.OpenQuery "⑨督促対象抽出－qslspot", acViewNormal, acEdit
    
    ' CSVファイルの出力
    cnt = DCount("顧客番号", "⑨福島印刷連携用－qslspot")
    DoCmd.TransferText acExportDelim, , "⑨福島印刷連携用－qslspot", stOutputDir & "\DATA\qsl_spot_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".csv", True
    ' PDFファイルの出力
    DoCmd.OutputTo acOutputReport, "⑨前日日付用－qslspot", acFormatPDF, stOutputDir & "\DATA\qsl_spot_" & cnt & "_" & Format$(Date, "yyyymmdd") & ".pdf"
    
    ' 後処理
    DoCmd.SetWarnings True
    Beep
    txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　qsl spot督促データ抽出が完了しました。QSLSpot" & vbNewLine
 
End Sub

