Option Compare Database
Option Explicit

Private Const TEMP_TABLE As String = "T_TMP_取込"
Private mLastError As String

Private Sub btnExcel取込_Click()
    Dim fd As FileDialog
    Dim filePath As String
    
    Dim basePath As String
    Dim srcDB As String
    Dim backupFolder As String
    Dim backupFile As String
    
    On Error GoTo ErrHandler
    
    '==============================
    'パス設定
    '==============================
    basePath = CurrentProject.Path
    srcDB = basePath & "\WORCS_データ.accdb"
    backupFolder = basePath & "\backup"
    
    '==============================
    'backupフォルダ作成（無ければ）
    '==============================
    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If
    
    '==============================
    'バックアップ作成
    '==============================
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    backupFile = backupFolder & "\WORCS_データ_" & Format(Now(), "yyyymmdd") & ".accdb"
    fso.CopyFile srcDB, backupFile, True
    
    '==============================
    'エクセルファイル選択
    '==============================
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "取り込むExcelファイルを選択してください(複数選択可)"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx; *.xls"
        .AllowMultiSelect = True
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "処理をキャンセルしました。", vbInformation
            Exit Sub
        End If
    End With
    
    DoCmd.SetWarnings False
    
    ' --- テンポラリテーブルを最初に1回だけ空にする ---
    CurrentDb.Execute "DELETE FROM [" & TEMP_TABLE & "]", dbFailOnError
    
    ' --- 選択した全ファイルをテンポラリへ積み上げ ---
    Dim i As Long
    Dim okFiles As Long      ' 取り込めたファイル数
    Dim totalRows As Long    ' 取り込めた合計行数
    Dim addedRows As Long    ' 1ファイル分の行数

    For i = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(i)

        addedRows = ImportOneFile(filePath)
        If addedRows < 0 Then
            MsgBox "ここまで " & okFiles & " ファイル（" & _
                Format(totalRows, "#,##0") & " 行）はテンポラリに取り込み済みです。" & _
                vbCrLf & vbCrLf & _
                "以下のファイルでエラーが発生したため中断しました。" & vbCrLf & _
                filePath & vbCrLf & vbCrLf & _
                "エラー内容: " & mLastError & vbCrLf & vbCrLf & _
                "取り込めた分はテンポラリに残っています。" & vbCrLf & _
                "このまま反映してから残りを再実行するか、" & vbCrLf & _
                "原因を直して最初からやり直してください。", _
                vbExclamation, "取込中断"
            Exit Sub
        End If

        okFiles = okFiles + 1
        totalRows = totalRows + addedRows
    Next i

    MsgBox okFiles & " ファイル（合計 " & Format(totalRows, "#,##0") & _
        " 行）をテンポラリに取り込みました。" & vbCrLf & _
        "続けて反映処理を実行してください。", _
        vbInformation, "取込完了"
    Exit Sub

ExitProc:
    Set fd = Nothing
    Exit Sub
    
ErrHandler:
    DoCmd.SetWarnings True
    MsgBox "エラーが発生しました。" & vbCrLf & Err.Number & " : " & Err.Description, vbInformation
    Resume ExitProc
    
End Sub

' ===== 1ファイル分の取込（テンポラリへ追記） =====
' 戻り値: 取り込んだ行数（エラー時は -1）
Private Function ImportOneFile(ByVal filePath As String) As Long
    On Error GoTo ErrHandler

    Dim beforeCount As Long
    beforeCount = DCount("*", TEMP_TABLE)

    ' ★引数（形式・HasFieldNames・Range）は既存の TransferSpreadsheet と
    '   必ず同じにすること。下は「1行目が見出し・シート全体」の例
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
        TEMP_TABLE, filePath, True

    ImportOneFile = DCount("*", TEMP_TABLE) - beforeCount
    Exit Function

ErrHandler:
    mLastError = Err.Number & ": " & Err.Description
    ImportOneFile = -1
End Function

