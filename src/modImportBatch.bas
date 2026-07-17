' 貼り付け先: WORCS_取込 の標準モジュール modImportBatch（新規作成）
Option Explicit

' =============================================================
' 複数 Excel ファイルのテンポラリテーブルへの一括取込
'
' 処理の流れ:
'   1. ファイル選択ダイアログで Excel を複数選択（Ctrl+クリック）
'   2. テンポラリテーブルを最初に1回だけ空にする
'   3. 選択した全ファイルを順に TransferSpreadsheet でテンポラリへ追記
'   4. 成功したファイルは「取込済み」サブフォルダへ移動（二重取込の防止）
'   5. 結果を表示して終了（本番系への反映は既存のクエリ処理で行う）
'
' ★★★ 貼り付け後にここだけ直す ★★★
'   ・TEMP_TABLE 定数（テンポラリテーブル名）
'   ・ImportOneFile 内の TransferSpreadsheet の引数を既存の取込処理と同じにする
' =============================================================

' ★テンポラリテーブル名（既存の取込先テーブル名に変更）
Private Const TEMP_TABLE As String = "<<テンポラリテーブル名>>"

' 処理済みファイルの移動先サブフォルダ名（元ファイルと同じ場所に自動作成）
Private Const DONE_FOLDER As String = "取込済み"

' 直近のエラー内容（結果表示用）
Private mLastError As String

' ===== メイン: これをボタン等から呼ぶ =====
Public Sub ImportExcelBatch()
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' 3 = msoFileDialogFilePicker

    With fd
        .AllowMultiSelect = True
        .Title = "取り込む Excel ファイルを選択してください（複数選択可）"
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx; *.xlsm; *.xls"
        If .Show = 0 Then Exit Sub       ' キャンセル
    End With

    ' --- テンポラリテーブルを最初に1回だけ空にする ---
    On Error GoTo ClearErr
    CurrentDb.Execute "DELETE FROM [" & TEMP_TABLE & "]", dbFailOnError
    On Error GoTo 0

    ' --- 選択した全ファイルをテンポラリへ積み上げ ---
    Dim i As Long
    Dim okCount As Long
    Dim filePath As String

    For i = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(i)

        If Not ImportOneFile(filePath) Then
            MsgBox okCount & " 件はテンポラリに取り込み済みです。" & vbCrLf & vbCrLf & _
                "以下のファイルでエラーが発生したため中断しました。" & vbCrLf & _
                filePath & vbCrLf & vbCrLf & _
                "エラー内容: " & mLastError & vbCrLf & vbCrLf & _
                "取り込めた分はテンポラリに残っています。" & vbCrLf & _
                "このまま反映してから残りを再実行するか、" & vbCrLf & _
                "原因を直して最初からやり直してください。", _
                vbExclamation, "取込中断"
            Exit Sub
        End If

        MoveToDone filePath
        okCount = okCount + 1
    Next i

    MsgBox okCount & " 件のファイルをテンポラリに取り込みました。" & vbCrLf & _
        "続けて反映処理を実行してください。", _
        vbInformation, "取込完了"
    Exit Sub

ClearErr:
    MsgBox "テンポラリテーブルを空にできませんでした。取込を中止します。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, _
        vbExclamation, "取込中止"
End Sub

' ===== 1ファイル分の取込（テンポラリへ追記） =====
Private Function ImportOneFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler

    ' ★引数（形式・HasFieldNames・Range）は既存の TransferSpreadsheet と
    '   必ず同じにすること。下は「1行目が見出し・シート全体」の例
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
        TEMP_TABLE, filePath, True

    ImportOneFile = True
    Exit Function

ErrHandler:
    mLastError = Err.Number & ": " & Err.Description
    ImportOneFile = False
End Function

' ===== 取込済みファイルをサブフォルダへ移動 =====
Private Sub MoveToDone(ByVal filePath As String)
    On Error GoTo ErrHandler   ' 移動失敗は取込成功を妨げない（警告のみ）

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim doneDir As String
    doneDir = fso.GetParentFolderName(filePath) & "\" & DONE_FOLDER
    If Not fso.FolderExists(doneDir) Then fso.CreateFolder doneDir

    Dim destPath As String
    destPath = doneDir & "\" & fso.GetFileName(filePath)

    ' 同名ファイルが既にある場合は日時を付けて退避
    If fso.FileExists(destPath) Then
        destPath = doneDir & "\" & _
            fso.GetBaseName(filePath) & "_" & _
            Format(Now, "yyyymmdd_hhnnss") & "." & _
            fso.GetExtensionName(filePath)
    End If

    fso.MoveFile filePath, destPath
    Exit Sub

ErrHandler:
    MsgBox "取込は成功しましたが、ファイルの移動に失敗しました。" & vbCrLf & _
        filePath & vbCrLf & _
        "二重取込を防ぐため、手動で移動してください。", vbExclamation
End Sub
