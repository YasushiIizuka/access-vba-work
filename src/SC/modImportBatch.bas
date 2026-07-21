' 貼り付け先: WORCS_取込 の標準モジュール modImportBatch（新規作成）
Option Explicit

' =============================================================
' 複数 Excel ファイルのテンポラリテーブルへの一括取込
'
' 処理の流れ:
'   1. ファイル選択ダイアログで Excel を複数選択（Ctrl+クリック）
'   2. テンポラリテーブルを最初に1回だけ空にする
'   3. 選択した全ファイルを順に TransferSpreadsheet でテンポラリへ追記
'   4. 結果（ファイル数＋合計行数）を表示して終了
'      （本番系への反映は既存のクエリ処理で行う）
'
' ★★★ 貼り付け後にここだけ直す ★★★
'   ・TEMP_TABLE 定数（テンポラリテーブル名）
'   ・ImportOneFile 内の TransferSpreadsheet の引数を既存の取込処理と同じにする
' =============================================================

' ★テンポラリテーブル名（既存の取込先テーブル名に変更）
Private Const TEMP_TABLE As String = "<<テンポラリテーブル名>>"

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
    Dim okFiles As Long      ' 取り込めたファイル数
    Dim totalRows As Long    ' 取り込めた合計行数
    Dim addedRows As Long    ' 1ファイル分の行数
    Dim filePath As String

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

ClearErr:
    MsgBox "テンポラリテーブルを空にできませんでした。取込を中止します。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, _
        vbExclamation, "取込中止"
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
