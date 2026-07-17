' 貼り付け先: WORCS_取込 の標準モジュール modImportBatch（新規作成）
Option Explicit

' =============================================================
' 複数 Excel ファイルの連続取込
'
' 処理の流れ:
'   1. ファイル選択ダイアログで Excel を複数選択（Ctrl+クリック）
'   2. バックアップを最初に1回だけ作成（modBackup の SafeCopyDatabase を使用）
'   3. ファイルごとに: テンポラリ削除 → TransferSpreadsheet → 反映クエリ
'   4. 成功したファイルは「取込済み」サブフォルダへ移動（二重取込の防止）
'   5. エラー時はそのファイルで停止（成功済み分は反映されたまま）
'
' ★★★ 貼り付け後にここだけ直す ★★★
'   ・下の定数3つ（テンポラリテーブル名・反映クエリ名・データDBのパス）
'   ・ImportOneFile 内の TransferSpreadsheet の引数を既存の取込処理と同じにする
' =============================================================

' ★テンポラリテーブル名（既存の取込先テーブル名に変更）
Private Const TEMP_TABLE As String = "<<テンポラリテーブル名>>"

' ★本番系へ追加する反映クエリ名（既存の追加クエリ名に変更）
Private Const APPEND_QUERY As String = "<<反映クエリ名>>"

' ★バックアップ対象（WORCS_データ）のフルパス（既存処理の srcDB と同じ値に変更）
Private Const SRC_DB As String = "<<WORCS_データのフルパス>>"

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

    ' --- バックアップ（最初に1回だけ） ---
    Dim backupFile As String
    backupFile = BuildBackupPath(SRC_DB)
    If Not SafeCopyDatabase(SRC_DB, backupFile) Then
        MsgBox "バックアップの作成に失敗したため、取込を中止します。", _
            vbExclamation, "取込中止"
        Exit Sub
    End If

    ' --- ファイルごとのループ ---
    Dim i As Long
    Dim okCount As Long
    Dim filePath As String

    For i = 1 To fd.SelectedItems.Count
        filePath = fd.SelectedItems(i)

        If Not ImportOneFile(filePath) Then
            MsgBox okCount & " 件は正常に取り込みました。" & vbCrLf & vbCrLf & _
                "以下のファイルでエラーが発生したため中断しました。" & vbCrLf & _
                filePath & vbCrLf & vbCrLf & _
                "エラー内容: " & mLastError & vbCrLf & vbCrLf & _
                "原因を確認後、残りのファイルだけ再実行してください。" & vbCrLf & _
                "（取り込めたファイルは「" & DONE_FOLDER & "」フォルダへ移動済みです）", _
                vbExclamation, "取込中断"
            Exit Sub
        End If

        MoveToDone filePath
        okCount = okCount + 1
    Next i

    MsgBox okCount & " 件のファイルをすべて取り込みました。", _
        vbInformation, "取込完了"
End Sub

' ===== 1ファイル分の処理: テンポラリ削除 → 取込 → 反映クエリ =====
Private Function ImportOneFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler

    ' (1) テンポラリテーブルを空にする
    CurrentDb.Execute "DELETE FROM [" & TEMP_TABLE & "]", dbFailOnError

    ' (2) Excel を取り込む
    ' ★引数（形式・HasFieldNames・Range）は既存の TransferSpreadsheet と
    '   必ず同じにすること。下は「1行目が見出し・シート全体」の例
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
        TEMP_TABLE, filePath, True

    ' (3) 反映クエリで本番系へ追加
    CurrentDb.Execute APPEND_QUERY, dbFailOnError

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

' ===== バックアップファイル名の組み立て（元DBと同じ場所に日時付きで作成） =====
Private Function BuildBackupPath(ByVal srcPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    BuildBackupPath = fso.GetParentFolderName(srcPath) & "\" & _
        fso.GetBaseName(srcPath) & "_bak_" & _
        Format(Now, "yyyymmdd_hhnnss") & "." & _
        fso.GetExtensionName(srcPath)
End Function
