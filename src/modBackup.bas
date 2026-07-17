' 貼り付け先: WORCS_取込 の標準モジュール modBackup（新規作成）
Option Explicit

' =============================================================
' FileCopy の代替バックアップ処理
'
' 背景:
'   VBA の FileCopy はコピー元を排他モードで開くため、
'   他の Access（WORCS_利用者など）が WORCS_データ を開いていると
'   エラー 70（書き込みできません）で失敗する。
'   FileSystemObject の CopyFile は共有モードで読むため、
'   使用中のファイルでもコピーできる。
'
' 使い方（既存の取込処理側の修正）:
'   変更前: FileCopy srcDB, backupFile
'   変更後:
'       If Not SafeCopyDatabase(srcDB, backupFile) Then
'           MsgBox "バックアップの作成に失敗したため、取込を中止します。", _
'               vbExclamation, "取込中止"
'           Exit Sub   ' ← 呼び出し元の構造に合わせて調整
'       End If
' =============================================================

' 使用中のファイルでもコピーできるバックアップ処理（検証付き）
' 戻り値: True = バックアップ作成・検証OK / False = 失敗
Public Function SafeCopyDatabase(ByVal srcPath As String, _
                                 ByVal backupPath As String) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(srcPath) Then
        SafeCopyDatabase = False
        Exit Function
    End If

    ' True = 既存の同名バックアップは上書き
    fso.CopyFile srcPath, backupPath, True

    ' コピー直後に、バックアップが Access として開けるか検証する
    SafeCopyDatabase = VerifyBackup(backupPath)
    Exit Function

ErrHandler:
    SafeCopyDatabase = False
End Function

' バックアップファイルが破損なく開けるかを確認する（共有・読み取り専用で開く）
Private Function VerifyBackup(ByVal dbPath As String) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(dbPath, False, True)
    db.Close
    Set db = Nothing

    VerifyBackup = True
    Exit Function

ErrHandler:
    VerifyBackup = False
End Function
