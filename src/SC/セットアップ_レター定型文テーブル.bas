' 貼り付け先: 標準モジュール modSetup（無ければ新規作成。実行後も残してよい）
' 実行場所: WORCS_データ（バックエンド）側で実行する。
'           実行後、フロント側で 外部データ→Access→リンクテーブルを作成 で
'           M_レター定型文 のリンクを張ること
' 使い方: 貼り付けたら Sub 内にカーソルを置いて F5（実行）
' 内容: 定型文テーブル M_レター定型文 を作成する（GUI でのテーブル作成は不要）
' 備考: DAO の参照設定が無いファイルでも動く（型・定数を使っていない）
Option Compare Database
Option Explicit

Public Sub Setup_レター定型文テーブル()
    'DAO の参照設定が無いファイル（データ側等）でも動くよう、
    'DAO の型・定数を使わずに書いてある
    Const DB_FAIL_ON_ERROR As Long = 128   'dbFailOnError と同値

    On Error GoTo ErrHandler

    '既にある場合は何もしない
    Dim tdf As Object
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = "M_レター定型文" Then
            MsgBox "M_レター定型文 は既に存在します。何もしませんでした。", vbInformation
            Exit Sub
        End If
    Next tdf

    CurrentDb.Execute _
        "CREATE TABLE M_レター定型文 (" & _
        " ID AUTOINCREMENT PRIMARY KEY," & _
        " 種別 TEXT(20)," & _
        " タイトル TEXT(100)," & _
        " 本文 LONGTEXT," & _
        " 更新日 DATETIME" & _
        ")", DB_FAIL_ON_ERROR

    '検索用インデックス（種別＋タイトル）
    CurrentDb.Execute _
        "CREATE INDEX idx種別タイトル ON M_レター定型文 (種別, タイトル)", DB_FAIL_ON_ERROR

    MsgBox "M_レター定型文 を作成しました。" & vbCrLf & _
           "（ナビゲーションウィンドウに出ない場合は F5 で最新表示）", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "作成に失敗しました。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
