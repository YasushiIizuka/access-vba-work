' 貼り付け先: 標準モジュール（T_WORCS にリンクしている Access ならどこでも可。
'             例: WORCS_取込 に mod一括修正 として貼り付け）
Option Compare Database
Option Explicit

' =============================================================
' 【1回だけ実行する処理】T_WORCS の登録日から時刻を除去する
'
' 背景:
'   これまで 登録日 に Now()（日付＋時刻）を入れていたが、
'   今後は Date（日付のみ）に変更する。既存データに時刻付きが
'   残っていると新旧が混在するため、この処理で揃える。
'
' 実行方法:
'   VBE のイミディエイトウィンドウ（Ctrl+G）で
'       Fix登録日時刻除去
'   と入力して Enter。実行前に件数と確認ダイアログが出る。
'
' 実行前に WORCS_データ のバックアップを取ること。
' 実行後はこのモジュールごと削除してよい。
' =============================================================

Public Sub Fix登録日時刻除去()
    On Error GoTo Err_Handler

    Dim targetCount As Long
    targetCount = DCount("*", "T_WORCS", "[登録日] <> DateValue([登録日])")

    If targetCount = 0 Then
        MsgBox "時刻付きの登録日はありません。修正は不要です。", vbInformation
        Exit Sub
    End If

    If MsgBox("時刻付きの登録日が " & Format(targetCount, "#,##0") & _
        " 件あります。時刻を除去して日付のみにします。" & vbCrLf & vbCrLf & _
        "実行前に WORCS_データ のバックアップを取りましたか？" & vbCrLf & _
        "（続行すると元に戻せません）", _
        vbYesNo + vbQuestion + vbDefaultButton2, "登録日の時刻除去") = vbNo Then
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute _
        "UPDATE T_WORCS SET 登録日 = DateValue([登録日]) " & _
        "WHERE [登録日] <> DateValue([登録日])", dbFailOnError

    MsgBox Format(db.RecordsAffected, "#,##0") & " 件の登録日を日付のみに修正しました。", _
        vbInformation, "完了"
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生しました：" & Err.Number & ": " & Err.Description, vbCritical
End Sub
