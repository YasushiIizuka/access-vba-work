' 貼り付け先: 標準モジュール（T_WORCS にリンクしている Access ならどこでも可。
'             例: WORCS_取込 に mod一括修正 として貼り付け）
Option Compare Database
Option Explicit

' =============================================================
' 【1回だけ実行する処理】T_WORCS の登録日・更新日から時刻を除去する
'
' 背景:
'   これまで 登録日・更新日 に Now()（日付＋時刻）を入れていたが、
'   今後は Date（日付のみ）に変更する。既存データに時刻付きが
'   残っていると新旧が混在するため、この処理で揃える。
'
' 実行方法:
'   VBE のイミディエイトウィンドウ（Ctrl+G）で
'       Fix時刻除去
'   と入力して Enter。実行前に件数と確認ダイアログが出る。
'
' 実行前に WORCS_データ のバックアップを取ること。
' 実行後はこのモジュールごと削除してよい。
' =============================================================

Public Sub Fix時刻除去()
    On Error GoTo Err_Handler

    Dim cnt登録 As Long
    Dim cnt更新 As Long
    cnt登録 = DCount("*", "T_WORCS", "[登録日] <> DateValue([登録日])")
    cnt更新 = DCount("*", "T_WORCS", "[更新日] <> DateValue([更新日])")

    If cnt登録 = 0 And cnt更新 = 0 Then
        MsgBox "時刻付きの登録日・更新日はありません。修正は不要です。", vbInformation
        Exit Sub
    End If

    If MsgBox("時刻付きのデータが以下の件数あります。" & vbCrLf & _
        "  登録日: " & Format(cnt登録, "#,##0") & " 件" & vbCrLf & _
        "  更新日: " & Format(cnt更新, "#,##0") & " 件" & vbCrLf & vbCrLf & _
        "時刻を除去して日付のみにします。" & vbCrLf & _
        "実行前に WORCS_データ のバックアップを取りましたか？" & vbCrLf & _
        "（続行すると元に戻せません）", _
        vbYesNo + vbQuestion + vbDefaultButton2, "登録日・更新日の時刻除去") = vbNo Then
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim fixed登録 As Long
    Dim fixed更新 As Long

    If cnt登録 > 0 Then
        db.Execute _
            "UPDATE T_WORCS SET 登録日 = DateValue([登録日]) " & _
            "WHERE [登録日] <> DateValue([登録日])", dbFailOnError
        fixed登録 = db.RecordsAffected
    End If

    If cnt更新 > 0 Then
        db.Execute _
            "UPDATE T_WORCS SET 更新日 = DateValue([更新日]) " & _
            "WHERE [更新日] <> DateValue([更新日])", dbFailOnError
        fixed更新 = db.RecordsAffected
    End If

    MsgBox "時刻を除去しました。" & vbCrLf & _
        "  登録日: " & Format(fixed登録, "#,##0") & " 件" & vbCrLf & _
        "  更新日: " & Format(fixed更新, "#,##0") & " 件", _
        vbInformation, "完了"
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生しました：" & Err.Number & ": " & Err.Description, vbCritical
End Sub
