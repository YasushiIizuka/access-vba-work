' 貼り付け先: フォームモジュール 一括処理（既存の一括処理画面のモジュールに追記）
'
' ■ 前提
'   - 標準モジュール mod一括SMS抽出 が貼り付け済みであること
'   - T_SMS配信抽出条件（ID=1）に条件が保存されていること
'   - T_SMS配信抽出条件 に「送信テキスト」列（短いテキスト or 長いテキスト）を
'     追加しておくこと（一括実行時は SMS配信画面のコンボが使えないため、
'     配信メッセージも条件テーブルから読む）
'
' ■ デザインビューでの準備
'   - 一括処理画面に チェックボックス chkSMS配信 を追加（ラベル例: SMS配信抽出）
'
' ■ 組み込み手順（2か所）
'   1) 下の Sub SMS配信一括処理 をこのままモジュール末尾に貼り付ける
'   2) btn一括処理開始_Click の中（chkココチモQSLSpot のブロックの後ろなど）に
'      次の4行を追加する:
'
'    If chkSMS配信 = True Then
'        txtLog.Value = txtLog.Value & Format$(Now(), "hh:mm:ss") & "　SMS配信データ抽出を開始します" & vbNewLine
'        Call SMS配信一括処理
'    End If
'
' ■ 出力先
'   他の処理と同じく カレントDBと同じ場所の \DATA フォルダに出力する
'   （SMS一括配信_yyyymmdd.xlsx / SMS一括配信_yyyymmdd.csv。同名は上書き）

Option Compare Database
Option Explicit

Private Sub SMS配信一括処理()
    Dim outDir As String
    Dim msgText As String

    outDir = Application.CurrentProject.Path & "\DATA\"
    msgText = Nz(DLookup("送信テキスト", "T_SMS配信抽出条件", "ID = 1"), "")

    ' 一括モードON: mod一括SMS抽出 内の MsgBox はすべて txtLog への追記になる
    gBatchMode = True
    Set g実行ログ = Me!txtLog

    If SMS抽出実行() Then
        SMSファイル出力 outDir, msgText
    End If

    ' 一括モードOFF（他の処理や画面単体利用に影響を残さない）
    gBatchMode = False
    Set g実行ログ = Nothing
End Sub
