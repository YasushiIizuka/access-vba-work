' 貼り付け先: コンボが載っているフォームのフォームモジュール
'   ※スクリーンショット確認の結果、対象コンボは F_特定5講座 の下部
'     （チェック状況サブフォーム＝F_SUB_特定5講座_チェック状況 と思われる）にある。
'     デザインビューでコンボを選択し、どのフォームに属しているか確認のうえ、
'     そのフォームの「コードの表示」に貼り付けること（メインフォーム側ではない）
' 内容: 2つのコンボの排他制御
'   ・代引きレター返送状況 に値が入ったら 再注レター返送状況 をグレーアウト（逆も同じ）
'   ・値を空に戻したら相手側を再び編集可能にする
'
' ★★★ 貼り付け前の確認（重要）★★★
'   1. コントロール名の確認: 画面のラベルは「再注レター“の”返送状況」となっている。
'      デザインビューでコンボ本体を選択 → プロパティシート「名前」を確認し、
'      名前が「再注レターの返送状況」だった場合は、VBE の置換（Ctrl+H）で
'      このコード内の 再注レター返送状況 → 再注レターの返送状況 に全置換する。
'      「代引きレター返送状況」側も同様に名前を確認する
'   2. 貼り付け先のフォームモジュールに既に Form_Current や各コンボの
'      AfterUpdate がある場合は、Sub を丸ごと貼らず、中身の
'      UpdateLetterComboState 呼び出しだけを既存 Sub に追記する
'      （日付フィルタのコードはメインフォーム側なので衝突しない）
'   3. フォームモジュールの先頭に Option Compare Database / Option Explicit が
'      既にある場合、下の2行は重複するので貼らない
Option Compare Database
Option Explicit

' 2つのコンボの状態を現在値に合わせて更新する（共通処理）
Private Sub UpdateLetterComboState()
    Dim hasDaibiki As Boolean
    Dim hasSaichu As Boolean

    ' Null・空文字のどちらも「空」とみなす
    hasDaibiki = (Nz(Me.代引きレター返送状況.Value, "") <> "")
    hasSaichu = (Nz(Me.再注レター返送状況.Value, "") <> "")

    ' フォーカスが乗っているコントロールは無効化できない（エラー2164）ため、
    ' これから無効化する側にフォーカスがある場合だけ相手側へ退避する
    On Error Resume Next
    If hasDaibiki And Me.ActiveControl.Name = "再注レター返送状況" Then
        Me.代引きレター返送状況.SetFocus
    End If
    If hasSaichu And Me.ActiveControl.Name = "代引きレター返送状況" Then
        Me.再注レター返送状況.SetFocus
    End If
    On Error GoTo 0

    Me.再注レター返送状況.Enabled = Not hasDaibiki
    Me.代引きレター返送状況.Enabled = Not hasSaichu
End Sub

Private Sub 代引きレター返送状況_AfterUpdate()
    UpdateLetterComboState
End Sub

Private Sub 再注レター返送状況_AfterUpdate()
    UpdateLetterComboState
End Sub

' 連結フォームの場合: レコード移動のたびに、その行の値に合わせて状態を揃える
Private Sub Form_Current()
    UpdateLetterComboState
End Sub
