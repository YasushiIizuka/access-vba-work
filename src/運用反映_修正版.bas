' 貼り付け先: 運用反映フォームのコードビハインド（モジュール全体を差し替え）
' 変更点:
'   1. 重複時の1件ごとの MsgBox を廃止。黙ってスキップして件数を数え、
'      最後に「追加◯件・重複スキップ◯件」を1回だけ表示する
'   2. 重複の内訳（受注番号・商品コード）は先頭20件まで結果表示に含める
'   3. コメントアウトされていた On Error GoTo Err_Handler を有効化
Option Compare Database
Option Explicit

Private Sub btn反映_Click()
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rsSrc As DAO.Recordset
    Dim rsDest As DAO.Recordset
    Dim totalCount As Long
    Dim currentCount As Long
    Dim addedCount As Long
    Dim dupCount As Long
    Dim dupList As String
    Const DUP_LIST_MAX As Long = 20   '結果表示に載せる重複の最大件数

    Set db = CurrentDb

    'インポート対象のデータを開く
    Set rsSrc = db.OpenRecordset("Q_TMP_取込確認", dbOpenSnapshot)

    'データが空なら終了
    If rsSrc.EOF Then
        MsgBox "対象データが存在しません。", vbExclamation
        GoTo Exit_Handler
    End If

    '全件数を取得
    rsSrc.MoveLast
    totalCount = rsSrc.RecordCount
    rsSrc.MoveFirst

    '書き込み先のテーブルを開く
    Set rsDest = db.OpenRecordset("T_WORCS", dbOpenDynaset)

    '進捗バーの表示
    SysCmd acSysCmdInitMeter, "データをWORCSに取込中...", totalCount
    DoCmd.Hourglass True

    '取込処理
    Do Until rsSrc.EOF
        '重複チェック
        rsDest.FindFirst "[受注番号] = '" & rsSrc![受注番号] & "' AND [商品コード] = '" & rsSrc![商品コード] & "'"
        '重複がなければ追加
        If rsDest.NoMatch Then
            rsDest.AddNew
            rsDest![選択] = rsSrc![選択]
            rsDest![受注番号] = rsSrc![受注番号]
            rsDest![顧客番号] = rsSrc![顧客番号]
            rsDest![ショップ区分] = rsSrc![ショップ区分]
            rsDest![販売経路] = rsSrc![販売経路]
            rsDest![組織コード] = rsSrc![組織コード]
            rsDest![センター名] = rsSrc![センター名]
            rsDest![所属_会社] = rsSrc![所属_会社]
            rsDest![受注担当ID] = rsSrc![受注担当ID]
            rsDest![受注年月日] = rsSrc![受注年月日]
            rsDest![受注時刻] = rsSrc![受注時刻]
            rsDest![受注ステータス] = rsSrc![受注ステータス]
            rsDest![社コード] = rsSrc![社コード]
            rsDest![商品コード] = rsSrc![商品コード]
            rsDest![商品名称] = rsSrc![商品名称]
            rsDest![数量] = rsSrc![数量]
            rsDest![販売額_税込] = rsSrc![販売額_税込]
            rsDest![支払方法] = rsSrc![支払方法]
            rsDest![顧客姓_漢字] = rsSrc![顧客姓_漢字]
            rsDest![顧客名_漢字] = rsSrc![顧客名_漢字]
            rsDest![郵便番号] = rsSrc![郵便番号]
            rsDest![都道府県] = rsSrc![都道府県]
            rsDest![市区郡] = rsSrc![市区郡]
            rsDest![町村番地] = rsSrc![町村番地]
            rsDest![建物名] = rsSrc![建物名]
            rsDest![様方姓名] = rsSrc![様方姓名]
            rsDest![保護者姓名] = rsSrc![保護者姓名]
            rsDest![電話番号_自宅または携帯] = rsSrc![電話番号_自宅または携帯]
            rsDest![電話番号_携帯] = rsSrc![電話番号_携帯]
            rsDest![電話番号_その他] = rsSrc![電話番号_その他]
            rsDest![メールアドレス_PC] = rsSrc![メールアドレス_PC]
            rsDest![メールアドレス_モバイル] = rsSrc![メールアドレス_モバイル]
            rsDest![通販用メールアドレス] = rsSrc![通販用メールアドレス]
            rsDest![取引停止] = rsSrc![取引停止]
            rsDest![ブラックリスト該当] = rsSrc![ブラックリスト該当]
            rsDest![年齢] = rsSrc![年齢]
            rsDest![性別] = rsSrc![性別]
            rsDest![媒体コード] = rsSrc![媒体コード]
            rsDest![印刷キー] = rsSrc![印刷キー]
            rsDest![媒体名称] = rsSrc![媒体名称]
            If IsNull(rsSrc![入金総額]) Or rsSrc![入金総額] = "" Then
                rsDest![入金総額] = 0
            Else
                rsDest![入金総額] = rsSrc![入金総額]
            End If

            rsDest![登録日] = Now()
            rsDest![更新日] = Now()
            rsDest.Update
            addedCount = addedCount + 1
        Else
            '重複はスキップして数える（1件ごとのダイアログは出さない）
            dupCount = dupCount + 1
            If dupCount <= DUP_LIST_MAX Then
                dupList = dupList & vbCrLf & _
                    "  受注番号=" & rsSrc![受注番号] & _
                    " / 商品コード=" & rsSrc![商品コード]
            End If
        End If
        '進捗バーを更新
        currentCount = currentCount + 1
        If currentCount Mod 100 = 0 Then '100件ごとに更新
            SysCmd acSysCmdUpdateMeter, currentCount
            DoEvents '画面のフリーズの防止
        End If
        rsSrc.MoveNext
    Loop

    '終了処理
    SysCmd acSysCmdRemoveMeter
    DoCmd.Hourglass False

    '結果を1回だけまとめて表示
    Dim msg As String
    msg = "反映が完了しました。" & vbCrLf & vbCrLf & _
        "対象: " & Format(totalCount, "#,##0") & " 件" & vbCrLf & _
        "追加: " & Format(addedCount, "#,##0") & " 件" & vbCrLf & _
        "登録済みのためスキップ: " & Format(dupCount, "#,##0") & " 件"
    If dupCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "スキップした内訳"
        If dupCount > DUP_LIST_MAX Then
            msg = msg & "（先頭 " & DUP_LIST_MAX & " 件のみ表示）"
        End If
        msg = msg & ":" & dupList
    End If
    MsgBox msg, vbInformation, "運用DB反映"

Exit_Handler:
    On Error Resume Next
    If Not rsSrc Is Nothing Then rsSrc.Close
    If Not rsDest Is Nothing Then rsDest.Close
    SysCmd acSysCmdRemoveMeter
    DoCmd.Hourglass False
    Set rsSrc = Nothing
    Set rsDest = Nothing
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical
    Resume Exit_Handler

End Sub

Private Sub 詳細_Click()

End Sub
