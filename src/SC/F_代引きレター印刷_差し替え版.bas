' 貼り付け先: F_代引きレター印刷 のフォームモジュール【全文差し替え】
'
' ★★★ 客先での準備（定型文機能・2026-07-31追加）★★★
'   1. テーブル M_レター定型文 を作成:
'        ID      オートナンバー（主キー）
'        種別    短いテキスト
'        タイトル 短いテキスト
'        本文    長いテキスト
'        更新日  日付/時刻
'      （代引き・再注で共用。1回作れば両フォーム分に対応）
'   2. フォーム右側にコントロールを追加:
'        コンボボックス   名前: cbo定型文
'          列数: 2 ／ 列幅: 0cm;8cm ／ 連結列: 1 ／ 値集合ソース: 空欄（コードが設定）
'          更新後処理: [イベント プロシージャ]
'        テキストボックス 名前: txt定型文プレビュー
'          編集ロック: はい ／ スクロールバー: 垂直 ／ ラベルは「定型文」等
'        ボタン           名前: btn定型文セット   標題: ← 本文にセット
'          クリック時: [イベント プロシージャ]
'        ボタン           名前: btn定型文登録     標題: 現在の本文を登録
'          クリック時: [イベント プロシージャ]
'        ラベル           標題: 差し込み可能: {商品名称} {返送期限}
'
' 2026-07-31 の変更点:
'   ・定型文機能を追加（選択→プレビュー→[← 本文にセット]で差し替え、
'     [現在の本文を登録]で画面の本文を定型文として保存・同名は上書き確認）
'   ・印刷済みの表示切替を「レコードソースの差し替え」から「Filter の ON/OFF」に変更
'   ・Form_Load で毎回「チェックオフ＋未印刷のみ」に初期化
'   ・印刷ボタンの処理は従来のまま（置き換え文字: {商品名称} {返送期限}）
Option Compare Database
Option Explicit

'このフォームが使う定型文の種別
Private Const TEMPLATE_KIND As String = "代引き"

Private Sub Form_Load()
    '開いたときは必ず「未印刷のみ表示」に統一
    Me!chk印刷済みFLG = False
    ApplyPrintedFilter

    '定型文のタイトル一覧をセット
    SetupTemplateCombo
End Sub

Private Sub chk印刷済みFLG_AfterUpdate()
    ApplyPrintedFilter
End Sub

'チェックボックスの状態に合わせてサブフォームの絞り込みを適用する
Private Sub ApplyPrintedFilter()
    With Me![F_SUB_代引きレター印刷].Form
        If Me!chk印刷済みFLG = True Then
            '印刷済みも表示（絞り込み解除）
            .FilterOn = False
        Else
            '未印刷のみ表示
            .Filter = "[印刷済みFLG] = False"
            .FilterOn = True
        End If
    End With
End Sub

'―――――――――――――――――――――――――――――――
' 定型文機能
'―――――――――――――――――――――――――――――――

'定型文コンボのタイトル一覧を設定（ID は非表示の連結列）
Private Sub SetupTemplateCombo()
    Me!cbo定型文.RowSource = _
        "SELECT ID, タイトル FROM M_レター定型文 " & _
        "WHERE 種別 = '" & TEMPLATE_KIND & "' ORDER BY タイトル"
End Sub

'定型文を選んだらプレビューに表示
Private Sub cbo定型文_AfterUpdate()
    Me!txt定型文プレビュー.Value = _
        DLookup("本文", "M_レター定型文", "ID = " & Nz(Me!cbo定型文.Value, 0))
End Sub

'プレビューの定型文を本文に差し替える
Private Sub btn定型文セット_Click()
    Dim strNew As String

    strNew = Nz(Me!txt定型文プレビュー.Value, "")
    If strNew = "" Then
        MsgBox "差し替える定型文をドロップダウンから選んでください。", vbExclamation
        Exit Sub
    End If

    '手で編集した本文をうっかり消さないよう、内容が変わる場合は確認する
    If Nz(Me!txt本文元.Value, "") <> "" And Nz(Me!txt本文元.Value, "") <> strNew Then
        If MsgBox("現在の本文を、選択した定型文で上書きします。よろしいですか？", _
            vbYesNo + vbQuestion, "本文の差し替え") = vbNo Then Exit Sub
    End If

    Me!txt本文元.Value = strNew
    '本文はテーブルに連結しているため、その場で確定して保存する
    If Me.Dirty Then Me.Dirty = False
End Sub

'現在の本文を定型文として登録する（同名タイトルは上書き確認）
Private Sub btn定型文登録_Click()
    Dim strTitle As String
    Dim strBody As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    strBody = Nz(Me!txt本文元.Value, "")
    If strBody = "" Then
        MsgBox "本文が空のため登録できません。", vbExclamation
        Exit Sub
    End If

    strTitle = Trim$(InputBox("この本文を定型文として登録します。" & vbCrLf & _
        "タイトルを入力してください。（例: 標準、期限延長のお願い）", "定型文の登録"))
    If strTitle = "" Then Exit Sub 'キャンセルまたは未入力

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM M_レター定型文 WHERE 種別 = '" & TEMPLATE_KIND & "' " & _
        "AND タイトル = '" & Replace(strTitle, "'", "''") & "'", dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!種別 = TEMPLATE_KIND
        rs!タイトル = strTitle
        rs!本文 = strBody
        rs!更新日 = Date
        rs.Update
        MsgBox "定型文「" & strTitle & "」を登録しました。"
    Else
        If MsgBox("同じタイトルの定型文があります。上書きしますか？", _
            vbYesNo + vbQuestion, "定型文の登録") = vbYes Then
            rs.Edit
            rs!本文 = strBody
            rs!更新日 = Date
            rs.Update
            MsgBox "定型文「" & strTitle & "」を上書きしました。"
        End If
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    Me!cbo定型文.Requery
End Sub

'―――――――――――――――――――――――――――――――
' 印刷（従来のまま）
'―――――――――――――――――――――――――――――――

Private Sub btn代引きレター印刷_Click()
    Dim db As DAO.Database
    Dim rsSub As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim str本文元 As String
    Dim str本文 As String

    Set db = CurrentDb

    CurrentDb.Execute "delete from T_レター出力", dbFailOnError

    str本文元 = Nz(Me.txt本文元.Value, "")

    If str本文元 = "" Then
        MsgBox "本文が空です。", vbExclamation
        Exit Sub
    End If

    'サブフォームのレコードを取得
    Set rsSub = Me!F_SUB_代引きレター印刷.Form.RecordsetClone

    If rsSub.RecordCount = 0 Then
        MsgBox "印刷対象のデータがありません", vbExclamation
        Exit Sub
    End If

    Set rsOut = db.OpenRecordset("T_レター出力", dbOpenDynaset)

    rsSub.MoveFirst

    Do Until rsSub.EOF
        If Nz(rsSub!印刷済みFLG, False) = False Then
            str本文 = str本文元
            str本文 = Replace(str本文, "{商品名称}", Nz(rsSub!商品名称, ""))
            str本文 = Replace(str本文, "{返送期限}", Nz(rsSub!代引きレター返送期限, ""))
            rsOut.AddNew
            rsOut![No] = rsSub![No]
            rsOut![本文] = str本文
            rsOut![住所1] = Nz(rsSub![都道府県], "") & Nz(rsSub![市区郡], "")
            rsOut![住所2] = Nz(rsSub![町村番地], "") & Nz(rsSub![建物名], "")
            rsOut![氏名] = Nz(rsSub![顧客姓_漢字], "") & " " & Nz(rsSub![顧客名_漢字], "")
            rsOut![お申込商品] = "「" & Nz(rsSub![商品名称], "") & "」を代金引換払いで申し込みます。"
            rsOut![作成日時] = Now()
            rsOut.Update
        End If
        rsSub.MoveNext
    Loop
    DoCmd.OpenReport "R_代引きレター_特定5講座", acViewPreview
End Sub
