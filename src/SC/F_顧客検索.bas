' 貼り付け先: F_顧客検索（メインフォーム）のフォームモジュール【新規】
'
' 検索画面の構成:
'   F_顧客検索（非連結のメインフォーム）
'     ├ txt電話番号      … 検索条件（テキストボックス）
'     ├ txtメールアドレス … 検索条件（テキストボックス）
'     ├ btn検索          … 検索実行ボタン
'     ├ btnリセット      … 条件リセットボタン
'     └ F_SUB_顧客検索   … 検索結果一覧（サブフォームコントロール）
'   検索は Q_顧客検索 の抽出条件ではなく、コードで Filter を組み立てて
'   サブフォームに適用する方式（ナビゲーションフォームに埋め込んでも動く）。
'   一致方法は部分一致（入力値がどこかに含まれればヒット）、
'   両方入力したときは OR（どちらかに一致すれば表示）。
'
' ★★★ 客先での準備 ★★★
'   1. サブフォーム用フォームを作る:
'      ・新規フォーム → レコードソースに Q_顧客検索 を指定
'      ・「既定のビュー」をデータシートビュー（または帳票形式）にする
'      ・一覧に出したい列だけ配置し、F_SUB_顧客検索 という名前で保存
'   2. メインフォームを作る:
'      ・新規フォーム（レコードソースは空のまま＝非連結）、F_顧客検索 で保存
'      ・ヘッダー部にテキストボックス2つとボタン2つを配置（ウィザードはキャンセル）:
'          名前: txt電話番号        ラベル: 電話番号
'          名前: txtメールアドレス  ラベル: メールアドレス
'          名前: btn検索            標題: 検索
'          名前: btnリセット        標題: リセット
'      ・詳細部にサブフォームコントロールを配置し、
'          名前:               F_SUB_顧客検索
'          ソースオブジェクト: F_SUB_顧客検索
'      ・btn検索 の「既定」プロパティを「はい」にする（Enter キーで検索できる）
'   3. btn検索・btnリセット の「クリック時」に [イベント プロシージャ] を選択し、
'      このファイル全文をフォームモジュールへ貼り付ける
'   4. 電話番号・メールアドレスのフィールド名が下の定数と違う場合は、
'      定数の値を実際のフィールド名に直す
Option Compare Database
Option Explicit

' ★検索対象のフィールド名（Q_顧客検索 は T_WORCS と T_WORCS_Check の結合クエリで
'   同名フィールドがあるため、必ずテーブル名で修飾したまま使う。
'   実際のフィールド名が違う場合はここだけ直せばよい）
Private Const FLD_TEL As String = "[T_WORCS].[電話番号]"
Private Const FLD_MAIL As String = "[T_WORCS].[メールアドレス]"

' ★サブフォームコントロールの名前（フォーム名ではなく、メインフォーム上の
'   サブフォームコントロールの「名前」プロパティを確認して合わせる）
Private Const SUB_LIST As String = "F_SUB_顧客検索"


Private Sub btn検索_Click()
    On Error GoTo ErrHandler

    Dim strTel As String
    Dim strMail As String
    Dim strFilter As String

    '入力値を整える（電話番号はハイフン・空白を除去して数字だけで比較する）
    strTel = NormalizeTel(Nz(Me!txt電話番号.Value, ""))
    strMail = Trim$(Nz(Me!txtメールアドレス.Value, ""))

    '電話番号: データ側もハイフン・空白を除去してから部分一致
    If Len(strTel) > 0 Then
        strFilter = "Replace(Replace(Nz(" & FLD_TEL & ",''),'-',''),' ','') " & _
            "Like '*" & EscapeLike(strTel) & "*'"
    End If

    'メールアドレス: 部分一致（両方入力されていれば OR＝どちらかに一致すれば表示）
    If Len(strMail) > 0 Then
        If Len(strFilter) > 0 Then strFilter = strFilter & " OR "
        strFilter = strFilter & FLD_MAIL & " Like '*" & EscapeLike(strMail) & "*'"
    End If

    With Me(SUB_LIST).Form
        If Len(strFilter) > 0 Then
            .Filter = strFilter
            .FilterOn = True
        Else
            '条件が空なら全件表示に戻す
            .FilterOn = False
        End If
    End With
    Exit Sub

ErrHandler:
    MsgBox "検索でエラーが発生しました。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub btnリセット_Click()
    Me!txt電話番号.Value = Null
    Me!txtメールアドレス.Value = Null
    Me(SUB_LIST).Form.FilterOn = False
    Me!txt電話番号.SetFocus
End Sub


'電話番号の入力値からハイフン（半角・全角・長音）と空白を除去する
Private Function NormalizeTel(ByVal strValue As String) As String
    Dim strResult As String
    strResult = Trim$(strValue)
    strResult = Replace(strResult, "-", "")
    strResult = Replace(strResult, "－", "")
    strResult = Replace(strResult, "ー", "")
    strResult = Replace(strResult, " ", "")
    strResult = Replace(strResult, "　", "")
    NormalizeTel = strResult
End Function


'Like 条件に埋め込む文字列のエスケープ（ワイルドカードと引用符を無効化する）
Private Function EscapeLike(ByVal strValue As String) As String
    Dim strResult As String
    strResult = Replace(strValue, "[", "[[]")
    strResult = Replace(strResult, "*", "[*]")
    strResult = Replace(strResult, "?", "[?]")
    strResult = Replace(strResult, "#", "[#]")
    strResult = Replace(strResult, "'", "''")
    EscapeLike = strResult
End Function
