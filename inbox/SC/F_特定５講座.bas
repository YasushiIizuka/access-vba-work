Option Compare Database
Option Explicit

' ★サブフォームコントロールの名前（フォーム名ではなく、メインフォーム上の
'   サブフォームコントロールの「名前」プロパティを確認して合わせる）
Private Const SUB_LIST As String = "F_SUB_特定5講座"

' ★登録日フィールド名（両サブフォームのレコードソースにあるフィールド）
Private Const DATE_FIELD As String = "登録日"

' ★ドロップダウンの日付一覧を作る元（登録日を持つテーブルまたはクエリ）
Private Const DATE_SOURCE As String = "T_WORCS"

Private Sub cbo対象日_AfterUpdate()
    RequerySubForms
End Sub


Private Sub Form_Load()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strAllData As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT 商品コード, 商品名称 FROM M_特定講座 ORDER BY 商品コード")
    strAllData = ""
    Do Until rs.EOF
        strAllData = strAllData & rs![商品名称] & "（" & rs![商品コード] & "）　"
        rs.MoveNext
    Loop
    Me.lbl対象講座.Caption = Me.lbl対象講座.Caption & strAllData
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'ドロップダウンに「データがある登録日」を新しい順で一覧表示
    Me!cbo対象日.RowSource = _
        "SELECT DISTINCT [登録日] FROM [" & DATE_SOURCE & "] " & _
        "ORDER BY [登録日] DESC"

    '初期表示は今日
    Me!cbo対象日.Value = Date
    RequerySubForms
End Sub


'2つのサブフォームを再読込（絞り込みはクエリ側で行われる）
Private Sub RequerySubForms()
    On Error GoTo ErrHandler

    Me(SUB_LIST).Form.Requery
    Exit Sub

ErrHandler:
    MsgBox "サブフォームの再読込に失敗しました。" & vbCrLf & _
        "定数のサブフォームコントロール名を確認してください。" & vbCrLf & _
        "エラー内容: " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

