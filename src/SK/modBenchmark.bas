' 貼り付け先: 標準モジュール modBenchmark
Option Compare Database
Option Explicit

'==============================================================================
' SK サーバー移行 性能検証用ベンチマークモジュール
'
' 前提構成: データ用Access(バックエンド)をネットワークフォルダーに置き、
'           各PCのUI用Access(フロントエンド)からリンクテーブルで参照する運用。
'
' 本番データには一切触りません。指定フォルダーに検証専用のバックエンド
' (bench_backend.mdb) を自動生成し、それに対して測定します。
'
' 使い方:
'   1) 空のAccessファイル(新規accdbでよい)を作り、VBEでこのモジュールを貼り付け
'   2) イミディエイトウィンドウ(Ctrl+G)で実行:
'        BenchRun "旧サーバー", "<<旧サーバーの検証用フォルダUNCパス>>"
'        BenchRun "新サーバー", "<<新サーバーの検証用フォルダUNCパス>>"
'   3) 結果は このAccessファイルと同じフォルダの bench_results.csv に追記される
'   4) 終わったら後片付け:
'        BenchCleanup "<<検証用フォルダUNCパス>>"
'
' 注意:
'   - 検証用フォルダは本番mdbと同じ共有上の「専用サブフォルダ」を推奨
'     (本番フォルダ直下に置かない)
'   - 業務時間帯とアイドル帯の両方で、各サーバー3回ずつ実行するのが目安
'==============================================================================

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

Private Const BACKEND_NAME As String = "bench_backend.mdb"
Private Const TBL As String = "BENCH_DATA"
Private Const LNK As String = "BENCH_LNK"
Private Const RESULT_CSV As String = "bench_results.csv"

Private mFreq As Currency

'------------------------------------------------------------------ 時間計測
Private Function TimerStart() As Currency
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    QueryPerformanceCounter TimerStart
End Function

Private Function TimerMs(ByVal t0 As Currency) As Double
    Dim t1 As Currency
    QueryPerformanceCounter t1
    TimerMs = (t1 - t0) / mFreq * 1000#
End Function

'------------------------------------------------------------------ メイン
' label      : 結果CSVに記録する識別名 (例: "旧サーバー")
' folderPath : 検証用フォルダのUNCパス (末尾の \ は不要)
Public Sub BenchRun(ByVal label As String, ByVal folderPath As String, _
                    Optional ByVal nBulkRows As Long = 2000, _
                    Optional ByVal nRowByRow As Long = 200, _
                    Optional ByVal nLookups As Long = 100, _
                    Optional ByVal nRuns As Long = 3)
    On Error GoTo EH

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "フォルダーにアクセスできません: " & folderPath, vbExclamation
        Exit Sub
    End If

    Dim backendPath As String
    backendPath = folderPath & "\" & BACKEND_NAME

    ' 環境情報を1行記録
    Dim bitness As String
    #If Win64 Then
        bitness = "64bit"
    #Else
        bitness = "32bit"
    #End If
    LogResult label, "環境情報", 0, 0, _
        "PC=" & Environ("COMPUTERNAME") & " User=" & Environ("USERNAME") & _
        " Access=" & SysCmd(acSysCmdAccessVer) & " " & bitness & " 対象=" & backendPath

    ' 検証用バックエンドを作り直し
    RecreateBackend backendPath

    Dim run As Long
    For run = 1 To nRuns
        RunOnce label, backendPath, run, nBulkRows, nRowByRow, nLookups
        DoEvents
    Next run

    Debug.Print "完了: " & label & " → 結果は " & ResultCsvPath()
    Exit Sub
EH:
    MsgBox "エラー " & Err.Number & ": " & Err.Description, vbCritical, "BenchRun"
End Sub

'------------------------------------------------------------------ 1回分の測定
Private Sub RunOnce(ByVal label As String, ByVal backendPath As String, _
                    ByVal run As Long, ByVal nBulkRows As Long, _
                    ByVal nRowByRow As Long, ByVal nLookups As Long)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim t0 As Currency
    Dim ms As Double
    Dim i As Long

    ' [1] DB接続 (共有モードでのオープン = ldbロックファイル作成を含む)
    t0 = TimerStart()
    Set db = DBEngine.OpenDatabase(backendPath, False, False)
    ms = TimerMs(t0)
    LogResult label, "1_DB接続(共有オープン)", run, ms, ""

    ' テーブルを空にしてから開始
    db.Execute "DELETE FROM " & TBL, dbFailOnError

    ' [2] 一括INSERT (トランザクションあり = 帯域寄りの指標)
    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0)
    t0 = TimerStart()
    ws.BeginTrans
    Set rs = db.OpenRecordset(TBL, dbOpenTable)
    For i = 1 To nBulkRows
        rs.AddNew
        rs!Val = i
        rs!Txt = "ベンチマークデータ" & Format(i, "000000")
        rs!Dt = Now
        rs.Update
    Next i
    rs.Close
    ws.CommitTrans
    ms = TimerMs(t0)
    LogResult label, "2_一括INSERT(トランザクション)", run, ms, _
        nBulkRows & "件 " & Format(nBulkRows / (ms / 1000#), "0") & "件/秒"

    ' [3] 1件ずつINSERT (トランザクションなし = 遅延(RTT)に最も敏感な指標)
    t0 = TimerStart()
    For i = 1 To nRowByRow
        db.Execute "INSERT INTO " & TBL & " (Val, Txt, Dt) VALUES (" & _
            (100000 + i) & ", '逐次挿入', Now())", dbFailOnError
    Next i
    ms = TimerMs(t0)
    LogResult label, "3_逐次INSERT(1件ずつ)", run, ms, _
        nRowByRow & "件 平均" & Format(ms / nRowByRow, "0.0") & "ms/件"

    ' [4] 全件スキャン (集計クエリ)
    t0 = TimerStart()
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C, SUM(Val) AS S FROM " & TBL, dbOpenSnapshot)
    ms = TimerMs(t0)
    LogResult label, "4_全件集計(COUNT/SUM)", run, ms, rs!C & "件"
    rs.Close

    ' [5] 全件読み取り (全行・全列をクライアントへ転送)
    t0 = TimerStart()
    Set rs = db.OpenRecordset("SELECT * FROM " & TBL, dbOpenSnapshot)
    Dim v As Variant
    Do Until rs.EOF
        v = rs!Val: v = rs!Txt: v = rs!Dt
        rs.MoveNext
    Loop
    ms = TimerMs(t0)
    LogResult label, "5_全件読み取り", run, ms, ""
    rs.Close

    ' [6] インデックス検索 (1件ずつの検索を繰り返す = 画面での検索操作に相当)
    Dim hit As Long
    Randomize
    t0 = TimerStart()
    For i = 1 To nLookups
        Set rs = db.OpenRecordset("SELECT * FROM " & TBL & _
            " WHERE Val = " & (Int(nBulkRows * Rnd) + 1), dbOpenSnapshot)
        If Not rs.EOF Then hit = hit + 1
        rs.Close
    Next i
    ms = TimerMs(t0)
    LogResult label, "6_インデックス検索", run, ms, _
        nLookups & "回 平均" & Format(ms / nLookups, "0.0") & "ms/回"

    ' [7] 一括UPDATE
    t0 = TimerStart()
    db.Execute "UPDATE " & TBL & " SET Val = Val + 1", dbFailOnError
    ms = TimerMs(t0)
    LogResult label, "7_一括UPDATE", run, ms, ""

    db.Close
    Set db = Nothing

    ' [8] リンクテーブル経由 (本番と同じ参照形態での体感を測る)
    RelinkTable backendPath
    t0 = TimerStart()
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & LNK & " WHERE Val = 500", dbOpenSnapshot)
    ms = TimerMs(t0)
    LogResult label, "8_リンクテーブル検索(初回)", run, ms, "リンク経由・接続確立込み"
    rs.Close
End Sub

'------------------------------------------------------------------ バックエンド生成
Private Sub RecreateBackend(ByVal backendPath As String)
    Dim db As DAO.Database
    If Dir(backendPath) <> "" Then Kill backendPath
    Set db = DBEngine.CreateDatabase(backendPath, dbLangJapanese, dbVersion40)
    db.Execute "CREATE TABLE " & TBL & " (" & _
        "ID COUNTER CONSTRAINT pk PRIMARY KEY, " & _
        "Val LONG, Txt TEXT(50), Dt DATETIME)", dbFailOnError
    db.Execute "CREATE INDEX ix_val ON " & TBL & " (Val)", dbFailOnError
    db.Close
End Sub

'------------------------------------------------------------------ リンク張り替え
Private Sub RelinkTable(ByVal backendPath As String)
    Dim cdb As DAO.Database
    Set cdb = CurrentDb
    On Error Resume Next
    cdb.TableDefs.Delete LNK
    On Error GoTo 0
    Dim td As DAO.TableDef
    Set td = cdb.CreateTableDef(LNK)
    td.Connect = ";DATABASE=" & backendPath
    td.SourceTableName = TBL
    cdb.TableDefs.Append td
End Sub

'------------------------------------------------------------------ 後片付け
Public Sub BenchCleanup(ByVal folderPath As String)
    On Error Resume Next
    CurrentDb.TableDefs.Delete LNK
    Kill folderPath & "\" & BACKEND_NAME
    Kill folderPath & "\" & Replace(BACKEND_NAME, ".mdb", ".ldb")
    On Error GoTo 0
    Debug.Print "後片付け完了: " & folderPath
End Sub

'------------------------------------------------------------------ 結果CSV
Private Function ResultCsvPath() As String
    ResultCsvPath = CurrentProject.Path & "\" & RESULT_CSV
End Function

Private Sub LogResult(ByVal label As String, ByVal testName As String, _
                      ByVal run As Long, ByVal ms As Double, ByVal detail As String)
    Dim f As Integer
    Dim needHeader As Boolean
    needHeader = (Dir(ResultCsvPath()) = "")
    f = FreeFile
    Open ResultCsvPath() For Append As #f
    If needHeader Then
        Print #f, "日時,サーバーラベル,テスト項目,実行回,ミリ秒,詳細"
    End If
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & "," & label & "," & _
              testName & "," & run & "," & Format(ms, "0.0") & "," & detail
    Close #f
End Sub
