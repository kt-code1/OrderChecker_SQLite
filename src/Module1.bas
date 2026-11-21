Attribute VB_Name = "Module1"
Option Explicit

Sub TestSqlite()

    Dim exePath As Variant
    Dim dbPath As Variant
    Dim resultPath As String
    Dim sql As String

    'sqlite3.exeを選択
    exePath = Application.GetOpenFilename("EXEファイル (*.exe),*.exe")
    If exePath = False Then Exit Sub

    'order_checker.dbを選択
    dbPath = Application.GetOpenFilename("DBファイル (*.db),*.db")
    If dbPath = False Then Exit Sub
    
    'SQLの結果を保存する一時ファイル
    resultPath = Environ$("TEMP") & "\sqlite_result.txt"

    '各シートの前回結果を削除
    ClearPreviousResult ThisWorkbook.Worksheets("照合結果"), 10, 7
    ClearPreviousResult ThisWorkbook.Worksheets("集計結果"), 2, 4

    'sql文
    sql = GetCheckSql

    '一時ファイルの後片付けのためにエラー検知
    On Error GoTo TestSqliteError

    'sqlite3.exeを実行
    If ExecuteSqlite(exePath, dbPath, resultPath, sql) = False Then
        GoTo CleanUp
    End If
    
    '結果出力
    If OutputResult(resultPath, ThisWorkbook.Worksheets("照合結果"), 10) = False Then
        GoTo CleanUp
    End If
    
    '集計SQL
    sql = GetSummarySql

    '集計SQLを実行
    If ExecuteSqlite(exePath, dbPath, resultPath, sql) = False Then
        GoTo CleanUp
    End If

    '集計結果を出力
    If OutputResult(resultPath, ThisWorkbook.Worksheets("集計結果"), 2) = False Then
        GoTo CleanUp
    End If
    
    '正常終了
    GoTo CleanUp

TestSqliteError:

    MsgBox Err.Description
    
    'エラー処理している状態を解除する
    Resume CleanUp
    
CleanUp:

    On Error GoTo CleanUpError

    '一時ファイル存在確認
    If Dir$(resultPath) <> "" Then
        '一時ファイルを削除
        Kill resultPath
    End If
    
    Exit Sub
    
CleanUpError:

    MsgBox "一時ファイル削除エラー：" & Err.Description

End Sub

Function GetCheckSql() As String

    GetCheckSql = _
        "SELECT " & _
            "orders.order_number, " & _
            "orders.product_code, " & _
            "products.product_name, " & _
            "orders.unit_price, " & _
            "products.standard_unit_price, " & _
            "CASE " & _
                "WHEN products.product_code IS NULL THEN 'NG' " & _
                "WHEN orders.unit_price <> products.standard_unit_price THEN 'NG' " & _
                "ELSE 'OK' " & _
            "END AS Result, " & _
            "CASE " & _
                "WHEN products.product_code IS NULL THEN '商品コードが商品マスタに存在しません' " & _
                "WHEN orders.unit_price <> products.standard_unit_price THEN '商品マスタとの単価不一致' " & _
                "ELSE '' " & _
            "END AS ErrorMessage " & _
        "FROM orders " & _
        "LEFT JOIN products " & _
            "ON orders.product_code = products.product_code;"

End Function

Sub ClearPreviousResult( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal lastDataCol As Long _
)

    Dim lastRow As Long

    '前回の出力結果がある場合のみ削除する
    '出力開始行より上の見出しなどを誤って削除しないため
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow >= startRow Then

        With ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastDataCol))
            .ClearContents
            .Borders.LineStyle = xlNone
        End With

    End If

End Sub

Function ExecuteSqlite( _
    ByVal exePath As Variant, _
    ByVal dbPath As Variant, _
    ByVal resultPath As String, _
    ByVal sql As String _
) As Boolean

    Dim cmd As String
    Dim proc As Object
    Dim limitTime As Date

    'コマンドを組み立てる
    '最終的なコマンドは以下の構成になる
    '【sqlit3.exeのパス + dbのパス +
    'sqlの実行結果を一時保存するファイルパス(sqlite3のonceというコマンドを使用) +
    'sql文】
    
    cmd = _
        """" & exePath & """" & _
        " " & _
        """" & dbPath & """" & _
        " " & _
        """" & ".once '" & resultPath & "'" & """" & _
        " " & _
        """" & sql & """"

    'sqlite3.exeを実行
    On Error GoTo ExecError
    Set proc = CreateObject("WScript.Shell").Exec(cmd)
    On Error GoTo 0

    '実行終了まで待つ
    limitTime = DateAdd("s", 10, Now)
    
    Do While proc.Status = 0

        If Now >= limitTime Then
            
            '外部プロセス(sqlite3.exe)終了
            On Error Resume Next 'Status確認後にsqlite3.exeが自然終了した場合の対策
            proc.Terminate
            On Error GoTo 0
            
            MsgBox "sqlite3.exeの実行がタイムアウトしました。"
            
            ExecuteSqlite = False
            
            Exit Function
        
        End If
        
        DoEvents
    
    Loop

    '実行・結果判定
    If proc.ExitCode <> 0 Then
        MsgBox "sqlite3.exeの実行エラー：" & proc.StdErr.ReadAll
        ExecuteSqlite = False
    Else
        ExecuteSqlite = True
    End If
    
    Exit Function
    
ExecError:
    
    MsgBox "VBAの実行時エラー：" & Err.Description
    ExecuteSqlite = False

End Function

Function OutputResult( _
    ByVal resultPath As String, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long _
) As Boolean
    
    Dim stm As Object
    Dim lines() As String
    Dim i As Long
    Dim fields() As String
    Dim j As Long
    Dim lastDataRow As Long
    Dim lastDataCol As Long
    
    'エラーが発生したら検知してFalseを返す
    '→TestSqliteで一時ファイルを削除する
    On Error GoTo OutputError

    'UTF-8としてファイルを読み取る
    Set stm = CreateObject("ADODB.Stream")

    'テキストとして読み込む
    stm.Type = 2
    'utf-8で読み込む
    stm.Charset = "utf-8"
    'ストリームを開く
    stm.Open
    
    'ストリームにテキストを読み込ませる
    stm.LoadFromFile resultPath

    'テキストを改行コードで分割
    lines = Split(Replace(stm.ReadText, vbCr, ""), vbLf)

    'セルに配置していく
    
    'データが0件の場合、罫線を引かないようにするため、
    '最終データ行の初期値を出力開始行の1行前にする。
    lastDataRow = startRow - 1
    
    For i = 0 To UBound(lines)

        If lines(i) <> "" Then
            
            '"|"で更に分割
            fields = Split(lines(i), "|")

            For j = 0 To UBound(fields)
                
                ws.Cells(i + startRow, j + 1).Value = fields(j)
                
            Next j
            
            'データを書き込んだ行を最終データ行として記録
            lastDataRow = i + startRow
            
            '最終列を記録
            lastDataCol = UBound(fields) + 1
            
        End If
            
    Next i
    
    '今回の出力範囲に罫線を設定
    If lastDataRow >= startRow Then

        ws.Range( _
            ws.Cells(startRow, 1), _
            ws.Cells(lastDataRow, lastDataCol) _
        ).Borders.LineStyle = xlContinuous

    End If

    stm.Close
    
    OutputResult = True
    
    Exit Function
    
OutputError:

    MsgBox "結果出力エラー：" & Err.Description

    'エラー発生時もストリームを閉じる
    '(stm.Close自体が失敗してもResume Nextで処理は最後まで進めるようにする)
    On Error Resume Next
    
    'stmがNothingではない
    If Not stm Is Nothing Then
    
        stm.Close
        
    End If
    
    On Error GoTo 0

    OutputResult = False

End Function

Function GetSummarySql() As String

    GetSummarySql = _
        "SELECT " & _
            "product_code, " & _
            "COUNT(*) AS order_count, " & _
            "SUM(quantity) AS total_quantity, " & _
            "SUM(quantity * unit_price) AS total_amount " & _
        "FROM orders " & _
        "GROUP BY product_code;"

End Function
