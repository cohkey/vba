Option Explicit

' =====================================================================================
' テスト用モジュール
' 概要:
'   - テストデータを Tool_Data / CSV_Data に生成
'   - 列の順序・列数・値をずらして差分を混在
'   - 比較処理を実行し、黄色セル数を集計して表示
' 参照:
'   - 本テストは、同一ブック内に CompareAndHighlightCsvDifferences が存在する前提
' 規約:
'   - ローカル変数: キャメルケース
'   - 引数名: スネークケース + ByVal/ByRef 明記
'   - 定数名: コンスタントネームケース
'   - 関数/サブ: パスカルケース + Public/Private 明記
'   - Call を付与
' =====================================================================================

Private Const SHEET_TOOL As String = "Tool_Data"
Private Const SHEET_CSV As String = "CSV_Data"
Private Const header_row As Long = 1
Private Const ID_HEADER As String = "ID"
Private Const YELLOW_BG As Long = vbYellow

' ------------------------------------------------------------
' 公開: スモールケース（手早く動作確認）
' 処理内容:
'   - 10行のフィクスチャ生成 → 比較実行 → 黄色セル数を表示
' 引数: なし
' 戻り値: なし
' ------------------------------------------------------------
Public Sub Test_SmallSample()
    On Error GoTo ERR_HANDLER
    Call BuildFixture_ToolAndCsv(10)
    Call ClearAllFills(ThisWorkbook.Worksheets(SHEET_CSV))
    
    Dim t As Double
    t = Timer
    Call CompareAndHighlightCsvDifferences
    Dim elapsed As Double
    elapsed = Timer - t
    
    Dim yellowCount As Long
    yellowCount = CountYellowCells(ThisWorkbook.Worksheets(SHEET_CSV), header_row)
    
    MsgBox "SmallSample 完了" & vbCrLf & _
           "黄色セル数: " & yellowCount & vbCrLf & _
           "処理時間(秒): " & Format(elapsed, "0.000"), vbInformation
    Exit Sub
ERR_HANDLER:
    MsgBox "Test_SmallSample エラー: " & Err.Number & " " & Err.Description, vbExclamation
End Sub

' ------------------------------------------------------------
' 公開: ベンチ用（既定 5,000 行）
' 処理内容:
'   - 5,000行のフィクスチャ生成（上限は環境に応じて変更可）
'   - 比較実行 → 黄色セル数と時間を表示
' 引数: なし
' 戻り値: なし
' ------------------------------------------------------------
Public Sub Test_Benchmark()
    On Error GoTo ERR_HANDLER
    Call BuildFixture_ToolAndCsv(5000) ' 環境に合わせて 40000 などに変更可
    Call ClearAllFills(ThisWorkbook.Worksheets(SHEET_CSV))
    
    Dim t As Double
    t = Timer
    Call CompareAndHighlightCsvDifferences
    Dim elapsed As Double
    elapsed = Timer - t
    
    Dim yellowCount As Long
    yellowCount = CountYellowCells(ThisWorkbook.Worksheets(SHEET_CSV), header_row)
    
    MsgBox "Benchmark 完了" & vbCrLf & _
           "黄色セル数: " & yellowCount & vbCrLf & _
           "処理時間(秒): " & Format(elapsed, "0.000"), vbInformation
    Exit Sub
ERR_HANDLER:
    MsgBox "Test_Benchmark エラー: " & Err.Number & " " & Err.Description, vbExclamation
End Sub

' ------------------------------------------------------------
' 処理内容:
'   Tool_Data / CSV_Data のテストデータを生成する。
'   - 両者ともに見出しあり
'   - 列セット/順序を変える（CSV側は列入替＆一部列欠落/追加）
'   - 値は一部ずらし/型違い/空白差 などを混ぜる
' 引数:
'   row_count : 生成するデータ件数
' 戻り値: なし
' ------------------------------------------------------------
Private Sub BuildFixture_ToolAndCsv(ByVal row_count As Long)
    Dim wsTool As Worksheet
    Dim wsCsv As Worksheet
    Set wsTool = PrepareSheet(SHEET_TOOL)
    Set wsCsv = PrepareSheet(SHEET_CSV)
    
    ' --- Tool 側の列（より多い列セット）
    ' ID, Name, Qty, Price, Date, Note, Extra
    Dim toolHeadersArr As Variant
    toolHeadersArr = Array("ID", "Name", "Qty", "Price", "Date", "Note", "Extra")
    
    ' --- CSV 側の列（順序を変更し、一部の列を欠落/別順）
    ' 例: Name, ID, Price, Qty, Note, Date  （Extra はCSVに出さない）
    Dim csvHeadersArr As Variant
    csvHeadersArr = Array("Name", "ID", "Price", "Qty", "Note", "Date")
    
    ' --- 見出し書き込み
    Call WriteHeaders(wsTool, toolHeadersArr)
    Call WriteHeaders(wsCsv, csvHeadersArr)
    
    ' --- データ本体（配列で一括書込）
    Call WriteToolData(wsTool, toolHeadersArr, row_count)
    Call WriteCsvDataWithDifferences(wsCsv, csvHeadersArr, row_count)
    
    ' 最終列幅調整（任意）
    wsTool.Columns.AutoFit
    wsCsv.Columns.AutoFit
End Sub

' ------------------------------------------------------------
' 処理内容:
'   指定シートを作成/初期化（既存ならクリア）
' 引数:
'   name_sheet : シート名
' 戻り値:
'   Worksheet : 準備済みシート
' ------------------------------------------------------------
Private Function PrepareSheet(ByVal name_sheet As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name_sheet)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = name_sheet
    End If
    
    ws.Cells.Clear
    Set PrepareSheet = ws
End Function

' ------------------------------------------------------------
' 処理内容:
'   見出し（1行目）を書き込む
' 引数:
'   ws_sheet      : 対象シート
'   headers_arr   : ヘッダー名のVariant配列(0-based)
' 戻り値: なし
' ------------------------------------------------------------
Private Sub WriteHeaders( _
    ByVal ws_sheet As Worksheet, _
    ByRef headers_arr As Variant _
)
    Dim i As Long
    For i = LBound(headers_arr) To UBound(headers_arr)
        ws_sheet.Cells(header_row, i + 1).Value = CStr(headers_arr(i))
    Next i
End Sub

' ------------------------------------------------------------
' 処理内容:
'   Tool_Data に row_count 分の行を生成して一括書込
'   値は比較に使いやすいように規則的に生成
' 引数:
'   ws_sheet      : Tool_Data シート
'   headers_arr   : ヘッダー名配列
'   row_count     : 行数
' 戻り値: なし
' ------------------------------------------------------------
Private Sub WriteToolData( _
    ByVal ws_sheet As Worksheet, _
    ByRef headers_arr As Variant, _
    ByVal row_count As Long _
)
    Dim colCount As Long
    colCount = UBound(headers_arr) - LBound(headers_arr) + 1
    
    Dim dataArr() As Variant
    ReDim dataArr(1 To row_count, 1 To colCount)
    
    Dim rowIndex As Long
    For rowIndex = 1 To row_count
        Dim idVal As String
        idVal = "ID" & Format$(rowIndex, "000000")
        
        Dim nameVal As String
        nameVal = "Name_" & rowIndex
        
        Dim qtyVal As Long
        qtyVal = (rowIndex Mod 10) + 1
        
        Dim priceVal As Double
        priceVal = 1000 + (rowIndex Mod 7) * 3.5
        
        Dim dateVal As Date
        dateVal = DateSerial(2024, ((rowIndex Mod 12) + 1), ((rowIndex Mod 27) + 1))
        
        Dim noteVal As String
        noteVal = "Note_" & rowIndex
        
        Dim extraVal As String
        extraVal = "EX" & (rowIndex Mod 5)
        
        Dim c As Long
        For c = 1 To colCount
            Select Case CStr(headers_arr(c - 1))
                Case "ID": dataArr(rowIndex, c) = idVal
                Case "Name": dataArr(rowIndex, c) = nameVal
                Case "Qty": dataArr(rowIndex, c) = qtyVal
                Case "Price": dataArr(rowIndex, c) = priceVal
                Case "Date": dataArr(rowIndex, c) = dateVal
                Case "Note": dataArr(rowIndex, c) = noteVal
                Case "Extra": dataArr(rowIndex, c) = extraVal
                Case Else: dataArr(rowIndex, c) = ""
            End Select
        Next c
    Next rowIndex
    
    ws_sheet.Range(ws_sheet.Cells(header_row + 1, 1), _
                   ws_sheet.Cells(header_row + row_count, colCount)).Value = dataArr
End Sub

' ------------------------------------------------------------
' 処理内容:
'   CSV_Data 用に Tool_Data と似たデータを生成するが、
'   ・列順を変更（ヘッダーで指定された順）
'   ・一部の行/列で意図的に差分を挿入
'   ・ID欠落/新規IDも混在（比較側の挙動確認用）
' 引数:
'   ws_sheet      : CSV_Data シート
'   headers_arr   : ヘッダー名配列
'   row_count     : 行数
' 戻り値: なし
' ------------------------------------------------------------
Private Sub WriteCsvDataWithDifferences( _
    ByVal ws_sheet As Worksheet, _
    ByRef headers_arr As Variant, _
    ByVal row_count As Long _
)
    Dim colCount As Long
    colCount = UBound(headers_arr) - LBound(headers_arr) + 1
    
    Dim dataArr() As Variant
    ReDim dataArr(1 To row_count, 1 To colCount)
    
    Dim rowIndex As Long
    For rowIndex = 1 To row_count
        ' Tool側と同じ規則でまず生成
        Dim idVal As String
        idVal = "ID" & Format$(rowIndex, "000000")
        
        Dim nameVal As String
        nameVal = "Name_" & rowIndex
        
        Dim qtyVal As Variant
        qtyVal = CStr((rowIndex Mod 10) + 1)          ' 文字列化して型差を演出
        
        Dim priceVal As Variant
        priceVal = CStr(1000 + (rowIndex Mod 7) * 3.5) ' 文字列化
        
        Dim dateVal As Variant
        dateVal = Format$(DateSerial(2024, ((rowIndex Mod 12) + 1), ((rowIndex Mod 27) + 1)), "yyyy/mm/dd")
        
        Dim noteVal As String
        noteVal = "Note_" & rowIndex
        
        ' --- 意図的な差分注入 ---
        ' 1) 50行おき: Name に末尾空白を付加（文字列Trim差の検証）
        If (rowIndex Mod 50) = 0 Then
            nameVal = nameVal & " "
        End If
        ' 2) 37行おき: Qty を +1（内容差）
        If (rowIndex Mod 37) = 0 Then
            qtyVal = CStr(((rowIndex Mod 10) + 1) + 1)
        End If
        ' 3) 101行おき: Price を 0.5 加算（内容差）
        If (rowIndex Mod 101) = 0 Then
            priceVal = CStr(CDbl(priceVal) + 0.5)
        End If
        ' 4) 77行おき: Note を空欄（内容差）
        If (rowIndex Mod 77) = 0 Then
            noteVal = ""
        End If
        ' 5) 111行おき: ID を未知に変更 → Tool側に存在しないIDを作成
        If (rowIndex Mod 111) = 0 Then
            idVal = "XID" & Format$(rowIndex, "000000")
        End If
        
        ' CSV側は Extra 列なしの想定（ヘッダー配列に含めていない）
        ' ヘッダー順で詰める
        Dim c As Long
        For c = 1 To colCount
            Select Case CStr(headers_arr(c - 1))
                Case "ID": dataArr(rowIndex, c) = idVal
                Case "Name": dataArr(rowIndex, c) = nameVal
                Case "Qty": dataArr(rowIndex, c) = qtyVal
                Case "Price": dataArr(rowIndex, c) = priceVal
                Case "Date": dataArr(rowIndex, c) = dateVal
                Case "Note": dataArr(rowIndex, c) = noteVal
                Case Else: dataArr(rowIndex, c) = ""
            End Select
        Next c
    Next rowIndex
    
    ws_sheet.Range(ws_sheet.Cells(header_row + 1, 1), _
                   ws_sheet.Cells(header_row + row_count, colCount)).Value = dataArr
End Sub

' ------------------------------------------------------------
' 処理内容:
'   指定シートのデータ部（ヘッダー以外）の塗りつぶしをクリア
' 引数:
'   ws_sheet : 対象シート
' 戻り値: なし
' ------------------------------------------------------------
Private Sub ClearAllFills(ByVal ws_sheet As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws_sheet.Cells(ws_sheet.Rows.Count, 1).End(xlUp).Row
    lastCol = ws_sheet.Cells(header_row, ws_sheet.Columns.Count).End(xlToLeft).Column
    If lastRow > header_row And lastCol >= 1 Then
        ws_sheet.Range(ws_sheet.Cells(header_row + 1, 1), ws_sheet.Cells(lastRow, lastCol)).Interior.Pattern = xlNone
    End If
End Sub

' ------------------------------------------------------------
' 処理内容:
'   指定シートで黄色セル（データ部のみ）の数をカウントする
' 引数:
'   ws_sheet   : 対象シート
'   header_row : ヘッダー行番号
' 戻り値:
'   Long       : 黄色セル数
' ------------------------------------------------------------
Private Function CountYellowCells( _
    ByVal ws_sheet As Worksheet, _
    ByVal header_row As Long _
) As Long
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws_sheet.Cells(ws_sheet.Rows.Count, 1).End(xlUp).Row
    lastCol = ws_sheet.Cells(header_row, ws_sheet.Columns.Count).End(xlToLeft).Column
    If lastRow <= header_row Or lastCol < 1 Then
        CountYellowCells = 0
        Exit Function
    End If
    
    Dim dataArr As Variant
    dataArr = ws_sheet.Range(ws_sheet.Cells(header_row + 1, 1), ws_sheet.Cells(lastRow, lastCol)).Value
    
    Dim r As Long, c As Long
    Dim cnt As Long
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            If ws_sheet.Cells(header_row + r, c).Interior.Color = YELLOW_BG Then
                cnt = cnt + 1
            End If
        Next c
    Next r
    CountYellowCells = cnt
End Function


