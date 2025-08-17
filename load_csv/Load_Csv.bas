Option Explicit

'============================================================
' CSV → ADODB(Text) → [Row x Col]配列 → シート貼付（テスト一式）
' ・schema.iniを自動生成して全列Text固定（先頭ゼロ/混在型も保持）
' ・include_header=True のときはCSVの1行目もデータとして配列に含める
' ・既定の文字コードはCP932（Shift-JIS）。UTF-8なら DEFAULT_CHARSET_CODE を "65001" に変更
'============================================================

Private Const AD_READLINE As Long = -2
Private Const DEFAULT_CHARSET_CODE As String = "932"   ' "932"=Shift-JIS / "65001"=UTF-8

'------------------------------------------------------------
' CSVをADODB(Text Driver)で読み込み、[Row x Col] 配列を返す
' 引数:
'   ByVal csv_path As String       : CSVフルパス
'   ByVal has_header As Boolean    : CSVの1行目がヘッダーならTrue（ドライバに渡す情報）
'   ByVal include_header As Boolean: 返す配列に1行目（ヘッダー行）も含めたいならTrue
'   ByRef out_arr As Variant       : 結果2次元配列
' 戻り値:
'   なし（out_arrに格納）
'------------------------------------------------------------
Public Sub LoadCsvToArray( _
    ByVal csv_path As String, _
    ByVal has_header As Boolean, _
    ByVal include_header As Boolean, _
    ByRef out_arr As Variant _
)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim folderPath As String
    Dim fileName As String
    Dim frArr As Variant
    Dim rcArr As Variant
    Dim effectiveHdr As Boolean
    Dim hdrFlag As String

    If LenB(csv_path) = 0 Then
        Err.Raise vbObjectError + 200, , "csv_path が空です。"
    End If

    ' include_header=True の場合は、1行目を「データとして残す」ためドライバにはHDR=Noを渡す
    effectiveHdr = has_header And (Not include_header)

    folderPath = GetFolderPath(csv_path)
    fileName = GetFileName(csv_path)

    ' schema.ini（全列Text固定）を作成
    Call EnsureSchemaIniAllText(csv_path, effectiveHdr, DEFAULT_CHARSET_CODE)

    hdrFlag = IIf(effectiveHdr, "Yes", "No")

    Set cn = New ADODB.Connection
    cn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & folderPath & ";" & _
        "Extended Properties=""text;HDR=" & hdrFlag & ";FMT=Delimited;IMEX=1"";"
    cn.Open

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM [" & fileName & "]", cn, adOpenForwardOnly, adLockReadOnly

    If Not (rs.EOF And rs.BOF) Then
        frArr = rs.GetRows()                  ' [Field x Record]
        rcArr = TransposeFRToRC(frArr)        ' [Row x Col]
        out_arr = rcArr
    Else
        out_arr = Empty
    End If

    If rs.State <> adStateClosed Then rs.Close
    If cn.State <> adStateClosed Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub

'------------------------------------------------------------
' schema.ini を [全列 Text] で作成
' 引数:
'   charset_code: "932"(Shift-JIS) / "65001"(UTF-8)
'------------------------------------------------------------
Private Sub EnsureSchemaIniAllText( _
    ByVal csv_path As String, _
    ByVal has_header As Boolean, _
    ByVal charset_code As String _
)
    Dim folderPath As String
    Dim fileName As String
    Dim schemaPath As String
    Dim colCount As Long
    Dim ts As Object
    Dim i As Long

    folderPath = GetFolderPath(csv_path)
    fileName = GetFileName(csv_path)
    schemaPath = folderPath & IIf(Right$(folderPath, 1) = "\", "", "\") & "schema.ini"

    colCount = CountCsvFieldsOnFirstLine(csv_path, charset_code)

    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(schemaPath, 2, True, -1) ' Unicode
    ts.WriteLine "[" & fileName & "]"
    ts.WriteLine "Format=CSVDelimited"
    ts.WriteLine "ColNameHeader=" & IIf(has_header, "True", "False")
    ts.WriteLine "MaxScanRows=0"
    ts.WriteLine "CharacterSet=" & charset_code
    For i = 1 To colCount
        ts.WriteLine "Col" & CStr(i) & "=F" & CStr(i) & " Text"
    Next i
    ts.Close
End Sub

'------------------------------------------------------------
' 先頭1行の列数を数える（引用符対応の簡易カウント）
'------------------------------------------------------------
Private Function CountCsvFieldsOnFirstLine( _
    ByVal csv_path As String, _
    ByVal charset_code As String _
) As Long
    Dim stm As ADODB.Stream
    Dim s As String
    Dim i As Long
    Dim inQuote As Boolean
    Dim c As String
    Dim cnt As Long
    Dim charsetName As String

    charsetName = IIf(charset_code = "65001", "utf-8", "shift_jis")

    Set stm = New ADODB.Stream
    stm.Type = adTypeText
    stm.Charset = charsetName
    stm.Open
    stm.LoadFromFile csv_path
    s = ReadLineFromStream(stm)
    stm.Close
    Set stm = Nothing

    If LenB(s) = 0 Then
        CountCsvFieldsOnFirstLine = 1
        Exit Function
    End If

    cnt = 1
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = """" Then
            ' 連続ダブルクォート "" はエスケープ
            If inQuote And i < Len(s) And Mid$(s, i + 1, 1) = """" Then
                i = i + 1
            Else
                inQuote = Not inQuote
            End If
        ElseIf c = "," And Not inQuote Then
            cnt = cnt + 1
        End If
    Next i
    CountCsvFieldsOnFirstLine = cnt
End Function

'------------------------------------------------------------
' [Field x Record] → [Row x Col] に転置
'------------------------------------------------------------
Private Function TransposeFRToRC(ByVal frArr As Variant) As Variant
    Dim f As Long, r As Long
    Dim fCnt As Long, rCnt As Long
    Dim rcArr As Variant

    fCnt = UBound(frArr, 1) - LBound(frArr, 1) + 1
    rCnt = UBound(frArr, 2) - LBound(frArr, 2) + 1
    ReDim rcArr(1 To rCnt, 1 To fCnt)

    For r = 1 To rCnt
        For f = 1 To fCnt
            rcArr(r, f) = frArr(f - 1, r - 1)
        Next f
    Next r

    TransposeFRToRC = rcArr
End Function

'------------------------------------------------------------
' 2次元配列([Row x Col])をシートへ貼り付け（文字列書式）
'------------------------------------------------------------
Public Sub PasteArrayToSheet( _
    ByVal ws_sheet As Worksheet, _
    ByRef data_arr As Variant _
)
    Dim rowCount As Long
    Dim colCount As Long
    Dim dest As Range

    If IsEmpty(data_arr) Then
        ws_sheet.Cells.Clear
        Exit Sub
    End If

    rowCount = UBound(data_arr, 1)
    colCount = UBound(data_arr, 2)

    ws_sheet.Cells.Clear
    ws_sheet.Columns("A").Resize(, colCount).NumberFormat = "@"

    Set dest = ws_sheet.Cells(1, 1).Resize(rowCount, colCount)
    dest.Value2 = data_arr
End Sub

'------------------------------------------------------------
' ADODB.Stream から1行読み取り（CR/LFどちらでもOK）
'------------------------------------------------------------
Private Function ReadLineFromStream(ByVal stm As ADODB.Stream) As String
    ReadLineFromStream = stm.ReadText(AD_READLINE)  ' = adReadLine
End Function

'------------------------------------------------------------
' パス補助
'------------------------------------------------------------
Private Function GetFolderPath(ByVal file_path As String) As String
    Dim p As Long
    p = InStrRev(file_path, "\")
    GetFolderPath = IIf(p = 0, CurDir$, Left$(file_path, p - 1))
End Function

Private Function GetFileName(ByVal file_path As String) As String
    Dim p As Long
    p = InStrRev(file_path, "\")
    GetFileName = IIf(p = 0, file_path, Mid$(file_path, p + 1))
End Function

'------------------------------------------------------------
' テスト: サンプルCSVをTEMPに作って読込 → Sheet1へ貼付
'------------------------------------------------------------
Public Sub Test_CsvImportToSheet()
    Const TEST_FILE_NAME As String = "csv_import_test_sample.csv"

    Dim tempPath As String
    Dim csvPath As String
    Dim dataArr As Variant

    tempPath = Environ$("TEMP")
    If Right$(tempPath, 1) <> "\" Then
        tempPath = tempPath & "\"
    End If
    csvPath = tempPath & TEST_FILE_NAME

    Call CreateSampleCsv(csvPath)                              ' サンプルCSV生成（CP932想定）
    Call LoadCsvToArray(csvPath, True, True, dataArr)          ' ヘッダー行も含めて配列化
    Call PasteArrayToSheet(ThisWorkbook.Worksheets("Sheet1"), dataArr)

    MsgBox "テスト完了: " & csvPath, vbInformation
End Sub

'------------------------------------------------------------
' サンプルCSVを作成（日本語含む。既定コードページ=CP932）
'------------------------------------------------------------
Private Sub CreateSampleCsv(ByVal csv_path As String)
    Dim fso As Object
    Dim ts As Object
    Dim linesArr() As String
    Dim i As Long

    ReDim linesArr(0 To 5)
    linesArr(0) = "ID,Name,Note,Zip,Amount"
    linesArr(1) = "1,""Alice"",""abc,def"",""0123"",10"
    linesArr(2) = "2,""Bob"",""He said """"hello""""."",""0456"",2000"
    linesArr(3) = "3,""Carol"",,""0010"",0"
    linesArr(4) = "4,""データ"",""日本語もOK" & vbCrLf & "CrLf改行あり" & vbLf & "Lf改行あり"",""100-0001"",999999"
    linesArr(5) = "5,""末尾空欄"",,,""007"""

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csv_path, 2, True, 0) ' ForWriting, Create, Default ANSI(CP932)
    For i = LBound(linesArr) To UBound(linesArr)
        ts.WriteLine linesArr(i)
    Next i
    ts.Close
End Sub


