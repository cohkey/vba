Option Explicit

'============================================================
' 親CSV + 複数子CSV を ID で結合して1枚のシートに出力
' 兼：テスト用CSVの自動生成
'
' 参照（前バインディング）:
' - Microsoft ActiveX Data Objects 2.x Library
' - Microsoft Scripting Runtime
'
' コーディング規約:
' - 変数: camelCase / 配列xxxArr / Dictionary xxxDic / Collection xxxColl
' - 引数: snake_case + ByVal/ByRef
' - 定数: UPPER_SNAKE_CASE
' - Sub/Function: PascalCase（Public/Private 明示）
' - 関数呼び出し: Call を付ける
'============================================================

'==================== 設定値 ====================
Private Const ADO_PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
Private Const TEXT_EXTENDED As String = "text;HDR=YES;FMT=Delimited"
Private Const CONCAT_DELIM As String = " | "

Private Const TEST_FOLDER As String = "C:\Data\Csv\"
Private Const PARENT_FILE As String = "Parent.csv"

'==================== エントリ（動作確認） ====================
'------------------------------------------------------------
' 処理内容 : テストCSVを作ってから結合テストを実行
' 引数     : なし
' 戻り値   : なし
'------------------------------------------------------------
Public Sub Run_GenerateCsvs_And_Join()
    Call CreateTestCsvFiles(TEST_FOLDER)
    Call Run_JoinAllCsvs_Test
End Sub

'------------------------------------------------------------
' 処理内容 : 親+子CSVの結合テストを実行
' 引数     : なし
' 戻り値   : なし
'------------------------------------------------------------
Public Sub Run_JoinAllCsvs_Test()
    Dim ws As Worksheet
    Set ws = GetOrAddSheet("Joined")
    
    Call JoinAllCsvs( _
        TEST_FOLDER, _
        PARENT_FILE, _
        "ID", _
        "ParentID", _
        ws _
    )
End Sub

'==================== 本体：結合処理 ====================
'------------------------------------------------------------
' 処理内容 : 親CSVと子CSV群を親IDで結合し、シートに出力
' 引数     : ByVal folder_path    CSVフォルダパス
'            ByVal parent_csv     親CSVファイル名
'            ByVal parent_id_col  親ID列名
'            ByVal child_id_col   子の親ID列名（全子で同一想定）
'            ByVal target_sheet   出力先シート
' 戻り値   : なし
'------------------------------------------------------------
Public Sub JoinAllCsvs( _
    ByVal folder_path As String, _
    ByVal parent_csv As String, _
    ByVal parent_id_col As String, _
    ByVal child_id_col As String, _
    ByVal target_sheet As Worksheet _
)
    Dim t0 As Double
    t0 = Timer
    
    Dim normFolder As String
    normFolder = NormalizeFolderPath(folder_path)
    
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.enableEvents = False
    
    On Error GoTo CleanFail
    
    ' 親読み込み
    Dim parentArr As Variant
    parentArr = LoadCsvToArray(normFolder, parent_csv)
    If IsEmpty(parentArr) Then Err.Raise vbObjectError + 100, , "親CSV読込失敗: " & parent_csv
    
    Dim parentHeaderArr As Variant
    parentHeaderArr = GetHeaderRow(parentArr)
    
    Dim parentIdColIndex As Long
    parentIdColIndex = GetColumnIndex(parentHeaderArr, parent_id_col)
    If parentIdColIndex < 1 Then Err.Raise vbObjectError + 101, , "親ID列が見つかりません: " & parent_id_col
    
    ' 結果配列は親のクローン
    Dim resultArr As Variant
    resultArr = CloneArray2D(parentArr)
    
    ' 親ID→行Index
    Dim idToRowDic As Scripting.Dictionary
    Set idToRowDic = BuildIdToRowIndex(parentArr, parentIdColIndex)
    
    ' 子CSV列挙（親は除外）
    Dim childFilesColl As Collection
    Set childFilesColl = ListCsvFiles(normFolder, parent_csv)
    
    Dim i As Long
    For i = 1 To childFilesColl.Count
        Dim childName As String
        childName = CStr(childFilesColl(i))
        
        Dim childArr As Variant
        childArr = LoadCsvToArray(normFolder, childName)
        If IsEmpty(childArr) Then GoTo ContinueNextChild
        
        Dim childHeaderArr As Variant
        childHeaderArr = GetHeaderRow(childArr)
        
        Dim childKeyColIndex As Long
        childKeyColIndex = GetColumnIndex(childHeaderArr, child_id_col)
        If childKeyColIndex < 1 Then GoTo ContinueNextChild
        
        Dim childValueColsArr As Variant
        childValueColsArr = ListNonKeyColumns(childHeaderArr, childKeyColIndex)
        If Not ArrayHasElements(childValueColsArr) Then GoTo ContinueNextChild
        
        Dim childBase As String
        childBase = GetFileBaseName(childName)
        
        Dim addedMapDic As Scripting.Dictionary
        Set addedMapDic = EnsureChildColumns(resultArr, childValueColsArr, childHeaderArr, childBase)
        
        Call AppendChildIntoResult( _
            resultArr, _
            childArr, _
            idToRowDic, _
            parentIdColIndex, _
            childKeyColIndex, _
            childValueColsArr, _
            addedMapDic _
        )
ContinueNextChild:
    Next i
    
    Call DumpArrayToSheet(target_sheet, resultArr)
    
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.enableEvents = True
    
    Debug.Print "JoinAllCsvs done in " & Format$(Timer - t0, "0.000") & " sec"
    Exit Sub
    
CleanFail:
    Application.screenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.enableEvents = True
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

'==================== ADOでCSV→配列 ====================
'------------------------------------------------------------
' 処理内容 : ADOでCSVを2次元配列化（ヘッダ行含む、Nullはそのまま）
' 引数     : ByVal folder_path, ByVal file_name
' 戻り値   : Variant(1 To rows, 1 To cols)
'------------------------------------------------------------
Private Function LoadCsvToArray( _
    ByVal folder_path As String, _
    ByVal file_name As String _
) As Variant
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set cn = New ADODB.Connection
    cn.Open "Provider=" & ADO_PROVIDER & ";" & _
            "Data Source=" & folder_path & ";" & _
            "Extended Properties='" & TEXT_EXTENDED & "';"
    
    sql = "SELECT * FROM [" & file_name & "]"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rs.EOF And rs.BOF Then
        LoadCsvToArray = Empty
        GoTo FinallyProc
    End If
    
    LoadCsvToArray = RecordsetToArray(rs)
    
FinallyProc:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Exit Function
    
ErrHandler:
    LoadCsvToArray = Empty
    Resume FinallyProc
End Function

'------------------------------------------------------------
' 処理内容 : Recordset全体をヘッダ付き2次元配列へ変換
' 引数     : ByVal rs As ADODB.Recordset
' 戻り値   : Variant(1 To rows, 1 To cols)
'------------------------------------------------------------
Private Function RecordsetToArray( _
    ByVal rs As ADODB.Recordset _
) As Variant
    Dim fldCount As Long
    fldCount = rs.Fields.Count
    
    Dim dataArr As Variant
    dataArr = rs.GetRows()
    
    Dim rowCount As Long
    rowCount = UBound(dataArr, 2) + 1
    
    Dim outArr() As Variant
    ReDim outArr(1 To rowCount + 1, 1 To fldCount)
    
    Dim c As Long
    For c = 1 To fldCount
        outArr(1, c) = rs.Fields(c - 1).Name
    Next c
    
    Dim r As Long
    For r = 1 To rowCount
        For c = 1 To fldCount
            outArr(r + 1, c) = dataArr(c - 1, r - 1)
        Next c
    Next r
    
    RecordsetToArray = outArr
End Function

'==================== 配列/列・辞書ヘルパ ====================
'------------------------------------------------------------
' 処理内容 : 2次元配列の先頭行（ヘッダ）を返す
' 引数     : ByVal dataArr
' 戻り値   : 1次元配列(1 To cols)
'------------------------------------------------------------
Private Function GetHeaderRow( _
    ByVal dataArr As Variant _
) As Variant
    Dim cols As Long
    cols = UBound(dataArr, 2)
    
    Dim headerArr() As Variant
    ReDim headerArr(1 To cols)
    
    Dim c As Long
    For c = 1 To cols
        headerArr(c) = CStr(dataArr(1, c))
    Next c
    
    GetHeaderRow = headerArr
End Function

'------------------------------------------------------------
' 処理内容 : 指定列名の列番号を返す（1始まり、見つからなければ0）
' 引数     : ByVal headerArr, ByVal col_name
' 戻り値   : Long
'------------------------------------------------------------
Private Function GetColumnIndex( _
    ByVal headerArr As Variant, _
    ByVal col_name As String _
) As Long
    Dim c As Long
    For c = 1 To UBound(headerArr)
        If StrComp(CStr(headerArr(c)), col_name, vbTextCompare) = 0 Then
            GetColumnIndex = c
            Exit Function
        End If
    Next c
    GetColumnIndex = 0
End Function

'------------------------------------------------------------
' 処理内容 : 親ID→親配列の行番号の辞書を作成（Null安全）
' 引数     : ByVal parentArr, ByVal parent_id_col_index
' 戻り値   : Scripting.Dictionary (Key: String, Item: Long)
'------------------------------------------------------------
Private Function BuildIdToRowIndex( _
    ByVal parentArr As Variant, _
    ByVal parent_id_col_index As Long _
) As Scripting.Dictionary
    Dim idToRowDic As Scripting.Dictionary
    Set idToRowDic = New Scripting.Dictionary
    idToRowDic.CompareMode = vbTextCompare
    
    Dim lastRow As Long
    lastRow = UBound(parentArr, 1)
    
    Dim r As Long
    For r = 2 To lastRow
        Dim key As String
        key = NzStr(parentArr(r, parent_id_col_index))
        If LenB(key) > 0 Then
            If Not idToRowDic.Exists(key) Then
                Call idToRowDic.Add(key, r)
            End If
        End If
    Next r
    
    Set BuildIdToRowIndex = idToRowDic
End Function

'------------------------------------------------------------
' 処理内容 : 2次元配列をクローン
' 引数     : ByVal srcArr
' 戻り値   : Variant
'------------------------------------------------------------
Private Function CloneArray2D( _
    ByVal srcArr As Variant _
) As Variant
    Dim r As Long
    Dim c As Long
    Dim rMax As Long
    Dim cMax As Long
    
    rMax = UBound(srcArr, 1)
    cMax = UBound(srcArr, 2)
    
    Dim dstArr() As Variant
    ReDim dstArr(1 To rMax, 1 To cMax)
    
    For r = 1 To rMax
        For c = 1 To cMax
            dstArr(r, c) = srcArr(r, c)
        Next c
    Next r
    
    CloneArray2D = dstArr
End Function

'------------------------------------------------------------
' 処理内容 : キー列以外の列インデックス配列(1-based)を返す。空なら要素0
' 引数     : ByVal headerArr, ByVal key_col_index
' 戻り値   : Variant (配列)
'------------------------------------------------------------
Private Function ListNonKeyColumns( _
    ByVal headerArr As Variant, _
    ByVal key_col_index As Long _
) As Variant
    Dim c As Long
    Dim tmpArr() As Long
    ReDim tmpArr(1 To UBound(headerArr) - 1)
    
    Dim k As Long
    k = 0
    
    For c = 1 To UBound(headerArr)
        If c <> key_col_index Then
            k = k + 1
            tmpArr(k) = c
        End If
    Next c
    
    If k = 0 Then
        Dim emptyArr() As Variant
        ReDim emptyArr(1 To 0)
        ListNonKeyColumns = emptyArr
    Else
        Dim outArr() As Variant
        ReDim outArr(1 To k)
        For c = 1 To k
            outArr(c) = tmpArr(c)
        Next c
        ListNonKeyColumns = outArr
    End If
End Function

'------------------------------------------------------------
' 処理内容 : 配列が要素を持つか判定（1 To 0 の空配列も考慮）
' 引数     : ByVal v
' 戻り値   : Boolean
'------------------------------------------------------------
Private Function ArrayHasElements( _
    ByVal v As Variant _
) As Boolean
    If Not IsArray(v) Then
        ArrayHasElements = False
        Exit Function
    End If
    On Error Resume Next
    ArrayHasElements = (LBound(v) <= UBound(v))
    On Error GoTo 0
End Function

'------------------------------------------------------------
' 処理内容 : 結果配列に子列（子名_元列名）を追加し、マッピング辞書を返す
' 引数     : ByRef resultArr
'            ByVal childValueColsArr
'            ByVal childHeaderArr
'            ByVal child_base_name
' 戻り値   : Scripting.Dictionary  '子元列Index(String) → 結果列Index(Long)
'------------------------------------------------------------
Private Function EnsureChildColumns( _
    ByRef resultArr As Variant, _
    ByVal childValueColsArr As Variant, _
    ByVal childHeaderArr As Variant, _
    ByVal child_base_name As String _
) As Scripting.Dictionary
    Dim mapDic As Scripting.Dictionary
    Set mapDic = New Scripting.Dictionary
    
    Dim existingHeaderArr As Variant
    existingHeaderArr = GetHeaderRow(resultArr)
    
    Dim c As Long
    For c = 1 To UBound(childValueColsArr)
        Dim srcCol As Long
        srcCol = CLng(childValueColsArr(c))
        
        Dim newHeader As String
        newHeader = child_base_name & "_" & CStr(childHeaderArr(srcCol))
        
        Dim existCol As Long
        existCol = GetColumnIndex(existingHeaderArr, newHeader)
        
        If existCol < 1 Then
            Call AddColumnToResult(resultArr, newHeader)
            existCol = UBound(resultArr, 2)
            existingHeaderArr = GetHeaderRow(resultArr)
        End If
        
        mapDic.Add CStr(srcCol), existCol
    Next c
    
    Set EnsureChildColumns = mapDic
End Function

'------------------------------------------------------------
' 処理内容 : 結果配列の右端に列を追加（ヘッダ設定）
' 引数     : ByRef resultArr, ByVal new_header
' 戻り値   : なし
'------------------------------------------------------------
Private Sub AddColumnToResult( _
    ByRef resultArr As Variant, _
    ByVal new_header As String _
)
    Dim rMax As Long
    Dim cMax As Long
    rMax = UBound(resultArr, 1)
    cMax = UBound(resultArr, 2)
    
    Dim tmpArr() As Variant
    ReDim tmpArr(1 To rMax, 1 To cMax + 1)
    
    Dim r As Long
    Dim c As Long
    For r = 1 To rMax
        For c = 1 To cMax
            tmpArr(r, c) = resultArr(r, c)
        Next c
    Next r
    
    tmpArr(1, cMax + 1) = new_header
    resultArr = tmpArr
End Sub

'------------------------------------------------------------
' 処理内容 : 子配列の値を結果配列へ反映（1:nはセル内連結、Null安全）
' 引数     : ByRef resultArr
'            ByVal childArr
'            ByVal idToRowDic
'            ByVal parent_id_col_index
'            ByVal child_key_col_index
'            ByVal childValueColsArr
'            ByVal addedMapDic
' 戻り値   : なし
'------------------------------------------------------------
Private Sub AppendChildIntoResult( _
    ByRef resultArr As Variant, _
    ByVal childArr As Variant, _
    ByVal idToRowDic As Scripting.Dictionary, _
    ByVal parent_id_col_index As Long, _
    ByVal child_key_col_index As Long, _
    ByVal childValueColsArr As Variant, _
    ByVal addedMapDic As Scripting.Dictionary _
)
    Dim lastChildRow As Long
    lastChildRow = UBound(childArr, 1)
    
    Dim r As Long
    For r = 2 To lastChildRow
        Dim key As String
        key = NzStr(childArr(r, child_key_col_index))
        If LenB(key) = 0 Then GoTo ContinueNextRow
        If Not idToRowDic.Exists(key) Then GoTo ContinueNextRow
        
        Dim parentRow As Long
        parentRow = CLng(idToRowDic(key))
        
        Dim k As Long
        For k = 1 To UBound(childValueColsArr)
            Dim srcCol As Long
            srcCol = CLng(childValueColsArr(k))
            
            Dim dstCol As Long
            dstCol = CLng(addedMapDic(CStr(srcCol)))
            
            Dim vText As String
            vText = NzStr(childArr(r, srcCol))
            If LenB(vText) = 0 Then GoTo ContinueNextK
            
            Dim curText As String
            curText = NzStr(resultArr(parentRow, dstCol))
            
            If LenB(curText) = 0 Then
                resultArr(parentRow, dstCol) = vText
            Else
                resultArr(parentRow, dstCol) = curText & CONCAT_DELIM & vText
                ' 1:1限定にするなら、初回のみ代入する実装に変更
            End If
ContinueNextK:
        Next k
ContinueNextRow:
    Next r
End Sub

'------------------------------------------------------------
' 処理内容 : 2次元配列をシートへ一括出力
' 引数     : ByVal ws_sheet, ByVal dataArr
' 戻り値   : なし
'------------------------------------------------------------
Private Sub DumpArrayToSheet( _
    ByVal ws_sheet As Worksheet, _
    ByVal dataArr As Variant _
)
    Call ws_sheet.Cells.Clear
    Dim rMax As Long
    Dim cMax As Long
    rMax = UBound(dataArr, 1)
    cMax = UBound(dataArr, 2)
    ws_sheet.Range(ws_sheet.Cells(1, 1), ws_sheet.Cells(rMax, cMax)).Value = dataArr
    Call ws_sheet.Columns.AutoFit
End Sub

'==================== 汎用ユーティリティ ====================
'------------------------------------------------------------
' 処理内容 : Null/Empty を "" にして返す（比較・連結で安全）
' 引数     : ByVal v_value
' 戻り値   : String
'------------------------------------------------------------
Private Function NzStr( _
    ByVal v_value As Variant _
) As String
    If IsNull(v_value) Then
        NzStr = vbNullString
    ElseIf IsEmpty(v_value) Then
        NzStr = vbNullString
    Else
        NzStr = CStr(v_value)
    End If
End Function

'------------------------------------------------------------
' 処理内容 : フォルダ内のCSVファイル一覧（親を除く）
' 引数     : ByVal folder_path, ByVal exclude_file
' 戻り値   : Collection（ファイル名）
'------------------------------------------------------------
Private Function ListCsvFiles( _
    ByVal folder_path As String, _
    ByVal exclude_file As String _
) As Collection
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim coll As New Collection
    Dim f As Scripting.File
    For Each f In fso.GetFolder(folder_path).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "csv" Then
            If StrComp(f.Name, exclude_file, vbTextCompare) <> 0 Then
                Call coll.Add(f.Name)
            End If
        End If
    Next f
    
    Set ListCsvFiles = coll
End Function

'------------------------------------------------------------
' 処理内容 : フォルダパスの末尾\を補正
' 引数     : ByVal folder_path
' 戻り値   : 文字列
'------------------------------------------------------------
Private Function NormalizeFolderPath( _
    ByVal folder_path As String _
) As String
    If Len(folder_path) = 0 Then
        NormalizeFolderPath = folder_path
    ElseIf Right$(folder_path, 1) = "\" Then
        NormalizeFolderPath = folder_path
    Else
        NormalizeFolderPath = folder_path & "\"
    End If
End Function

'------------------------------------------------------------
' 処理内容 : 拡張子を除いたベース名を取得
' 引数     : ByVal file_name
' 戻り値   : 文字列
'------------------------------------------------------------
Private Function GetFileBaseName( _
    ByVal file_name As String _
) As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    GetFileBaseName = fso.GetBaseName(file_name)
End Function

'------------------------------------------------------------
' 処理内容 : 指定シートを取得。無ければ作成
' 引数     : ByVal sheet_name
' 戻り値   : Worksheet
'------------------------------------------------------------
Private Function GetOrAddSheet( _
    ByVal sheet_name As String _
) As Worksheet
    On Error Resume Next
    Set GetOrAddSheet = ThisWorkbook.Worksheets(sheet_name)
    On Error GoTo 0
    If GetOrAddSheet Is Nothing Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheet_name
        Set GetOrAddSheet = ws
    End If
End Function

'==================== テスト用CSV生成 ====================
'------------------------------------------------------------
' 処理内容 : テスト用CSVを一括生成
' 引数     : ByVal folder_path
' 戻り値   : なし
'------------------------------------------------------------
Public Sub CreateTestCsvFiles( _
    ByVal folder_path As String _
)
    Dim normFolder As String
    normFolder = NormalizeFolderPath(folder_path)
    
    Call EnsureFolder(normFolder)
    
    Call Write_ParentCsv(normFolder & PARENT_FILE)
    Call Write_OrdersCsv(normFolder & "Orders.csv")
    Call Write_NotesCsv(normFolder & "Notes.csv")
    Call Write_TagsCsv(normFolder & "Tags.csv")
    Call Write_ContactsCsv(normFolder & "Contacts.csv")
    Call Write_ScoresCsv(normFolder & "Scores.csv")
    
    MsgBox "テストCSVを作成しました: " & normFolder, vbInformation
End Sub

'------------------------------------------------------------
' 処理内容 : Parent.csv を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_ParentCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ID", "Name", "Country", "JoinDate", "Score")))
    Call linesColl.Add(MakeCsvLine(Array("1", "Alice", "US", "2024-01-10", "85")))
    Call linesColl.Add(MakeCsvLine(Array("2", "Bob", "UK", "2024-02-03", "92")))
    Call linesColl.Add(MakeCsvLine(Array("3", "Carol", "JP", "2024-03-21", "77")))
    Call linesColl.Add(MakeCsvLine(Array("4", "データ太郎", "JP", "2024-04-15", "88")))
    Call linesColl.Add(MakeCsvLine(Array("5", "Eve", "DE", "2024-05-01", "90")))
    Call linesColl.Add(MakeCsvLine(Array("6", "Frank", "FR", "2024-05-20", "73")))
    Call linesColl.Add(MakeCsvLine(Array("7", "Grace", "US", "2024-06-02", "95")))
    Call linesColl.Add(MakeCsvLine(Array("8", "Heidi", "CA", "2024-06-17", "81")))
    Call linesColl.Add(MakeCsvLine(Array("9", "Ivan", "RU", "2024-07-04", "69")))
    Call linesColl.Add(MakeCsvLine(Array("10", "Judy", "AU", "2024-07-25", "88")))
    Call linesColl.Add(MakeCsvLine(Array("11", "Ken", "SG", "2024-08-01", "84")))
    Call linesColl.Add(MakeCsvLine(Array("12", "Lena", "SE", "2024-08-09", "91")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Orders.csv（1:n）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_OrdersCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "OrderNo", "Amount", "Item")))
    Call linesColl.Add(MakeCsvLine(Array("1", "O-1001", "120.50", "Standard item")))
    Call linesColl.Add(MakeCsvLine(Array("1", "O-1002", "89.99", "abc,def")))
    Call linesColl.Add(MakeCsvLine(Array("2", "O-2001", "45.00", "He said ""hello"".")))
    Call linesColl.Add(MakeCsvLine(Array("3", "O-3001", "999999.00", "大容量パック")))
    Call linesColl.Add(MakeCsvLine(Array("3", "O-3002", "15.75", "Refill")))
    Call linesColl.Add(MakeCsvLine(Array("4", "O-4001", "0", "Free sample")))
    Call linesColl.Add(MakeCsvLine(Array("5", "O-5001", "250.00", "Bundle A")))
    Call linesColl.Add(MakeCsvLine(Array("7", "O-7001", "10.00", "Small")))
    Call linesColl.Add(MakeCsvLine(Array("7", "O-7002", "20.00", "Medium")))
    Call linesColl.Add(MakeCsvLine(Array("7", "O-7003", "30.00", "Large")))
    Call linesColl.Add(MakeCsvLine(Array("12", "O-12001", "1.00", "Last minute")))
    Call linesColl.Add(MakeCsvLine(Array("999", "O-X", "123", "No parent")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Notes.csv（1:n）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_NotesCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Note")))
    Call linesColl.Add(MakeCsvLine(Array("1", "First contact completed.")))
    Call linesColl.Add(MakeCsvLine(Array("2", "見積済み。追加要件あり。")))
    Call linesColl.Add(MakeCsvLine(Array("3", "Special handling required, see ticket #42.")))
    Call linesColl.Add(MakeCsvLine(Array("3", "'急ぎ'対応。担当: 佐藤")))
    Call linesColl.Add(MakeCsvLine(Array("5", "Ready to ship, hold until payment.")))
    Call linesColl.Add(MakeCsvLine(Array("8", "請求書再発行の依頼あり。")))
    Call linesColl.Add(MakeCsvLine(Array("11", "Follow-up next week.")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Tags.csv（1:n）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_TagsCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Tag")))
    Call linesColl.Add(MakeCsvLine(Array("1", "priority")))
    Call linesColl.Add(MakeCsvLine(Array("1", "beta")))
    Call linesColl.Add(MakeCsvLine(Array("2", "vip")))
    Call linesColl.Add(MakeCsvLine(Array("3", "internal")))
    Call linesColl.Add(MakeCsvLine(Array("3", "jp")))
    Call linesColl.Add(MakeCsvLine(Array("7", "bulk")))
    Call linesColl.Add(MakeCsvLine(Array("10", "oversea")))
    Call linesColl.Add(MakeCsvLine(Array("12", "promo")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Contacts.csv（1:1想定、欠損混在）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_ContactsCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Email", "Phone")))
    Call linesColl.Add(MakeCsvLine(Array("1", "alice@example.com", "+1-202-555-0101")))
    Call linesColl.Add(MakeCsvLine(Array("2", "bob@example.co.uk", "")))
    Call linesColl.Add(MakeCsvLine(Array("3", "", "+81-3-0000-0000")))
    Call linesColl.Add(MakeCsvLine(Array("4", "data.taro@example.jp", "+81-90-1234-5678")))
    Call linesColl.Add(MakeCsvLine(Array("5", "eve@example.de", "")))
    Call linesColl.Add(MakeCsvLine(Array("7", "grace@example.com", "+1-415-555-0199")))
    Call linesColl.Add(MakeCsvLine(Array("10", "judy@example.au", "")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Scores.csv（1:1想定）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_ScoresCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Quiz", "Rank")))
    Call linesColl.Add(MakeCsvLine(Array("1", "78", "B")))
    Call linesColl.Add(MakeCsvLine(Array("2", "91", "A")))
    Call linesColl.Add(MakeCsvLine(Array("3", "66", "C")))
    Call linesColl.Add(MakeCsvLine(Array("4", "88", "B")))
    Call linesColl.Add(MakeCsvLine(Array("7", "97", "A")))
    Call linesColl.Add(MakeCsvLine(Array("12", "59", "D")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'==================== CSV書き出しヘルパ ====================
'------------------------------------------------------------
' 処理内容 : フォルダを作成（存在しなければ）
' 引数     : ByVal folder_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub EnsureFolder( _
    ByVal folder_path As String _
)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    If Not fso.FolderExists(folder_path) Then
        Call fso.CreateFolder(folder_path)
    End If
End Sub

'------------------------------------------------------------
' 処理内容 : CSV1行を作成（正しいエスケープ）
' 引数     : ByVal valuesArr
' 戻り値   : 文字列
'------------------------------------------------------------
Private Function MakeCsvLine( _
    ByVal valuesArr As Variant _
) As String
    Dim i As Long
    Dim partsArr() As String
    ReDim partsArr(0 To UBound(valuesArr) - LBound(valuesArr))
    
    Dim idx As Long
    idx = 0
    
    For i = LBound(valuesArr) To UBound(valuesArr)
        partsArr(idx) = CsvEscape(CStr(valuesArr(i)))
        idx = idx + 1
    Next i
    
    MakeCsvLine = Join(partsArr, ",")
End Function

'------------------------------------------------------------
' 処理内容 : CSV用の値エスケープ（, / " / 改行）
' 引数     : ByVal s
' 戻り値   : 文字列
'------------------------------------------------------------
Private Function CsvEscape( _
    ByVal s As String _
) As String
    Dim needQuote As Boolean
    needQuote = (InStr(s, ",") > 0) _
                Or (InStr(s, """") > 0) _
                Or (InStr(s, vbCr) > 0) _
                Or (InStr(s, vbLf) > 0)
    
    If InStr(s, """") > 0 Then
        s = Replace$(s, """", """""")
    End If
    
    If needQuote Or Len(s) = 0 Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function

'------------------------------------------------------------
' 処理内容 : 行コレクションをファイルへ（CRLF、ANSI）
' 引数     : ByVal file_path, ByVal linesColl
' 戻り値   : なし
'------------------------------------------------------------
Private Sub SaveLinesToFile( _
    ByVal file_path As String, _
    ByVal linesColl As Collection _
)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim ts As Scripting.TextStream
    Set ts = fso.CreateTextFile(file_path, True, False) ' overwrite:=True, Unicode:=False(ANSI)
    
    Dim i As Long
    For i = 1 To linesColl.Count
        ts.Write linesColl(i)
        ts.Write vbCrLf
    Next i
    
    Call ts.Close
End Sub


