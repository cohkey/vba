Option Explicit

'============================================================
' 概要   : 配列操作、文字列、安全変換、FS、シート入出力の汎用関数群
' 参照   : Microsoft Scripting Runtime
' 規約   : 事前バインディング / 1ステートメント1行
'============================================================

'-------------------- Null/Empty 安全化 ----------------------
'------------------------------------------------------------
' 処理内容 : Null/Empty を "" に変換して返す（比較・連結で安全）
' 引数     : ByVal vValue : 任意
' 戻り値   : String
' 例外     : なし
'------------------------------------------------------------
Public Function NzStr( _
    ByVal vValue As Variant _
) As String
    If IsNull(vValue) Then
        NzStr = vbNullString
    ElseIf IsEmpty(vValue) Then
        NzStr = vbNullString
    Else
        NzStr = CStr(vValue)
    End If
End Function

'-------------------- ファイルシステム -----------------------
'------------------------------------------------------------
' 処理内容 : フォルダパスの末尾\ を補正して返す
' 引数     : ByVal folderPath : 入力パス
' 戻り値   : String           : 末尾\ 付きに正規化
' 例外     : なし
'------------------------------------------------------------
Public Function NormalizeFolderPath( _
    ByVal folderPath As String _
) As String
    If Len(folderPath) = 0 Then
        NormalizeFolderPath = folderPath
    ElseIf Right$(folderPath, 1) = "\" Then
        NormalizeFolderPath = folderPath
    Else
        NormalizeFolderPath = folderPath & "\"
    End If
End Function

'------------------------------------------------------------
' 処理内容 : フォルダ内のCSVファイル一覧（指定ファイルを除く）
' 引数     : ByVal folderPath : フォルダパス
'          : ByVal excludeFile: 除外するファイル名（一致時）
' 戻り値   : Collection       : CSVファイル名の一覧
' 例外     : フォルダ不存在等で実行時エラー発生の可能性
'------------------------------------------------------------
Public Function ListCsvFiles( _
    ByVal folderPath As String, _
    ByVal excludeFile As String _
) As Collection
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject

    Dim fileColl As New Collection
    Dim f As Scripting.File
    For Each f In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "csv" Then
            If StrComp(f.Name, excludeFile, vbTextCompare) <> 0 Then
                Call fileColl.Add(f.Name)
            End If
        End If
    Next f

    Set ListCsvFiles = fileColl
End Function

'------------------------------------------------------------
' 処理内容 : 拡張子を除いたベース名を取得
' 引数     : ByVal fileName : ファイル名
' 戻り値   : String
' 例外     : なし
'------------------------------------------------------------
Public Function GetFileBaseName( _
    ByVal fileName As String _
) As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    GetFileBaseName = fso.GetBaseName(fileName)
End Function

'-------------------- シート入出力 ---------------------------
'------------------------------------------------------------
' 処理内容 : 2次元配列をシートへ一括出力（ヘッダ含む）
' 引数     : ByVal wsSheet : 出力先シート
'          : ByVal dataArr : 2次元配列
' 戻り値   : なし
' 例外     : なし
'------------------------------------------------------------
Public Sub DumpArrayToSheet( _
    ByVal wsSheet As Worksheet, _
    ByVal dataArr As Variant _
)
    Call wsSheet.Cells.Clear
    Dim rMax As Long
    Dim cMax As Long
    rMax = UBound(dataArr, 1)
    cMax = UBound(dataArr, 2)
    wsSheet.Range(wsSheet.Cells(1, 1), wsSheet.Cells(rMax, cMax)).Value = dataArr
    Call wsSheet.Columns.AutoFit
End Sub

'------------------------------------------------------------
' 処理内容 : 指定シートを取得。無ければ末尾に作成して返す
' 引数     : ByVal sheetName : シート名
' 戻り値   : Worksheet
' 例外     : なし
'------------------------------------------------------------
Public Function GetOrAddSheet( _
    ByVal sheetName As String _
) As Worksheet
    On Error Resume Next
    Set GetOrAddSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrAddSheet Is Nothing Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) _
        )
        ws.Name = sheetName
        Set GetOrAddSheet = ws
    End If
End Function

'-------------------- 配列/ヘッダ操作 -----------------------
'------------------------------------------------------------
' 処理内容 : 2次元配列の先頭行（ヘッダ）を1次元配列で返す
' 引数     : ByVal dataArr : 2次元配列（ヘッダ含む）
' 戻り値   : Variant(1 To cols)
' 例外     : なし
'------------------------------------------------------------
Public Function GetHeaderRow( _
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
' 引数     : ByVal headerArr : ヘッダ配列
'          : ByVal colName   : 列名
' 戻り値   : Long
' 例外     : なし
'------------------------------------------------------------
Public Function GetColumnIndex( _
    ByVal headerArr As Variant, _
    ByVal colName As String _
) As Long
    Dim c As Long
    For c = 1 To UBound(headerArr)
        If StrComp(CStr(headerArr(c)), colName, vbTextCompare) = 0 Then
            GetColumnIndex = c
            Exit Function
        End If
    Next c
    GetColumnIndex = 0
End Function

'------------------------------------------------------------
' 処理内容 : 2次元配列をクローンして返す
' 引数     : ByVal srcArr : 2次元配列
' 戻り値   : Variant
' 例外     : なし
'------------------------------------------------------------
Public Function CloneArray2D( _
    ByVal srcArr As Variant _
) As Variant
    Dim rMax As Long
    Dim cMax As Long
    rMax = UBound(srcArr, 1)
    cMax = UBound(srcArr, 2)

    Dim dstArr() As Variant
    ReDim dstArr(1 To rMax, 1 To cMax)

    Dim r As Long
    Dim c As Long
    For r = 1 To rMax
        For c = 1 To cMax
            dstArr(r, c) = srcArr(r, c)
        Next c
    Next r

    CloneArray2D = dstArr
End Function

'------------------------------------------------------------
' 処理内容 : キー列以外の列インデックス配列(1-based)を返す。空なら要素0
' 引数     : ByVal headerArr   : ヘッダ配列
'          : ByVal keyColIndex : キー列Index
' 戻り値   : Variant(配列)
' 例外     : なし
'------------------------------------------------------------
Public Function ListNonKeyColumns( _
    ByVal headerArr As Variant, _
    ByVal keyColIndex As Long _
) As Variant
    Dim c As Long
    Dim tmpArr() As Long
    ReDim tmpArr(1 To UBound(headerArr) - 1)

    Dim k As Long
    k = 0

    For c = 1 To UBound(headerArr)
        If c <> keyColIndex Then
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
' 引数     : ByVal v : 任意
' 戻り値   : Boolean
' 例外     : なし
'------------------------------------------------------------
Public Function ArrayHasElements( _
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
' 処理内容 : 結果配列の右端に新規列を追加（ヘッダ設定）
' 引数     : ByRef resultArr : 2次元配列
'          : ByVal newHeader : 新規ヘッダー名
' 戻り値   : なし
' 例外     : なし
'------------------------------------------------------------
Public Sub AddColumnToResult( _
    ByRef resultArr As Variant, _
    ByVal newHeader As String _
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

    tmpArr(1, cMax + 1) = newHeader
    resultArr = tmpArr
End Sub

'------------------------------------------------------------
' 処理内容 : 親ID→親配列の行番号の辞書を作成（Null安全）
' 引数     : ByVal parentArr        : 親2次元配列（ヘッダ含む）
'          : ByVal parentIdColIndex : 親ID列Index
' 戻り値   : Scripting.Dictionary   : Key=ID(String) / Item=行Index(Long)
' 例外     : なし
'------------------------------------------------------------
Public Function BuildIdToRowIndex( _
    ByVal parentArr As Variant, _
    ByVal parentIdColIndex As Long _
) As Scripting.Dictionary
    Dim idToRowDic As Scripting.Dictionary
    Set idToRowDic = New Scripting.Dictionary
    idToRowDic.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = UBound(parentArr, 1)

    Dim r As Long
    For r = 2 To lastRow
        Dim key As String
        key = NzStr(parentArr(r, parentIdColIndex))
        If LenB(key) > 0 Then
            If Not idToRowDic.Exists(key) Then
                Call idToRowDic.Add(key, r)
            End If
        End If
    Next r

    Set BuildIdToRowIndex = idToRowDic
End Function



'==================== ModCommon 追加分 ====================

'------------------------------------------------------------
' 処理内容 : ヘッダー行からの最終行・最終列を基準に、表全体を配列化
' 引数     : ByVal wsSheet    出力元シート
'          : ByVal headerRow  見出し行番号（通常は1）
'          : ByRef outArr     [out] 2次元配列（1-based, ヘッダ含む）
'          : ByRef lastRow    [out] 最終行
'          : ByRef lastCol    [out] 最終列
' 戻り値   : なし
'------------------------------------------------------------
Public Sub ReadUsedBodyAsArray( _
    ByVal wsSheet As Worksheet, _
    ByVal headerRow As Long, _
    ByRef outArr As Variant, _
    ByRef lastRow As Long, _
    ByRef lastCol As Long _
)
    lastRow = wsSheet.Cells(wsSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSheet.Cells(headerRow, wsSheet.Columns.Count).End(xlToLeft).Column

    If lastRow < headerRow Then
        lastRow = headerRow
    End If
    If lastCol < 1 Then
        lastCol = 1
    End If

    outArr = wsSheet.Range(wsSheet.Cells(1, 1), wsSheet.Cells(lastRow, lastCol)).Value
End Sub

'------------------------------------------------------------
' 処理内容 : キー用に値を正規化（IsError/Null/Empty→""、Trim）
' 引数     : ByVal vValue 任意
' 戻り値   : String
'------------------------------------------------------------
Public Function NormalizeKey( _
    ByVal vValue As Variant _
) As String
    If IsError(vValue) Then
        NormalizeKey = vbNullString
        Exit Function
    End If
    NormalizeKey = Trim$(NzStr(vValue))
End Function

'------------------------------------------------------------
' 処理内容 : 差分フラグ配列に従い、データ部を一括着色
' 引数     : ByVal wsSheet     対象シート
'          : ByVal headerRow   見出し行番号
'          : ByVal lastRow     最終行
'          : ByVal lastCol     最終列
'          : ByRef diffFlagArr (1..行数, 1..列数) のブール配列
'          : ByVal colorValue  塗り色（例: vbYellow）
' 戻り値   : なし
'------------------------------------------------------------
Public Sub ApplyHighlightByFlags( _
    ByVal wsSheet As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByRef diffFlagArr() As Boolean, _
    ByVal colorValue As Long _
)
    Dim r As Long
    Dim c As Long
    Dim baseRow As Long
    baseRow = headerRow + 1

    For r = baseRow To lastRow
        For c = 1 To lastCol
            If diffFlagArr(r - headerRow, c) Then
                wsSheet.Cells(r, c).Interior.Color = colorValue
            End If
        Next c
    Next r
End Sub

