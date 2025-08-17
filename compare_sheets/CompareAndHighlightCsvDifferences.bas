Option Explicit

' =====================================================================================
' 概要  :
'   Tool_Dataシート と CSV_Dataシート を ID で突合し、
'   CSV_Data側の列見出しに対応する Tool_Data の同名列と値を比較。
'   値が異なる CSV_Data 側セルのみ黄色背景にする。
'   Tool_Data に存在しない ID の行は、CSV_Data のその行の全データ列を黄色にする。
' 方式:
'   - 両シートをVariant配列に一括読込
'   - ヘッダー名→列番号のマップ(Dictionary)作成
'   - Tool_Dataは ID→行配列 のDictionary化（O(1)で行参照）
'   - CSV_Data を走査し差分をフラグ配列に保持 → 最後に一括で塗り
' 前提:
'   - 列見出しは1行目（定数 HEADER_ROW で変更可）
'   - 両シートに ID 列が存在（定数 ID_HEADER で見出し名を指定）
' 注意:
'   - 参照設定: Microsoft Scripting Runtime
' =====================================================================================

Private Const SHEET_TOOL As String = "Sheet1"
Private Const SHEET_CSV As String = "Sheet2"
Private Const ID_HEADER As String = "ID"
Private Const header_row As Long = 1
Private Const YELLOW_BG As Long = vbYellow

' ------------------------------------------
' 公開エントリポイント
' ------------------------------------------
' 処理内容:
'   上記の通り。差分セルを CSV_Data 側で黄色塗り。
Public Sub CompareAndHighlightCsvDifferences()
    Dim appCalcMode As XlCalculation
    Dim screenUpdating As Boolean
    Dim enableEvents As Boolean
    
    ' --- 高速化オプション退避＆OFF ---
    appCalcMode = Application.Calculation
    screenUpdating = Application.screenUpdating
    enableEvents = Application.enableEvents
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CLEANUP
    
    Dim wsTool As Worksheet
    Dim wsCsv As Worksheet
    Set wsTool = ThisWorkbook.Worksheets(SHEET_TOOL)
    Set wsCsv = ThisWorkbook.Worksheets(SHEET_CSV)
    
    Dim toolArr As Variant
    Dim csvArr As Variant
    Dim toolLastRow As Long, toolLastCol As Long
    Dim csvLastRow As Long, csvLastCol As Long
    
    ' --- 範囲確定＆配列取得 ---
    Call GetUsedBodyAsArray(wsTool, header_row, toolArr, toolLastRow, toolLastCol)
    Call GetUsedBodyAsArray(wsCsv, header_row, csvArr, csvLastRow, csvLastCol)
    
    If toolLastRow < header_row + 1 Or csvLastRow < header_row + 1 Then
        ' データ無し
        GoTo CLEANUP
    End If
    
    ' --- ヘッダーマップ（列名→列番号） ---
    Dim toolHeaderDic As Scripting.Dictionary
    Dim csvHeaderDic As Scripting.Dictionary
    Set toolHeaderDic = CreateHeaderIndexMap(toolArr, toolLastCol)
    Set csvHeaderDic = CreateHeaderIndexMap(csvArr, csvLastCol)
    
    ' ID列インデックス取得（両シート）
    Dim toolIdCol As Long
    Dim csvIdCol As Long
    toolIdCol = GetRequiredHeaderIndex(toolHeaderDic, ID_HEADER)
    csvIdCol = GetRequiredHeaderIndex(csvHeaderDic, ID_HEADER)
    
    ' --- Tool 側: ID→行配列 の辞書 ---
    Dim toolRowDic As Scripting.Dictionary
    Set toolRowDic = BuildIdToRowDictionary(toolArr, toolLastRow, toolIdCol)
    
    ' --- CSV側と比較: 差分フラグ配列（データ部のみ 2次元Boolean） ---
    Dim rowCount As Long, colCount As Long
    rowCount = csvLastRow - header_row               ' データ行数
    colCount = csvLastCol                            ' データ列数（ヘッダ含む列幅で処理）
    Dim diffFlagArr() As Boolean
    ReDim diffFlagArr(1 To rowCount, 1 To colCount)  ' 行:1=2行目, 列:1=列1
    
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim csvId As String
    Dim toolRowArr As Variant
    Dim hasToolRow As Boolean
    
    ' 比較対象は「CSVの各列見出しがToolにも存在する」という要件に基づき、
    ' CSVのヘッダーを主としてループ。ID列は除外して比較。
    For rowIndex = header_row + 1 To csvLastRow
        csvId = NormalizeId(csvArr(rowIndex, csvIdCol))
        hasToolRow = toolRowDic.Exists(csvId)
        If hasToolRow Then
            toolRowArr = toolRowDic(csvId)  ' 1-based 行配列
        End If
        
        For colIndex = 1 To csvLastCol
            ' ヘッダー行・ID列はスキップ（IDがTool未存在時は後で一括塗り）
            If rowIndex > header_row Then
                If colIndex <> csvIdCol Then
                    Dim headerName As String
                    headerName = CStr(csvArr(header_row, colIndex))
                    
                    If Len(headerName) > 0 Then
                        ' CSVの列はToolに必ず存在する前提（要件）
                        Dim toolCol As Long
                        toolCol = GetRequiredHeaderIndex(toolHeaderDic, headerName)
                        
                        If hasToolRow Then
                            Dim vCsv As Variant
                            Dim vTool As Variant
                            vCsv = csvArr(rowIndex, colIndex)
                            vTool = toolRowArr(toolCol)
                            
                            If Not AreCellValuesEqual(vCsv, vTool) Then
                                diffFlagArr(rowIndex - header_row, colIndex) = True
                            End If
                        Else
                            ' Toolに該当IDが無い → 当該行の全データ列を差分扱い
                            diffFlagArr(rowIndex - header_row, colIndex) = True
                        End If
                    End If
                End If
            End If
        Next colIndex
    Next rowIndex
    
    ' --- 差分フラグに基づき CSV_Data 側を一括で黄色塗り ---
    Call ApplyHighlightByFlags(wsCsv, header_row, csvLastRow, csvLastCol, diffFlagArr, YELLOW_BG)
    
CLEANUP:
    ' --- 復元 ---
    Application.Calculation = appCalcMode
    Application.screenUpdating = screenUpdating
    Application.enableEvents = enableEvents
End Sub

' ------------------------------------------
' 配列取得: ヘッダー行含むUsedRange相当を安全に配列化
' ------------------------------------------
' 引数:
'   ws_sheet     : 対象シート
'   header_row   : 見出しの行番号
'   outArr       : [out] Variant 2次元配列(1-based)
'   lastRow/Col  : [out] 終端行・列
Private Sub GetUsedBodyAsArray( _
    ByVal ws_sheet As Worksheet, _
    ByVal header_row As Long, _
    ByRef outArr As Variant, _
    ByRef lastRow As Long, _
    ByRef lastCol As Long _
)
    ' UsedRangeに依存せず、ヘッダー行からの最終行・列を推定
    Dim r As Long, c As Long
    lastRow = ws_sheet.Cells(ws_sheet.Rows.Count, 1).End(xlUp).Row
    lastCol = ws_sheet.Cells(header_row, ws_sheet.Columns.Count).End(xlToLeft).Column
    
    If lastRow < header_row Then lastRow = header_row
    If lastCol < 1 Then lastCol = 1
    
    outArr = ws_sheet.Range(ws_sheet.Cells(1, 1), ws_sheet.Cells(lastRow, lastCol)).Value
End Sub

' ------------------------------------------
' ヘッダー行(1行目)から列名→列番号の辞書を作る
' ------------------------------------------
' 返り値:
'   Key: ヘッダー名(String)、Item: 列番号(Long)
Private Function CreateHeaderIndexMap( _
    ByRef dataArr As Variant, _
    ByVal last_col As Long _
) As Scripting.Dictionary
    Dim mapDic As Scripting.Dictionary
    Set mapDic = New Scripting.Dictionary
    mapDic.CompareMode = TextCompare
    
    Dim colIndex As Long
    For colIndex = 1 To last_col
        Dim headerName As String
        headerName = CStr(dataArr(header_row, colIndex))
        If Len(headerName) > 0 Then
            If Not mapDic.Exists(headerName) Then
                Call mapDic.Add(headerName, colIndex)
            End If
        End If
    Next colIndex
    
    Set CreateHeaderIndexMap = mapDic
End Function

' ------------------------------------------
' 必須ヘッダーの列番号を取得（無ければ実行時エラー）
' ------------------------------------------
Private Function GetRequiredHeaderIndex( _
    ByVal header_dic As Scripting.Dictionary, _
    ByVal header_name As String _
) As Long
    If Not header_dic.Exists(header_name) Then
        Err.Raise vbObjectError + 101, "GetRequiredHeaderIndex", _
                  "必要なヘッダーが見つかりません: " & header_name
    End If
    GetRequiredHeaderIndex = CLng(header_dic(header_name))
End Function

' ------------------------------------------
' Tool側の「ID→行配列」辞書を作る
' ------------------------------------------
' 返り値:
'   Key: 正規化ID(String)、Item: 行配列(Variant; 1-based, 全列)
Private Function BuildIdToRowDictionary( _
    ByRef toolArr As Variant, _
    ByVal tool_last_row As Long, _
    ByVal tool_id_col As Long _
) As Scripting.Dictionary
    Dim rowDic As Scripting.Dictionary
    Set rowDic = New Scripting.Dictionary
    rowDic.CompareMode = TextCompare
    
    Dim rowIndex As Long
    For rowIndex = header_row + 1 To tool_last_row
        Dim idKey As String
        idKey = NormalizeId(toolArr(rowIndex, tool_id_col))
        If Len(idKey) > 0 Then
            If Not rowDic.Exists(idKey) Then
                ' 参照のまま保持（速度優先）
                Call rowDic.Add(idKey, toolArr) ' まず配列ごと保持…
                ' ただし、行配列だけを取り出して保持したいので再代入
                rowDic(idKey) = GetRowAsArray(toolArr, rowIndex)
            End If
        End If
    Next rowIndex
    
    Set BuildIdToRowDictionary = rowDic
End Function

' ------------------------------------------
' 配列の指定行を 1-based の行配列として取り出す
' ------------------------------------------
Private Function GetRowAsArray( _
    ByRef srcArr As Variant, _
    ByVal row_index As Long _
) As Variant
    Dim lastCol As Long
    lastCol = UBound(srcArr, 2)
    Dim rowArr() As Variant
    ReDim rowArr(1 To lastCol)
    
    Dim colIndex As Long
    For colIndex = 1 To lastCol
        rowArr(colIndex) = srcArr(row_index, colIndex)
    Next colIndex
    
    GetRowAsArray = rowArr
End Function

' ------------------------------------------
' IDの正規化（文字列化＋Trim）
' ------------------------------------------
Private Function NormalizeId(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then
        s = ""
    ElseIf IsNull(v) Or v = "" Then
        s = ""
    Else
        s = CStr(v)
    End If
    NormalizeId = Trim$(s)
End Function

' ------------------------------------------
' 値の等価判定
' ルール:
'   - 両方数値として解釈可 → CDbl比較
'   - 両方日付として解釈可 → CDbl(Date)比較（内部シリアルで一致判定）
'   - それ以外            → 文字列Trim同士の一致（大文字小文字区別）
' ------------------------------------------
Private Function AreCellValuesEqual( _
    ByVal v1 As Variant, _
    ByVal v2 As Variant _
) As Boolean
    ' 空判定の先にエラーを潰す
    If IsError(v1) Or IsError(v2) Then
        AreCellValuesEqual = False
        Exit Function
    End If
    
    ' null/空文字は同一扱い
    If (IsNull(v1) Or v1 = "") And (IsNull(v2) Or v2 = "") Then
        AreCellValuesEqual = True
        Exit Function
    End If
    
    ' 数値比較
    If IsNumeric(v1) And IsNumeric(v2) Then
        AreCellValuesEqual = (CDbl(v1) = CDbl(v2))
        Exit Function
    End If
    
    ' 日付比較（両方Dateとして解釈できる場合）
    If IsDate(v1) And IsDate(v2) Then
        AreCellValuesEqual = (CDbl(CDate(v1)) = CDbl(CDate(v2)))
        Exit Function
    End If
    
    ' 文字列比較（厳密：ケースセンシティブ、不要なら StrComp の vbTextCompare へ変更）
    Dim s1 As String
    Dim s2 As String
    s1 = Trim$(CStr(v1))
    s2 = Trim$(CStr(v2))
    AreCellValuesEqual = (StrComp(s1, s2, vbBinaryCompare) = 0)
End Function

' ------------------------------------------
' 差分フラグ配列に基づき一括で着色
' diffFlagArr: (1..行数, 1..列数)
' ------------------------------------------
Private Sub ApplyHighlightByFlags( _
    ByVal ws_sheet As Worksheet, _
    ByVal header_row As Long, _
    ByVal last_row As Long, _
    ByVal last_col As Long, _
    ByRef diffFlagArr() As Boolean, _
    ByVal color_value As Long _
)
    Dim r As Long, c As Long
    Dim baseRow As Long
    baseRow = header_row + 1
    
    ' 一度だけのループで塗る（Union多用は避ける）
    For r = baseRow To last_row
        For c = 1 To last_col
            If diffFlagArr(r - header_row, c) Then
                ws_sheet.Cells(r, c).Interior.Color = color_value
            Else
                ' 以前の色が残っている可能性に配慮するなら解除
                ' ws_sheet.Cells(r, c).Interior.Pattern = xlNone
            End If
        Next c
    Next r
End Sub


