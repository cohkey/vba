Option Explicit

'--------------------------------------------------------------------------------
' 列インデックス (0-based)
'--------------------------------------------------------------------------------
Private Const CONST_COL_NAME As Long = 3
Private Const CONST_COL_CONTROLTYPEID As Long = 5
Private Const CONST_COL_CONTROLTYPELABEL As Long = 6
Private Const CONST_COL_FRAMEWORKID As Long = 9
Private Const CONST_COL_ARIAROLE As Long = 17

'==================================================================
' メインのエントリポイント
'==================================================================
Public Sub MainCompareProcedure()
    Dim wsExec As Worksheet
    Set wsExec = ThisWorkbook.Worksheets("Execution")

    Dim filePathOld As String, filePathNew As String
    filePathOld = wsExec.Range("C2").Value
    filePathNew = wsExec.Range("C3").Value

    If filePathOld = "" Or filePathNew = "" Then
        MsgBox "Excelファイルのパスが未入力です。ExecutionシートのC2, C3を確認してください。", vbExclamation
        Exit Sub
    End If

    ' 新規ブック作成
    Dim wbCompare As Workbook
    Set wbCompare = Workbooks.Add(xlWBATWorksheet)

    ' ファイル名からシート名生成（不正文字変換も実施）
    Dim rawOldName As String, rawNewName As String
    rawOldName = MakeSheetNameSafe(GetFileName(filePathOld))
    rawNewName = MakeSheetNameSafe(GetFileName(filePathNew))

    Dim sheetOldName As String, sheetNewName As String
    sheetOldName = "Old_" & rawOldName
    sheetNewName = "New_" & rawNewName

    ' シート名は最大31文字まで
    If Len(sheetOldName) > 31 Then sheetOldName = Left(sheetOldName, 31)
    If Len(sheetNewName) > 31 Then sheetNewName = Left(sheetNewName, 31)

    ' 既定シートを "Result" に変更
    Dim wsResult As Worksheet
    Set wsResult = wbCompare.Worksheets(1)
    wsResult.Name = "Result"

    ' 旧データ、新データ用シートを追加
    Dim wsOld As Worksheet, wsNew As Worksheet
    Set wsOld = wbCompare.Worksheets.Add(After:=wsResult)
    wsOld.Name = sheetOldName
    Set wsNew = wbCompare.Worksheets.Add(After:=wsOld)
    wsNew.Name = sheetNewName

    ' Excelファイルからデータインポート（1シート目のみ）
    ImportExcelData filePathOld, wbCompare, sheetOldName
    ImportExcelData filePathNew, wbCompare, sheetNewName

    ' 旧シートと新シートを比較し、結果シートを整列（Level順）する
    Dim threshold As Double: threshold = 0.4
    CompareSheets wbCompare, sheetOldName, sheetNewName, "Result", threshold

    wbCompare.Activate
    MsgBox "処理が完了しました。新規ブックに Old / New / Result シートがあります。", vbInformation
End Sub

'==================================================================
' ImportExcelData: 指定ファイルの1シート目のデータを、対象ブックの指定シートにコピー
'==================================================================
Private Sub ImportExcelData(ByVal filePath As String, ByRef wbTarget As Workbook, ByVal targetSheetName As String)
    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = wbTarget.Worksheets(targetSheetName)
    On Error GoTo 0

    If wsTarget Is Nothing Then
        MsgBox "ImportExcelData: " & targetSheetName & " シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim wbSource As Workbook
    On Error Resume Next
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    If wbSource Is Nothing Then
        MsgBox "ImportExcelData: ファイルを開けませんでした。 " & filePath, vbExclamation
        Exit Sub
    End If

    Dim wsSource As Worksheet
    Set wsSource = wbSource.Worksheets(1) ' 1つ目のシートを使用

    ' ソースのUsedRangeをコピーして、対象シートのA1に貼り付け
    wsSource.UsedRange.Copy Destination:=wsTarget.Range("A1")

    wbSource.Close SaveChanges:=False
End Sub

'==================================================================
' CompareSheets: 旧シートと新シートを比較して結果シートに出力＋ソート処理
'==================================================================
Private Sub CompareSheets(ByRef wb As Workbook, ByVal oldSheetName As String, ByVal newSheetName As String, ByVal resultSheetName As String, ByVal threshold As Double)
    Dim wsOld As Worksheet, wsNew As Worksheet, wsResult As Worksheet
    Set wsOld = wb.Worksheets(oldSheetName)
    Set wsNew = wb.Worksheets(newSheetName)
    Set wsResult = wb.Worksheets(resultSheetName)

    Dim lastRowOld As Long, lastRowNew As Long, lastCol As Long
    lastRowOld = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column

    wsResult.Cells.Clear
    wsResult.Name = resultSheetName
    Call CopyHeader(wsOld, wsNew, wsResult, lastCol)

    Dim oldDataArr As Variant, newDataArr As Variant
    If lastRowOld > 1 Then
        oldDataArr = wsOld.Range(wsOld.Cells(2, 1), wsOld.Cells(lastRowOld, lastCol)).Value
    Else
        ReDim oldDataArr(1 To 1, 1 To lastCol)
    End If
    If lastRowNew > 1 Then
        newDataArr = wsNew.Range(wsNew.Cells(2, 1), wsNew.Cells(lastRowNew, lastCol)).Value
    Else
        ReDim newDataArr(1 To 1, 1 To lastCol)
    End If

    Dim newMatched() As Boolean
    ReDim newMatched(1 To UBound(newDataArr, 1))

    Dim i As Long, resultRow As Long: resultRow = 2
    Dim oldRowArr() As Variant, newRowArr() As Variant
    For i = 1 To UBound(oldDataArr, 1)
        oldRowArr = GetRowArray(oldDataArr, i, lastCol)
        Dim matchedIndex As Long
        matchedIndex = FindBestMatchIndex(oldRowArr, newDataArr, newMatched, threshold)
        If matchedIndex > 0 Then
            newRowArr = GetRowArray(newDataArr, matchedIndex, lastCol)
            If CompareArrays(oldRowArr, newRowArr) Then
                WriteResultRow wsResult, resultRow, oldRowArr, newRowArr, "一致", lastCol
            Else
                WriteResultRow wsResult, resultRow, oldRowArr, newRowArr, "変更", lastCol
                HighlightDiffCells wsResult, resultRow, oldRowArr, newRowArr, 16750280, lastCol ' PURPLE_BG
            End If
            newMatched(matchedIndex) = True
        Else
            WriteResultRow wsResult, resultRow, oldRowArr, EmptyArray(lastCol), "削除", lastCol
            ColorRow wsResult, resultRow, 1, lastCol, 13027071 ' RED_BG
        End If
        resultRow = resultRow + 1
    Next i

    Dim j As Long
    For j = 1 To UBound(newDataArr, 1)
        If Not newMatched(j) Then
            newRowArr = GetRowArray(newDataArr, j, lastCol)
            WriteResultRow wsResult, resultRow, EmptyArray(lastCol), newRowArr, "追加", lastCol
            ColorRow wsResult, resultRow, lastCol + 2, lastCol * 2 + 1, 15123099 ' BLUE_BG
            resultRow = resultRow + 1
        End If
    Next j

    wsResult.Columns.AutoFit
    SortResultSheetByLevel wsResult, lastCol, resultRow
End Sub

'==================================================================
' SortResultSheetByLevel: ヘルパー列を追加してLevel（階層）順に結果シートをソート
'==================================================================
Private Sub SortResultSheetByLevel(ByRef wsResult As Worksheet, ByVal lastCol As Long, ByVal resultRow As Long)
    Dim helperCol As Long
    helperCol = lastCol * 2 + 2  ' 既存の列の右隣にヘルパー列を配置

    Dim r As Long
    For r = 2 To resultRow - 1
        Dim status As String
        status = wsResult.Cells(r, lastCol + 1).Value
        Dim levelValue As Variant
        If status = "追加" Then
            levelValue = wsResult.Cells(r, lastCol + 2).Value
        Else
            levelValue = wsResult.Cells(r, 1).Value
        End If
        wsResult.Cells(r, helperCol).Value = levelValue
    Next r
    wsResult.Cells(1, helperCol).Value = "SortLevel"

    With wsResult.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsResult.Range(wsResult.Cells(2, helperCol), wsResult.Cells(resultRow - 1, helperCol)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsResult.Range(wsResult.Cells(1, 1), wsResult.Cells(resultRow - 1, helperCol))
        .Header = xlYes
        .Apply
    End With

    wsResult.Columns(helperCol).Hidden = True
End Sub

'==================================================================
' 以下、補助関数群
'==================================================================
Private Function GetFileName(ByVal fullPath As String) As String
    Dim fName As String: fName = fullPath
    Dim pos As Long
    Do While InStr(fName, "\") > 0
        pos = InStr(fName, "\")
        fName = Mid(fName, pos + 1)
    Loop
    Do While InStr(fName, "/") > 0
        pos = InStr(fName, "/")
        fName = Mid(fName, pos + 1)
    Loop
    GetFileName = fName
End Function

Private Function MakeSheetNameSafe(ByVal rawName As String) As String
    Dim tmp As String: tmp = rawName
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", "[", "]")
    Dim c As Variant
    For Each c In invalidChars
        tmp = Replace(tmp, c, "")
    Next c
    tmp = Replace(tmp, ".", "_") ' ピリオドは使えないのでアンダースコアに変換
    If tmp = "" Then tmp = "Sheet"
    MakeSheetNameSafe = tmp
End Function

Private Sub CopyHeader(ByRef wsOld As Worksheet, ByRef wsNew As Worksheet, ByRef wsResult As Worksheet, ByVal lastCol As Long)
    Dim oldHeader As Variant, newHeader As Variant
    Dim c As Long
    oldHeader = wsOld.Range(wsOld.Cells(1, 1), wsOld.Cells(1, lastCol)).Value
    For c = 1 To lastCol
        wsResult.Cells(1, c).Value = oldHeader(1, c)
    Next c
    wsResult.Cells(1, lastCol + 1).Value = "Status"
    newHeader = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(1, lastCol)).Value
    For c = 1 To lastCol
        wsResult.Cells(1, lastCol + 1 + c).Value = newHeader(1, c)
    Next c
End Sub

Private Sub WriteResultRow(ByRef wsResult As Worksheet, ByVal resultRow As Long, ByRef arrOld As Variant, ByRef arrNew As Variant, ByVal statusStr As String, ByVal lastCol As Long)
    Dim c As Long
    For c = 0 To lastCol - 1
        wsResult.Cells(resultRow, c + 1).Value = arrOld(c)
    Next c
    wsResult.Cells(resultRow, lastCol + 1).Value = statusStr
    For c = 0 To lastCol - 1
        wsResult.Cells(resultRow, lastCol + 2 + c).Value = arrNew(c)
    Next c
End Sub

Private Sub HighlightDiffCells(ByRef wsResult As Worksheet, ByVal resultRow As Long, ByRef arrOld As Variant, ByRef arrNew As Variant, ByVal colorCode As Long, ByVal lastCol As Long)
    Dim c As Long
    For c = 0 To UBound(arrOld)
        If CStr(arrOld(c)) <> CStr(arrNew(c)) Then
            wsResult.Cells(resultRow, c + 1).Interior.Color = colorCode
            wsResult.Cells(resultRow, lastCol + 2 + c).Interior.Color = colorCode
        End If
    Next c
End Sub

Private Sub ColorRow(ByRef wsResult As Worksheet, ByVal resultRow As Long, ByVal startCol As Long, ByVal endCol As Long, ByVal colorCode As Long)
    Dim c As Long
    For c = startCol To endCol
        wsResult.Cells(resultRow, c).Interior.Color = colorCode
    Next c
End Sub

Private Function EmptyArray(ByVal colCount As Long) As Variant
    Dim arr() As Variant
    ReDim arr(0 To colCount - 1)
    Dim i As Long
    For i = 0 To colCount - 1
        arr(i) = ""
    Next i
    EmptyArray = arr
End Function

Private Function FindBestMatchIndex(ByRef arrOldRow As Variant, ByRef newDataArr As Variant, ByRef newMatched() As Boolean, ByVal matchThreshold As Double) As Long
    Dim bestIndex As Long: bestIndex = 0
    Dim bestScore As Double: bestScore = 0#
    Dim rowCount As Long: rowCount = UBound(newDataArr, 1)
    Dim colCount As Long: colCount = UBound(arrOldRow) + 1
    Dim i As Long
    For i = 1 To rowCount
        If Not newMatched(i) Then
            Dim arrNewRow() As Variant
            arrNewRow = GetRowArray(newDataArr, i, colCount)
            If IsSameKeyColumns(arrOldRow, arrNewRow) Then
                Dim score As Double
                score = CalculateMatchRatio(arrOldRow, arrNewRow)
                If score >= matchThreshold And score > bestScore Then
                    bestScore = score
                    bestIndex = i
                End If
            End If
        End If
    Next i
    FindBestMatchIndex = bestIndex
End Function

Private Function IsSameKeyColumns(ByRef arr1 As Variant, ByRef arr2 As Variant) As Boolean
    If UBound(arr1) <> UBound(arr2) Then
        IsSameKeyColumns = False
        Exit Function
    End If
    Dim nameOk As Boolean
    nameOk = (CStr(arr1(CONST_COL_NAME)) = CStr(arr2(CONST_COL_NAME)))
    Dim ctidOk As Boolean
    ctidOk = (CStr(arr1(CONST_COL_CONTROLTYPEID)) = CStr(arr2(CONST_COL_CONTROLTYPEID)))
    If nameOk And ctidOk Then
        IsSameKeyColumns = True
        Exit Function
    End If
    Dim labelOk As Boolean
    labelOk = (CStr(arr1(CONST_COL_CONTROLTYPELABEL)) = CStr(arr2(CONST_COL_CONTROLTYPELABEL)))
    Dim fwOk As Boolean
    fwOk = (CStr(arr1(CONST_COL_FRAMEWORKID)) = CStr(arr2(CONST_COL_FRAMEWORKID)))
    Dim ariaOk As Boolean
    ariaOk = (CStr(arr1(CONST_COL_ARIAROLE)) = CStr(arr2(CONST_COL_ARIAROLE)))
    If ctidOk And labelOk And fwOk And ariaOk Then
        IsSameKeyColumns = True
    Else
        IsSameKeyColumns = False
    End If
End Function

Private Function CalculateMatchRatio(ByRef arr1 As Variant, ByRef arr2 As Variant) As Double
    If UBound(arr1) <> UBound(arr2) Then
        CalculateMatchRatio = 0#
        Exit Function
    End If
    Dim matchCount As Long: matchCount = 0
    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
        If CStr(arr1(i)) = CStr(arr2(i)) Then matchCount = matchCount + 1
    Next i
    CalculateMatchRatio = matchCount / (UBound(arr1) - LBound(arr1) + 1)
End Function

Private Function CompareArrays(ByRef arr1 As Variant, ByRef arr2 As Variant) As Boolean
    If UBound(arr1) <> UBound(arr2) Then
        CompareArrays = False
        Exit Function
    End If
    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
        If CStr(arr1(i)) <> CStr(arr2(i)) Then
            CompareArrays = False
            Exit Function
        End If
    Next i
    CompareArrays = True
End Function

Private Function GetRowArray(ByRef data_array As Variant, ByVal row_num As Long, ByVal last_col As Long) As Variant
    Dim arr() As Variant
    ReDim arr(0 To last_col - 1)
    Dim c As Long
    For c = 1 To last_col
        arr(c - 1) = data_array(row_num, c)
    Next c
    GetRowArray = arr
End Function
