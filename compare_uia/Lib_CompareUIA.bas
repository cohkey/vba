Option Explicit

'--------------------------------------------------------------------------------
' 列インデックス(0ベース)
'   Name            = 4列目 → 3
'   ControlTypeID   = 6列目 → 5
'   ControlTypeLabel= 7列目 → 6
'   FrameworkId     = 10列目→ 9
'   AriaRole        = 18列目→ 17
'--------------------------------------------------------------------------------
Private Const CONST_ColIndex_Name             As Long = 3
Private Const CONST_ColIndex_ControlTypeID    As Long = 5
Private Const CONST_ColIndex_ControlTypeLabel As Long = 6
Private Const CONST_ColIndex_FrameworkId      As Long = 9
Private Const CONST_ColIndex_AriaRole         As Long = 17

'==================================================================
' メインのエントリポイント
'==================================================================
Public Sub MainCompareProcedure()

    Dim wsExec As Worksheet
    Set wsExec = ThisWorkbook.Worksheets("Execution")

    Dim filePathOld As String
    Dim filePathNew As String
    Dim encodeOld As String
    Dim encodeNew As String

    filePathOld = wsExec.Range("C2").Value
    filePathNew = wsExec.Range("C3").Value
    encodeOld = wsExec.Range("B2").Value
    encodeNew = wsExec.Range("B3").Value

    If filePathOld = "" Or filePathNew = "" Then
        MsgBox "CSVのパスが未入力です。ExecutionシートのC2, C3を確認してください。", vbExclamation
        Exit Sub
    End If
    If encodeOld = "" Or encodeNew = "" Then
        MsgBox "文字コードが未入力です。ExecutionシートのB2, B3を確認してください。", vbExclamation
        Exit Sub
    End If

    ' 1) 新規ブックを作成
    Dim wbCompare As Workbook
    Set wbCompare = Workbooks.Add(xlWBATWorksheet)  ' 新規ブック(1シート)

    ' CSVファイル名からシート名生成
    Dim oldSheetName As String, newSheetName As String
    oldSheetName = "Old_" & MakeSheetNameSafe(GetFileName(filePathOld))
    newSheetName = "New_" & MakeSheetNameSafe(GetFileName(filePathNew))

    ' 既定で存在するシートを "Result" に
    Dim wsTemp As Worksheet
    Set wsTemp = wbCompare.Worksheets(1)
    wsTemp.Name = "Result"

    ' Old_xxx, New_xxx シートを追加
    Dim wsOld As Worksheet, wsNew As Worksheet
    Set wsOld = wbCompare.Worksheets.Add(After:=wsTemp)
    wsOld.Name = oldSheetName

    Set wsNew = wbCompare.Worksheets.Add(After:=wsOld)
    wsNew.Name = newSheetName

    ' 2) CSVインポート
    Call ImportCSV(filePathOld, encodeOld, wbCompare, oldSheetName)
    Call ImportCSV(filePathNew, encodeNew, wbCompare, newSheetName)

    ' 3) 比較(例: 閾値=0.4 → 40%)
    Dim threshold As Double
    threshold = 0.4
    Call CompareOldAndNew_VarThreshold(wbCompare, oldSheetName, newSheetName, "Result", threshold)

    ' 完了メッセージ
    wbCompare.Activate
    MsgBox "処理が完了しました。新規ブックに Old / New / Result シートがあります。", vbInformation

End Sub

'----------------------------------------------
Private Sub ImportCSV(ByVal filePath As String, _
                      ByVal encodeType As String, _
                      ByRef wbTarget As Workbook, _
                      ByVal targetSheetName As String)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wbTarget.Worksheets(targetSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "ImportCSV: " & targetSheetName & " シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim platformCode As Long
    platformCode = GetPlatformCode(encodeType)
    If platformCode = 0 Then
        MsgBox "サポート外の文字コードです。utf-8 or shift-jis を指定してください。", vbExclamation
        Exit Sub
    End If

    Dim qt As QueryTable
    For Each qt In ws.QueryTables
        qt.Delete
    Next qt

    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .TextFilePlatform = platformCode
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With

End Sub

Private Function GetPlatformCode(ByVal encodeType As String) As Long

    Select Case LCase(encodeType)
        Case "utf-8", "utf8"
            GetPlatformCode = 65001
        Case "shift-jis", "sjis", "shift_jis"
            GetPlatformCode = 932
        Case Else
            GetPlatformCode = 0
    End Select

End Function

Private Function GetFileName(ByVal fullPath As String) As String
    Dim fName As String
    fName = fullPath

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
    Dim tmp As String
    tmp = rawName

    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", "[", "]")

    Dim c As Variant
    For Each c In invalidChars
        tmp = Replace(tmp, c, "")
    Next c

    If Len(tmp) > 31 Then
        tmp = Left(tmp, 31)
    End If

    If tmp = "" Then tmp = "Sheet"

    MakeSheetNameSafe = tmp
End Function

'******************************************************
Public Sub CompareOldAndNew_VarThreshold( _
    ByRef wb As Workbook, _
    ByVal oldSheetName As String, _
    ByVal newSheetName As String, _
    ByVal resultSheetName As String, _
    ByVal matchThreshold As Double)

    Const RED_BG As Long = 13027071
    Const BLUE_BG As Long = 15123099
    Const PURPLE_BG As Long = 16750280

    Dim wsOld As Worksheet, wsNew As Worksheet, wsResult As Worksheet
    Set wsOld = wb.Worksheets(oldSheetName)
    Set wsNew = wb.Worksheets(newSheetName)
    Set wsResult = wb.Worksheets(resultSheetName)

    Dim lastRowOld As Long, lastRowNew As Long, lastCol As Long
    lastRowOld = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).row
    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).row
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

    Dim i As Long, resultRow As Long
    resultRow = 2

    Dim oldRowArr() As Variant, newRowArr() As Variant

    For i = 1 To UBound(oldDataArr, 1)

        oldRowArr = GetRowArray(oldDataArr, i, lastCol)

        ' FindBestMatchIndex
        Dim matchedIndex As Long
        matchedIndex = FindBestMatchIndex(oldRowArr, newDataArr, newMatched, matchThreshold)

        If matchedIndex > 0 Then
            newRowArr = GetRowArray(newDataArr, matchedIndex, lastCol)

            If CompareArrays(oldRowArr, newRowArr) Then
                Call WriteResultRow(wsResult, resultRow, oldRowArr, newRowArr, "一致", lastCol)
            Else
                Call WriteResultRow(wsResult, resultRow, oldRowArr, newRowArr, "変更", lastCol)
                Call HighlightDiffCells(wsResult, resultRow, oldRowArr, newRowArr, PURPLE_BG, lastCol)
            End If
            newMatched(matchedIndex) = True
        Else
            Call WriteResultRow(wsResult, resultRow, oldRowArr, EmptyArray(lastCol), "削除", lastCol)
            Call ColorRow(wsResult, resultRow, 1, lastCol, RED_BG)
        End If

        resultRow = resultRow + 1
    Next i

    Dim j As Long
    For j = 1 To UBound(newDataArr, 1)
        If Not newMatched(j) Then
            newRowArr = GetRowArray(newDataArr, j, lastCol)
            Call WriteResultRow(wsResult, resultRow, EmptyArray(lastCol), newRowArr, "追加", lastCol)
            Call ColorRow(wsResult, resultRow, lastCol + 2, lastCol * 2 + 1, BLUE_BG)
            resultRow = resultRow + 1
        End If
    Next j

    wsResult.Columns.AutoFit
End Sub

'******************************************************
' FindBestMatchIndex
'******************************************************
Private Function FindBestMatchIndex( _
    ByRef arrOldRow As Variant, _
    ByRef newDataArr As Variant, _
    ByRef newMatched() As Boolean, _
    ByVal matchThreshold As Double) As Long

    Dim bestIndex As Long: bestIndex = 0
    Dim bestScore As Double: bestScore = 0#

    Dim rowCount As Long
    rowCount = UBound(newDataArr, 1)

    Dim colCount As Long
    colCount = UBound(arrOldRow) + 1

    Dim i As Long
    For i = 1 To rowCount
        If Not newMatched(i) Then
            Dim arrNewRow() As Variant
            arrNewRow = GetRowArray(newDataArr, i, colCount)

            ' (A) キー列が一致するか
            If IsSameKeyColumns(arrOldRow, arrNewRow) Then
                ' (B) 行全体の類似度
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

'******************************************************
' IsSameKeyColumns
'   ・OR 条件に変更：
'     (1) Name+ControlTypeID が一致
'         または
'     (2) ControlTypeID & ControlTypeLabel & FrameworkId & AriaRole がすべて一致
'******************************************************
Private Function IsSameKeyColumns(ByRef arr1 As Variant, _
                                  ByRef arr2 As Variant) As Boolean

    If UBound(arr1) <> UBound(arr2) Then
        IsSameKeyColumns = False
        Exit Function
    End If

    ' 1) Name + ControlTypeID が一致？
    Dim nameOk As Boolean
    nameOk = (CStr(arr1(CONST_ColIndex_Name)) = CStr(arr2(CONST_ColIndex_Name)))

    Dim ctidOk As Boolean
    ctidOk = (CStr(arr1(CONST_ColIndex_ControlTypeID)) = CStr(arr2(CONST_ColIndex_ControlTypeID)))

    If nameOk And ctidOk Then
        ' これだけで同じ要素と見なす
        IsSameKeyColumns = True
        Exit Function
    End If

    ' 2) ControlTypeID/Label/FrameworkId/AriaRole が全部一致？
    Dim labelOk As Boolean
    labelOk = (CStr(arr1(CONST_ColIndex_ControlTypeLabel)) = CStr(arr2(CONST_ColIndex_ControlTypeLabel)))

    Dim fwOk As Boolean
    fwOk = (CStr(arr1(CONST_ColIndex_FrameworkId)) = CStr(arr2(CONST_ColIndex_FrameworkId)))

    Dim ariaOk As Boolean
    ariaOk = (CStr(arr1(CONST_ColIndex_AriaRole)) = CStr(arr2(CONST_ColIndex_AriaRole)))

    If ctidOk And labelOk And fwOk And ariaOk Then
        IsSameKeyColumns = True
    Else
        IsSameKeyColumns = False
    End If

End Function

'******************************************************
Private Function CalculateMatchRatio(ByRef arr1 As Variant, _
                                     ByRef arr2 As Variant) As Double

    If UBound(arr1) <> UBound(arr2) Then
        CalculateMatchRatio = 0#
        Exit Function
    End If

    Dim matchCount As Long: matchCount = 0
    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
        If CStr(arr1(i)) = CStr(arr2(i)) Then
            matchCount = matchCount + 1
        End If
    Next i

    CalculateMatchRatio = matchCount / (UBound(arr1) - LBound(arr1) + 1)
End Function

Private Function CompareArrays(ByRef arr1 As Variant, _
                               ByRef arr2 As Variant) As Boolean

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

Private Function GetRowArray(ByRef data_array As Variant, _
                             ByVal row_num As Long, _
                             ByVal last_col As Long) As Variant

    Dim arr() As Variant
    ReDim arr(0 To last_col - 1)

    Dim c As Long
    For c = 1 To last_col
        arr(c - 1) = data_array(row_num, c)
    Next c

    GetRowArray = arr
End Function

Private Sub CopyHeader(ByRef ws_old As Worksheet, _
                       ByRef ws_new As Worksheet, _
                       ByRef ws_result As Worksheet, _
                       ByVal last_col As Long)

    Dim oldHeader As Variant
    Dim newHeader As Variant
    Dim c As Long

    oldHeader = ws_old.Range(ws_old.Cells(1, 1), ws_old.Cells(1, last_col)).Value
    For c = 1 To last_col
        ws_result.Cells(1, c).Value = oldHeader(1, c)
    Next c

    ws_result.Cells(1, last_col + 1).Value = "Status"

    newHeader = ws_new.Range(ws_new.Cells(1, 1), ws_new.Cells(1, last_col)).Value
    For c = 1 To last_col
        ws_result.Cells(1, last_col + 1 + c).Value = newHeader(1, c)
    Next c

End Sub

Private Sub WriteResultRow(ByRef ws_result As Worksheet, _
                           ByVal result_row As Long, _
                           ByRef arr_old As Variant, _
                           ByRef arr_new As Variant, _
                           ByVal status_str As String, _
                           ByVal last_col As Long)

    Dim c As Long
    For c = 0 To last_col - 1
        ws_result.Cells(result_row, 1 + c).Value = arr_old(c)
    Next c

    ws_result.Cells(result_row, last_col + 1).Value = status_str

    For c = 0 To last_col - 1
        ws_result.Cells(result_row, last_col + 2 + c).Value = arr_new(c)
    Next c

End Sub

Private Sub HighlightDiffCells(ByRef ws_result As Worksheet, _
                               ByVal result_row As Long, _
                               ByRef arr_old As Variant, _
                               ByRef arr_new As Variant, _
                               ByVal color_code As Long, _
                               ByVal last_col As Long)

    Dim c As Long
    For c = 0 To UBound(arr_old)
        If CStr(arr_old(c)) <> CStr(arr_new(c)) Then
            ws_result.Cells(result_row, 1 + c).Interior.Color = color_code
            ws_result.Cells(result_row, last_col + 2 + c).Interior.Color = color_code
        End If
    Next c

End Sub

Private Sub ColorRow(ByRef ws_result As Worksheet, _
                     ByVal result_row As Long, _
                     ByVal start_col As Long, _
                     ByVal end_col As Long, _
                     ByVal color_code As Long)

    Dim c As Long
    For c = start_col To end_col
        ws_result.Cells(result_row, c).Interior.Color = color_code
    Next c

End Sub

Private Function EmptyArray(ByVal col_count As Long) As Variant

    Dim arr() As Variant
    ReDim arr(0 To col_count - 1)

    Dim i As Long
    For i = 0 To col_count - 1
        arr(i) = ""
    Next i

    EmptyArray = arr

End Function
