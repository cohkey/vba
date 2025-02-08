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

    filePathOld = wsExec.Range("C2").Value
    filePathNew = wsExec.Range("C3").Value

    If filePathOld = "" Or filePathNew = "" Then
        MsgBox "Excelファイルのパスが未入力です。ExecutionシートのC2, C3を確認してください。", vbExclamation
        Exit Sub
    End If

    ' 1) 新規ブックを作成
    Dim wbCompare As Workbook
    Set wbCompare = Workbooks.Add(xlWBATWorksheet)  ' 新規ブック(1シート)

    ' ファイル名からシート名生成（不正文字の変換も実施）
    Dim rawOldName As String, rawNewName As String
    rawOldName = MakeSheetNameSafe(GetFileName(filePathOld))
    rawNewName = MakeSheetNameSafe(GetFileName(filePathNew))

    Dim oldSheetName As String, newSheetName As String
    oldSheetName = "Old_" & rawOldName
    newSheetName = "New_" & rawNewName

    ' Excelのシート名は31文字までのため、超えている場合は先頭31文字に切り詰める
    If Len(oldSheetName) > 31 Then
        oldSheetName = Left(oldSheetName, 31)
    End If
    If Len(newSheetName) > 31 Then
        newSheetName = Left(newSheetName, 31)
    End If

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

    ' 2) Excelファイルインポート（※最初のシートを選択）
    Call ImportExcelData(filePathOld, wbCompare, oldSheetName)
    Call ImportExcelData(filePathNew, wbCompare, newSheetName)

    ' 3) 比較(例: 閾値=0.4 → 40%)
    Dim threshold As Double
    threshold = 0.4
    Call CompareOldAndNew_VarThreshold(wbCompare, oldSheetName, newSheetName, "Result", threshold)

    ' 完了メッセージ
    wbCompare.Activate
    MsgBox "処理が完了しました。新規ブックに Old / New / Result シートがあります。", vbInformation

End Sub


'----------------------------------------------
' Excelファイルから1つ目のシートのデータを指定シートにコピーする
'----------------------------------------------
Private Sub ImportExcelData(ByVal filePath As String, _
                            ByRef wbTarget As Workbook, _
                            ByVal targetSheetName As String)

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
    Set wsSource = wbSource.Worksheets(1) ' 1つ目のシートを選択

    ' ソースのUsedRangeをコピーして、対象シートのA1セルに貼り付け
    wsSource.UsedRange.Copy Destination:=wsTarget.Range("A1")

    wbSource.Close SaveChanges:=False
End Sub

'----------------------------------------------
' 以下、残りは元のコード（GetFileName, MakeSheetNameSafe, 比較処理など）
'----------------------------------------------
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

    ' ピリオドはシート名に使用できないのでアンダースコアに変換
    tmp = Replace(tmp, ".", "_")

    ' 空文字になってしまった場合の代替文字列
    If tmp = "" Then tmp = "Sheet"

    ' ※ここではトリミングは行わず、Mainでプレフィックスを付けた後に全体の長さをチェックします。
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

    Dim i As Long, resultRow As Long
    resultRow = 2

    Dim oldRowArr() As Variant, newRowArr() As Variant

    '【旧データ側のループ】
    For i = 1 To UBound(oldDataArr, 1)
        oldRowArr = GetRowArray(oldDataArr, i, lastCol)
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

    '【新データ側で未マッチの行を追加】
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

    '==== ここからヘルパー列によるソート処理 ====

    ' ヘルパー列の位置を決定（結果シートは：古いデータ: 1～lastCol, Status: lastCol+1, 新しいデータ: lastCol+2～lastCol*2+1）
    ' ヘルパー列を追加する位置は、最後の列の右隣（例: lastCol*2+2）
    Dim helperCol As Long
    helperCol = lastCol * 2 + 2

    Dim r As Long
    For r = 2 To resultRow - 1
        Dim levelValue As Variant
        ' ステータスが「追加」の場合は、旧データ側は空なので、新しいデータ側（列 lastCol+2）の値を使用
        If wsResult.Cells(r, lastCol + 1).Value = "追加" Then
            levelValue = wsResult.Cells(r, lastCol + 2).Value
        Else
            levelValue = wsResult.Cells(r, 1).Value
        End If
        wsResult.Cells(r, helperCol).Value = levelValue
    Next r

    ' ヘッダー行にも「Level」項目名を付与（ヘッダーは1行目）
    wsResult.Cells(1, helperCol).Value = "SortLevel"

    ' 結果シート全体を、ヘルパー列（SortLevel）をキーにして昇順ソート
    With wsResult.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsResult.Range(wsResult.Cells(2, helperCol), wsResult.Cells(resultRow - 1, helperCol)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsResult.Range(wsResult.Cells(1, 1), wsResult.Cells(resultRow - 1, helperCol))
        .Header = xlYes
        .Apply
    End With

    ' ※必要ならヘルパー列を非表示にする
    wsResult.Columns(helperCol).Hidden = True
    '==== ここまでソート処理 ====

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
