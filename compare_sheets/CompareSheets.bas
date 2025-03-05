Option Explicit

' 差異の総数をカウントするサブルーチン
Private Sub CompareSheetsSummary(ByRef ws1 As Worksheet, ByRef ws2 As Worksheet, ParamArray rngList() As Variant)
    Dim lastRow1 As Long, lastRow2 As Long, lastRow As Long
    Dim diffCount As Long
    Dim i As Long
    Dim evalFormula As String
    Dim sh1Name As String, sh2Name As String

    ' 各シートのA列の最終行を取得し、短い方を比較範囲の終了行とする
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Min(lastRow1, lastRow2)

    ' シート名のシングルクォーテーション処理
    sh1Name = "'" & ws1.Name & "'"
    sh2Name = "'" & ws2.Name & "'"

    diffCount = 0
    For i = LBound(rngList) To UBound(rngList)
        ' 例: rngList(i) = "A3:B"　なら "A3:B10" という形で最終行番号を連結
        evalFormula = "SUMPRODUCT(--(" & sh1Name & "!" & rngList(i) & lastRow & _
                      "<>" & sh2Name & "!" & rngList(i) & lastRow & "))"
        diffCount = diffCount + Evaluate(evalFormula)
    Next i

    ' 結果はイミディエイトウィンドウに出力
    Debug.Print "Total differences: " & diffCount
End Sub

' 各セルごとに比較し、差異のあるセルのリストを出力するサブルーチン
Private Sub CompareSheetsDetails(ByRef ws1 As Worksheet, ByRef ws2 As Worksheet, ParamArray rngList() As Variant)
    Dim lastRow1 As Long, lastRow2 As Long, lastRow As Long
    Dim i As Long, r As Long, c As Long
    Dim rngStr As String
    Dim rngWs1 As Range, rngWs2 As Range
    Dim rowCount As Long, colCount As Long
    Dim cell1 As Range, cell2 As Range

    ' 各シートのA列の最終行を取得し、短い方を比較範囲の終了行とする
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Min(lastRow1, lastRow2)

    For i = LBound(rngList) To UBound(rngList)
        rngStr = rngList(i) ' 例: "A3:B" や "F3:H" など

        On Error Resume Next
        Set rngWs1 = ws1.Range(rngStr & lastRow)
        Set rngWs2 = ws2.Range(rngStr & lastRow)
        On Error GoTo 0

        If rngWs1 Is Nothing Or rngWs2 Is Nothing Then
            Debug.Print "Error: Invalid range " & rngStr
            GoTo NextRange
        End If

        rowCount = rngWs1.Rows.Count
        colCount = rngWs1.Columns.Count

        ' 各セルをループして比較
        For r = 1 To rowCount
            For c = 1 To colCount
                Set cell1 = rngWs1.Cells(r, c)
                Set cell2 = rngWs2.Cells(r, c)
                If cell1.Value <> cell2.Value Then
                    Debug.Print "Difference at " & ws1.Name & " " & cell1.Address & _
                                " (" & cell1.Value & ") vs " & ws2.Name & " " & cell2.Address & _
                                " (" & cell2.Value & ")"
                End If
            Next c
        Next r

NextRange:
        Set rngWs1 = Nothing
        Set rngWs2 = Nothing
    Next i
End Sub
