Sub CompareSheets(ws1 As Worksheet, ws2 As Worksheet, ParamArray rngList() As Variant)
    Dim lastRow1 As Long, lastRow2 As Long, lastRow As Long
    Dim diffCount As Long
    Dim i As Long
    Dim sh1Name As String, sh2Name As String
    Dim evalFormula As String

    ' 各シートのA列の最終行を取得し、短い方を比較対象の終了行とする
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Min(lastRow1, lastRow2)

    ' シート名にスペースが含まれる可能性を考慮してシングルクォーテーションで囲む
    sh1Name = "'" & ws1.Name & "'"
    sh2Name = "'" & ws2.Name & "'"

    diffCount = 0

    ' ParamArrayで渡された各範囲に対してループ処理
    ' 各rngList(i)は例として "A3:B" や "F3:H" のように指定
    For i = LBound(rngList) To UBound(rngList)
        ' Evaluate用の文字列を作成
        evalFormula = "SUMPRODUCT(--(" & sh1Name & "!" & rngList(i) & lastRow & _
                      "<>" & sh2Name & "!" & rngList(i) & lastRow & "))"
        ' Evaluateで差分をカウントして加算
        diffCount = diffCount + Evaluate(evalFormula)
    Next i

    ' 結果の表示
    If diffCount = 0 Then
        MsgBox "シートは一致しています"
    Else
        MsgBox diffCount & "個の違いがあります"
    End If
End Sub
