Sub CompareSheets(ws1 As Worksheet, ws2 As Worksheet)
    Dim lastRow1 As Long, lastRow2 As Long, lastRow As Long
    Dim diffCount1 As Variant, diffCount2 As Variant, diffCount As Variant
    Dim sh1Name As String, sh2Name As String

    ' 各シートのA列の最終行を取得し、短い方を採用（比較範囲の終了行）
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Min(lastRow1, lastRow2)

    ' シート名にスペースが含まれている場合に備え、シングルクォーテーションで囲む
    sh1Name = "'" & ws1.Name & "'"
    sh2Name = "'" & ws2.Name & "'"

    ' EvaluateとSUMPRODUCTでA列〜B列（3行目以降）の差異をカウント
    diffCount1 = Evaluate("SUMPRODUCT(--(" & sh1Name & "!A3:B" & lastRow & "<>" & sh2Name & "!A3:B" & lastRow & "))")

    ' 同様に、F列〜H列（3行目以降）の差異をカウント
    diffCount2 = Evaluate("SUMPRODUCT(--(" & sh1Name & "!F3:H" & lastRow & "<>" & sh2Name & "!F3:H" & lastRow & "))")

    ' 両方の差異を合計
    diffCount = diffCount1 + diffCount2

    ' 結果の表示
    If diffCount = 0 Then
        MsgBox "シートは一致しています"
    Else
        MsgBox diffCount & "個の違いがあります"
    End If
End Sub
