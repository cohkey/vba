Sub CompareSheetsUsingEvaluate()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRow As Long
    Dim diffCount As Variant
    Dim sh1Name As String, sh2Name As String

    ' 対象のシートを設定（シート名は適宜変更してください）
    Set ws1 = Sheets("Sheet1")
    Set ws2 = Sheets("Sheet2")

    ' A列の最終行を取得（3行目以降の比較）
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow = Application.WorksheetFunction.Min(lastRow1, lastRow2)

    ' シート名にスペースが含まれる場合に備え、シングルクォーテーションで囲む
    sh1Name = "'" & ws1.Name & "'"
    sh2Name = "'" & ws2.Name & "'"

    ' Evaluateで2つのシートの範囲を比較し、違いのあるセル数をカウント
    diffCount = Evaluate("SUMPRODUCT(--(" & sh1Name & "!A3:H" & lastRow & "<>" & sh2Name & "!A3:H" & lastRow & "))")

    ' 結果の表示
    If diffCount = 0 Then
        MsgBox "シートは一致しています"
    Else
        MsgBox diffCount & "個の違いがあります"
    End If
End Sub
