Option Explicit

' ========== CSV読み込み関数 ==========
Public Function ReadCsvFile(ByVal filePath As String) As Variant
    Dim fileNum As Integer
    Dim line As String
    Dim resultArr() As Variant
    Dim rowIndex As Long

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    rowIndex = 0
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        rowIndex = rowIndex + 1
        ReDim Preserve resultArr(1 To rowIndex)
        resultArr(rowIndex) = Split(line, ",")
    Loop

    Close #fileNum
    ReadCsvFile = resultArr
End Function
