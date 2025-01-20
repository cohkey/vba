Option Explicit


' ========== アサート関数 ==========
Public Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        Debug.Print testName & " : OK (" & CStr(actual) & ")"
    Else
        Debug.Print testName & " : NG (Expected: " & CStr(expected) & ", Got: " & CStr(actual) & ")"
    End If
End Sub

Public Sub AssertNotEqual(ByVal testName As String, ByVal notExpected As Variant, ByVal actual As Variant)
    If notExpected = actual Then
        Debug.Print testName & " : NG (NotExpected: " & CStr(notExpected) & ", ButGot: " & CStr(actual) & ")"
    Else
        Debug.Print testName & " : OK (" & CStr(actual) & ")"
    End If
End Sub

Public Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    If condition Then
        Debug.Print testName & " : OK (True)"
    Else
        Debug.Print testName & " : NG (Expected True, Got False)"
    End If
End Sub

Public Sub AssertFalse(ByVal testName As String, ByVal condition As Boolean)
    If Not condition Then
        Debug.Print testName & " : OK (False)"
    Else
        Debug.Print testName & " : NG (Expected False, Got True)"
    End If
End Sub

' ========== テスト実行関数 ==========
Public Sub RunAllTests()
    Call RunUnitTests
    Call RunIntegrationTests
End Sub

Public Sub RunUnitTests()
    Debug.Print "--- Running Unit Tests ---"
    Call Test_Unit_Addition
    Call Test_Unit_Subtraction
End Sub

Public Sub RunIntegrationTests()
    Debug.Print "--- Running Integration Tests ---"
    Call Test_Integration_CsvRead
End Sub

' ========== テストケース ==========
Public Sub Test_Unit_Addition()
    Debug.Print "Running Test_Unit_Addition"
    Dim calc As New Calculator
    Call AssertEqual("Addition Test 1", 5, calc.Add(2, 3))
    Call AssertEqual("Addition Test 2", 0, calc.Add(-1, 1))
End Sub

Public Sub Test_Unit_Subtraction()
    Debug.Print "Running Test_Unit_Subtraction"
    Dim calc As New Calculator
    Call AssertEqual("Subtraction Test 1", 1, calc.Subtract(3, 2))
    Call AssertEqual("Subtraction Test 2", -2, calc.Subtract(-1, 1))
End Sub

Public Sub Test_Integration_CsvRead()
    Debug.Print "Running Test_Integration_CsvRead"
    ' テスト用CSVファイルパスを設定
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\testData.csv"

    If Dir(csvPath) = "" Then
        Debug.Print "Test_Integration_CsvRead : SKIP (File not found)"
        Exit Sub
    End If

    Dim result As Variant
    result = ReadCsvFile(csvPath)

    Call AssertEqual("CsvRead Row Count", 3, UBound(result, 1))
    Call AssertEqual("CsvRead Column Count", 2, UBound(result, 2))
End Sub





