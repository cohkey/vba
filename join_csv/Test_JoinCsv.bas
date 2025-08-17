Option Explicit

'============================================================
' テスト用CSVを生成するユーティリティ
' - 出力先: C:\Data\Csv（既定。必要ならConstを書き換え）
' - 親: Parent.csv（列: ID, Name, Country, JoinDate, Score）
' - 子: Orders/Notes/Tags/Contacts/Scores（すべて ParentID を外部キーとする）
' 特徴:
' - 1:n（同一ParentIDの複数行）を含む子CSVを用意
' - カンマ/ダブルクォート/日本語を含む値を適切にエスケープ
' - 親に存在しないParentIDの子行も混ぜ、無視される流れを検証可能
'============================================================

'==================== 設定値 ====================
Private Const TEST_FOLDER As String = "C:\Data\Csv\"
Private Const PARENT_FILE As String = "Parent.csv"

'==================== エントリ ====================
'------------------------------------------------------------
' 処理内容 : テスト用CSVを生成し、既存のJoinテストを実行
' 引数     : なし
' 戻り値   : なし
'------------------------------------------------------------
Public Sub Run_GenerateCsvs_And_Join()
    ' 1) テストCSV生成
    Call CreateTestCsvFiles(TEST_FOLDER)
    
    ' 2) あなたの既存テスト（Run_JoinAllCsvs_Test）を実行
    '    既出の JoinAllCsvs を使う想定（親: Parent.csv, 親ID: ID, 子ID: ParentID）
    Call Run_JoinAllCsvs_Test
End Sub

'------------------------------------------------------------
' 処理内容 : テスト用CSVを一括生成
' 引数     : ByVal folder_path  出力先フォルダ
' 戻り値   : なし
'------------------------------------------------------------
Public Sub CreateTestCsvFiles( _
    ByVal folder_path As String _
)
    Dim normFolder As String
    normFolder = NormalizeFolderPath(folder_path)
    
    ' フォルダ作成
    Call EnsureFolder(normFolder)
    
    ' 親CSV
    Call Write_ParentCsv(normFolder & PARENT_FILE)
    
    ' 子CSV（1:n含む）
    Call Write_OrdersCsv(normFolder & "Orders.csv")
    Call Write_NotesCsv(normFolder & "Notes.csv")
    Call Write_TagsCsv(normFolder & "Tags.csv")
    Call Write_ContactsCsv(normFolder & "Contacts.csv")
    Call Write_ScoresCsv(normFolder & "Scores.csv")
    
    MsgBox "テストCSVを作成しました: " & normFolder, vbInformation
End Sub

'==================== 各CSVの定義 ====================

'------------------------------------------------------------
' 処理内容 : Parent.csv を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_ParentCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    ' ヘッダ
    Call linesColl.Add(MakeCsvLine(Array("ID", "Name", "Country", "JoinDate", "Score")))
    
    ' データ（IDは一意）
    Call linesColl.Add(MakeCsvLine(Array("1", "Alice", "US", "2024-01-10", "85")))
    Call linesColl.Add(MakeCsvLine(Array("2", "Bob", "UK", "2024-02-03", "92")))
    Call linesColl.Add(MakeCsvLine(Array("3", "Carol", "JP", "2024-03-21", "77")))
    Call linesColl.Add(MakeCsvLine(Array("4", "データ太郎", "JP", "2024-04-15", "88")))
    Call linesColl.Add(MakeCsvLine(Array("5", "Eve", "DE", "2024-05-01", "90")))
    Call linesColl.Add(MakeCsvLine(Array("6", "Frank", "FR", "2024-05-20", "73")))
    Call linesColl.Add(MakeCsvLine(Array("7", "Grace", "US", "2024-06-02", "95")))
    Call linesColl.Add(MakeCsvLine(Array("8", "Heidi", "CA", "2024-06-17", "81")))
    Call linesColl.Add(MakeCsvLine(Array("9", "Ivan", "RU", "2024-07-04", "69")))
    Call linesColl.Add(MakeCsvLine(Array("10", "Judy", "AU", "2024-07-25", "88")))
    Call linesColl.Add(MakeCsvLine(Array("11", "Ken", "SG", "2024-08-01", "84")))
    Call linesColl.Add(MakeCsvLine(Array("12", "Lena", "SE", "2024-08-09", "91")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Orders.csv（1:n）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
' 備考     : カンマ/ダブルクォート含む値を混在、存在しないParentIDも含める
'------------------------------------------------------------
Private Sub Write_OrdersCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "OrderNo", "Amount", "Item")))
    
    Call linesColl.Add(MakeCsvLine(Array("1", "O-1001", "120.50", "Standard item")))
    Call linesColl.Add(MakeCsvLine(Array("1", "O-1002", "89.99", "abc,def")))                    ' カンマ
    Call linesColl.Add(MakeCsvLine(Array("2", "O-2001", "45.00", "He said ""hello"".")))         ' ダブルクォート
    Call linesColl.Add(MakeCsvLine(Array("3", "O-3001", "999999.00", "大容量パック")))
    Call linesColl.Add(MakeCsvLine(Array("3", "O-3002", "15.75", "Refill")))
    Call linesColl.Add(MakeCsvLine(Array("4", "O-4001", "0", "Free sample")))
    Call linesColl.Add(MakeCsvLine(Array("5", "O-5001", "250.00", "Bundle A")))
    Call linesColl.Add(MakeCsvLine(Array("7", "O-7001", "10.00", "Small")))
    Call linesColl.Add(MakeCsvLine(Array("7", "O-7002", "20.00", "Medium")))
    Call linesColl.Add(MakeCsvLine(Array("7", "O-7003", "30.00", "Large")))
    Call linesColl.Add(MakeCsvLine(Array("12", "O-12001", "1.00", "Last minute")))
    Call linesColl.Add(MakeCsvLine(Array("999", "O-X", "123", "No parent")))                     ' 親に無いID
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Notes.csv（1:n、長文/日本語含む）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_NotesCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Note")))
    
    Call linesColl.Add(MakeCsvLine(Array("1", "First contact completed.")))
    Call linesColl.Add(MakeCsvLine(Array("2", "見積済み。追加要件あり。")))
    Call linesColl.Add(MakeCsvLine(Array("3", "Special handling required, see ticket #42.")))
    Call linesColl.Add(MakeCsvLine(Array("3", "'急ぎ'対応。担当: 佐藤"))) ' 全角引用符含む
    Call linesColl.Add(MakeCsvLine(Array("5", "Ready to ship, hold until payment.")))
    Call linesColl.Add(MakeCsvLine(Array("8", "請求書再発行の依頼あり。")))
    Call linesColl.Add(MakeCsvLine(Array("11", "Follow-up next week.")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Tags.csv（1:n、単語タグ）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_TagsCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Tag")))
    
    Call linesColl.Add(MakeCsvLine(Array("1", "priority")))
    Call linesColl.Add(MakeCsvLine(Array("1", "beta")))
    Call linesColl.Add(MakeCsvLine(Array("2", "vip")))
    Call linesColl.Add(MakeCsvLine(Array("3", "internal")))
    Call linesColl.Add(MakeCsvLine(Array("3", "jp")))
    Call linesColl.Add(MakeCsvLine(Array("7", "bulk")))
    Call linesColl.Add(MakeCsvLine(Array("10", "oversea")))
    Call linesColl.Add(MakeCsvLine(Array("12", "promo")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Contacts.csv（1:1想定、欠損混在）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_ContactsCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Email", "Phone")))
    
    Call linesColl.Add(MakeCsvLine(Array("1", "alice@example.com", "+1-202-555-0101")))
    Call linesColl.Add(MakeCsvLine(Array("2", "bob@example.co.uk", "")))
    Call linesColl.Add(MakeCsvLine(Array("3", "", "+81-3-0000-0000")))
    Call linesColl.Add(MakeCsvLine(Array("4", "data.taro@example.jp", "+81-90-1234-5678")))
    Call linesColl.Add(MakeCsvLine(Array("5", "eve@example.de", "")))
    Call linesColl.Add(MakeCsvLine(Array("7", "grace@example.com", "+1-415-555-0199")))
    Call linesColl.Add(MakeCsvLine(Array("10", "judy@example.au", "")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'------------------------------------------------------------
' 処理内容 : Scores.csv（1:1想定、数値/ランク）を作成
' 引数     : ByVal file_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub Write_ScoresCsv( _
    ByVal file_path As String _
)
    Dim linesColl As Collection
    Set linesColl = New Collection
    
    Call linesColl.Add(MakeCsvLine(Array("ParentID", "Quiz", "Rank")))
    
    Call linesColl.Add(MakeCsvLine(Array("1", "78", "B")))
    Call linesColl.Add(MakeCsvLine(Array("2", "91", "A")))
    Call linesColl.Add(MakeCsvLine(Array("3", "66", "C")))
    Call linesColl.Add(MakeCsvLine(Array("4", "88", "B")))
    Call linesColl.Add(MakeCsvLine(Array("7", "97", "A")))
    Call linesColl.Add(MakeCsvLine(Array("12", "59", "D")))
    
    Call SaveLinesToFile(file_path, linesColl)
End Sub

'==================== ヘルパ ====================

'------------------------------------------------------------
' 処理内容 : フォルダパスの末尾\を補正
' 引数     : ByVal folder_path
' 戻り値   : 文字列
'------------------------------------------------------------
Private Function NormalizeFolderPath( _
    ByVal folder_path As String _
) As String
    If Len(folder_path) = 0 Then
        NormalizeFolderPath = folder_path
    ElseIf Right$(folder_path, 1) = "\" Then
        NormalizeFolderPath = folder_path
    Else
        NormalizeFolderPath = folder_path & "\"
    End If
End Function

'------------------------------------------------------------
' 処理内容 : フォルダを作成（存在しなければ）
' 引数     : ByVal folder_path
' 戻り値   : なし
'------------------------------------------------------------
Private Sub EnsureFolder( _
    ByVal folder_path As String _
)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If Not fso.FolderExists(folder_path) Then
        Call fso.CreateFolder(folder_path)
    End If
End Sub

'------------------------------------------------------------
' 処理内容 : 1レコードの配列をCSV行（正しいエスケープ付き）へ変換
' 引数     : ByVal valuesArr  1次元配列(Variant)
' 戻り値   : 文字列（CSVの1行）
'------------------------------------------------------------
Private Function MakeCsvLine( _
    ByVal valuesArr As Variant _
) As String
    Dim i As Long
    Dim partsArr() As String
    ReDim partsArr(0 To UBound(valuesArr) - LBound(valuesArr))
    
    Dim idx As Long
    idx = 0
    
    For i = LBound(valuesArr) To UBound(valuesArr)
        partsArr(idx) = CsvEscape(CStr(valuesArr(i)))
        idx = idx + 1
    Next i
    
    MakeCsvLine = Join(partsArr, ",")
End Function

'------------------------------------------------------------
' 処理内容 : CSV用の値エスケープ（カンマ/ダブルクォート/改行があればクォート）
' 引数     : ByVal s
' 戻り値   : 文字列（必要に応じて""で囲み、内部の""を重ねる）
'------------------------------------------------------------
Private Function CsvEscape( _
    ByVal s As String _
) As String
    Dim needQuote As Boolean
    needQuote = (InStr(s, ",") > 0) _
                Or (InStr(s, """") > 0) _
                Or (InStr(s, vbCr) > 0) _
                Or (InStr(s, vbLf) > 0)
    
    If InStr(s, """") > 0 Then
        s = Replace$(s, """", """""")
    End If
    
    If needQuote Or Len(s) = 0 Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function

'------------------------------------------------------------
' 処理内容 : 行コレクションを書き出し（CRLF区切り）
' 引数     : ByVal file_path, ByVal linesColl
' 戻り値   : なし
'------------------------------------------------------------
Private Sub SaveLinesToFile( _
    ByVal file_path As String, _
    ByVal linesColl As Collection _
)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim ts As Scripting.TextStream
    Set ts = fso.CreateTextFile(file_path, True, False) ' overwrite, ASCII(=False)
    
    Dim i As Long
    For i = 1 To linesColl.Count
        ts.Write linesColl(i)
        ts.Write vbCrLf
    Next i
    
    Call ts.Close
End Sub


