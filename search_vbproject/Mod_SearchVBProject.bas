Option Explicit

Private Sub TestSearchKeywordInModules()
    Dim folderPath As String, keyword As String
    folderPath = "C:\path\to\YourMacroBookFolder"
    keyword = "Keyword"

'    Call ExportSearchResultsToSheet(wbPath, keyword)
    Call ExportFolderSearchAndSummary(folderPath, keyword, True)
End Sub


'--------------------------------------------------
' 指定フォルダ内のマクロファイルを
' ・SearchResults シートに行単位で出力
' ・Summary シートにファイル/モジュール単位でヒット数を集計
'
' 引数:
'   folder_path    : 対象フォルダのパス
'   search_keyword : 検索文字列
'   case_sensitive : True=大文字小文字区別, False=非区別 (省略可)
'--------------------------------------------------
Public Sub ExportFolderSearchAndSummary( _
    ByVal folder_path As String, _
    ByVal search_keyword As String, _
    Optional ByVal case_sensitive As Boolean = False _
)
    Dim wsDetail    As Worksheet
    Dim wsSummary   As Worksheet
    Dim fileName    As String
    Dim filePath    As String
    Dim foundColl   As Collection
    Dim summaryColl As Collection: Set summaryColl = New Collection
    Dim currentRow  As Long
    Dim currentTs   As String

    '── SearchResults シート準備 ──
    On Error Resume Next
    Application.DisplayAlerts = False
      ThisWorkbook.Worksheets("SearchResults").Delete
      ThisWorkbook.Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' ← After にはオブジェクトを渡す
    Set wsDetail = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) _
    )
    wsDetail.Name = "SearchResults"
    wsDetail.Range("A1:K1").Value = Array( _
        "検索キーワード", "結果", "エラー内容", "タイムスタンプ", _
        "ファイルパス", "ファイル名", "モジュール名", _
        "モジュール種類", "プロシージャ名", "行番号", "コード内容" _
    )
    currentRow = 2

    '── フォルダ内ループ ──
    fileName = Dir(folder_path & "\*.*")
    Do While fileName <> ""
        filePath = folder_path & "\" & fileName

        If LCase$(Right$(fileName, 5)) = ".xlsm" _
        Or LCase$(Right$(fileName, 5)) = ".xlsb" _
        Or LCase$(Right$(fileName, 4)) = ".xls" _
        Or LCase$(Right$(fileName, 5)) = ".xlam" Then

            currentTs = Format(Now, "yyyy-mm-dd hh:nn:ss")
            On Error Resume Next
            Set foundColl = FindModulesByKeywordInFile( _
                filePath, search_keyword, case_sensitive _
            )
            If Err.Number <> 0 Then
                With wsDetail
                    .Cells(currentRow, 1).Value = search_keyword
                    .Cells(currentRow, 2).Value = "Error"
                    .Cells(currentRow, 3).Value = Err.Description
                    .Cells(currentRow, 4).Value = currentTs
                    .Cells(currentRow, 5).Value = filePath
                    .Cells(currentRow, 6).Value = fileName
                End With
                summaryColl.Add Array(filePath, fileName, "", 0, Err.Description)
                Err.Clear
                currentRow = currentRow + 1
            Else
                ' 全モジュール名の取得
                Dim tmpWb      As Workbook
                Dim comp       As VBIDE.VBComponent
                Dim allModules As Collection: Set allModules = New Collection
                Dim extName    As String

                Set tmpWb = Application.Workbooks.Open( _
                    fileName:=filePath, ReadOnly:=True _
                )
                For Each comp In tmpWb.VBProject.VBComponents
                    Select Case comp.Type
                    Case vbext_ct_StdModule:   extName = ".bas"
                    Case vbext_ct_ClassModule: extName = ".cls"
                    Case vbext_ct_MSForm:      extName = ".frm"
                    Case Else:                 extName = ""
                    End Select
                    If extName <> "" Then allModules.Add comp.Name & extName
                Next
                tmpWb.Close SaveChanges:=False

                ' Detail 出力
                If foundColl.Count > 0 Then
                    Dim resultItem As Variant
                    For Each resultItem In foundColl
                        Dim moduleType As String
                        moduleType = Mid$( _
                            resultItem(1), InStrRev(resultItem(1), ".") + 1 _
                        )
                        With wsDetail
                            .Cells(currentRow, 1).Value = search_keyword
                            .Cells(currentRow, 2).Value = "OK"
                            .Cells(currentRow, 3).Value = ""
                            .Cells(currentRow, 4).Value = currentTs
                            .Cells(currentRow, 5).Value = resultItem(0)
                            .Cells(currentRow, 6).Value = fileName
                            .Cells(currentRow, 7).Value = resultItem(1)
                            .Cells(currentRow, 8).Value = moduleType
                            .Cells(currentRow, 9).Value = resultItem(2)
                            .Cells(currentRow, 10).Value = resultItem(3)
                            .Cells(currentRow, 11).Value = resultItem(4)
                        End With
                        currentRow = currentRow + 1
                    Next
                Else
                    With wsDetail
                        .Cells(currentRow, 1).Value = search_keyword
                        .Cells(currentRow, 2).Value = "OK"
                        .Cells(currentRow, 3).Value = "No hits"
                        .Cells(currentRow, 4).Value = currentTs
                        .Cells(currentRow, 5).Value = filePath
                        .Cells(currentRow, 6).Value = fileName
                    End With
                    currentRow = currentRow + 1
                End If

                ' Summary 用ヒット数集計
                Dim dictHits As Object: Set dictHits = CreateObject("Scripting.Dictionary")
                Dim mName As Variant
                For Each mName In allModules: dictHits(mName) = 0: Next
                For Each resultItem In foundColl
                    dictHits(resultItem(1)) = dictHits(resultItem(1)) + 1
                Next
                For Each mName In dictHits.Keys
                    summaryColl.Add Array(filePath, fileName, mName, dictHits(mName), "")
                Next
            End If
            On Error GoTo 0
        End If
        fileName = Dir()
    Loop

    '── Summary シート出力 ──
    Set wsSummary = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) _
    )
    wsSummary.Name = "Summary"
    wsSummary.Range("A1:E1").Value = Array( _
        "ファイルパス", "ファイル名", "モジュール名", "ヒット件数", "エラー内容" _
    )
    Dim i As Long
    For i = 1 To summaryColl.Count
        wsSummary.Cells(i + 1, 1).Resize(1, 5).Value = summaryColl(i)
    Next
    wsSummary.Columns("A:E").AutoFit
    wsDetail.Columns("A:K").AutoFit

    MsgBox "SearchResults と Summary の出力が完了しました。", vbInformation
End Sub






'--------------------------------------------------
' 機能  : 検索結果を二次元配列にまとめ、
'         新規シートへヘッダー付きで貼り付ける
' 引数  : ByVal file_path      - マクロファイルのフルパス
'         ByVal search_keyword - 検索する文字列
' 戻値  : なし
'--------------------------------------------------
Public Sub ExportSearchResultsToSheet( _
    ByVal file_path As String, _
    ByVal search_keyword As String _
)

    Dim foundColl     As Collection
    Dim resultItem    As Variant
    Dim resultArr     As Variant
    Dim i             As Long
    Dim rowCount      As Long
    Dim colCount      As Long
    Dim ws            As Worksheet

    ' 検索実行
    Set foundColl = FindModulesByKeywordInFile( _
        file_path, _
        search_keyword _
    )

    rowCount = foundColl.Count

    If rowCount = 0 Then
        MsgBox "キーワード「" & search_keyword & "」は見つかりませんでした。"
        Exit Sub
    End If

    ' 生成する配列の行数＝ヒット件数＋1（ヘッダー用）、列数＝5
    colCount = 5
    ReDim resultArr(1 To rowCount + 1, 1 To colCount)

    ' ヘッダーをセット
    resultArr(1, 1) = "ファイルパス"
    resultArr(1, 2) = "モジュール名"
    resultArr(1, 3) = "プロシージャ名"
    resultArr(1, 4) = "行番号"
    resultArr(1, 5) = "コード内容"

    ' データ行をセット
    For i = 1 To rowCount
        resultItem = foundColl(i)
        resultArr(i + 1, 1) = resultItem(0)
        resultArr(i + 1, 2) = resultItem(1)
        resultArr(i + 1, 3) = resultItem(2)
        resultArr(i + 1, 4) = resultItem(3)
        resultArr(i + 1, 5) = resultItem(4)
    Next i

    ' 新規シート作成（既存なら削除してから）
    On Error Resume Next
    Application.DisplayAlerts = False
    Call ThisWorkbook.Worksheets("SearchResults").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count) _
    )
    ws.Name = "SearchResults"

    ' 配列をシートに貼り付け
    ws.Range(ws.Cells(1, 1), _
        ws.Cells(rowCount + 1, colCount) _
    ).Value = resultArr

    ' 列幅を自動調整
    ws.Columns("A:E").AutoFit

    MsgBox "シート 'SearchResults' に " & rowCount & " 件を出力しました。"

End Sub



'--------------------------------------------------
' 機能  : 指定マクロファイル内の全モジュールを行単位で検索し、
'         ヒット箇所を返す
' 引数  : ByVal file_path      - マクロファイルのフルパス
'         ByVal search_keyword - 検索する文字列
'         Optional ByVal case_sensitive As Boolean = False
'               True ＝大文字小文字を区別
'               False＝区別しない（既定）
' 返値  : Collection
'         各要素は Variant 配列:
'         (0)=ファイルパス
'         (1)=モジュール名＋拡張子
'         (2)=プロシージャ名 または "(モジュールレベル)"
'         (3)=行番号
'         (4)=該当行のコード
'--------------------------------------------------
Public Function FindModulesByKeywordInFile( _
    ByVal file_path As String, _
    ByVal search_keyword As String, _
    Optional ByVal case_sensitive As Boolean = False _
) As Collection

    Dim resultsColl   As Collection: Set resultsColl = New Collection
    Dim targetWb      As Workbook
    Dim vbComp        As VBIDE.VBComponent
    Dim codeMod       As VBIDE.CodeModule
    Dim totalLines    As Long
    Dim ext           As String
    Dim lineIndex     As Long
    Dim lineText      As String
    Dim procName      As String
    Dim procKind      As VBIDE.vbext_ProcKind
    Dim compareMode   As VbCompareMethod

    ' 比較モード設定
    If case_sensitive Then
        compareMode = vbBinaryCompare
    Else
        compareMode = vbTextCompare
    End If

    ' 読み取り専用で開く
    Set targetWb = Application.Workbooks.Open( _
        fileName:=file_path, _
        ReadOnly:=True _
    )

    For Each vbComp In targetWb.VBProject.VBComponents
        Select Case vbComp.Type
        Case vbext_ct_StdModule:   ext = ".bas"
        Case vbext_ct_ClassModule: ext = ".cls"
        Case vbext_ct_MSForm:      ext = ".frm"
        Case Else:                 ext = ""
        End Select

        If ext <> "" Then
            Set codeMod = vbComp.CodeModule
            totalLines = codeMod.CountOfLines

            For lineIndex = 1 To totalLines
                lineText = codeMod.Lines(lineIndex, 1)
                If InStr(1, lineText, search_keyword, compareMode) > 0 Then
                    ' プロシージャ名取得（モジュールレベルは空になる）
                    On Error Resume Next
                    procName = codeMod.ProcOfLine(lineIndex, procKind)
                    On Error GoTo 0
                    If procName = "" Then procName = "(モジュールレベル)"

                    resultsColl.Add Array( _
                        file_path, _
                        vbComp.Name & ext, _
                        procName, _
                        lineIndex, _
                        Trim$(lineText) _
                    )
                End If
            Next
        End If
    Next

    targetWb.Close SaveChanges:=False
    Set FindModulesByKeywordInFile = resultsColl
End Function


