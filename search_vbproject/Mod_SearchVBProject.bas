Option Explicit

Private Sub TestSearchKeywordInModules()
    Dim exportFolder As String, keyword As String
    exportFolder = "C:\path\to\ExportedModules"
    keyword = "Keyword"

    ' モジュールを先にエクスポート
    Call ExportAllModulesFromFolder("C:\path\to\MacroBooks", exportFolder)
    ' エクスポート済みモジュールを検索
    Call SearchExportedModulesAndSummary(exportFolder, keyword, True)
End Sub



'--------------------------------------------------
' Sub  : ExportAllModulesFromFolder
' 機能  : 指定フォルダ内のマクロブックを開き、すべてのモジュールをエクスポート
' 引数  : ByVal source_folder As String - マクロブック格納フォルダ
'        : ByVal export_folder As String - モジュール出力先基底フォルダ
' 戻値  : なし
'--------------------------------------------------
Public Sub ExportAllModulesFromFolder(ByVal source_folder As String, ByVal export_folder As String)
    On Error GoTo ErrHandler
    Dim fso As Object
    Dim folderItem As Object
    Dim fileItem As Object
    Dim wbPath As String
    Dim targetWb As Workbook
    Dim comp As VBIDE.VBComponent
    Dim baseName As String
    Dim outFolder As String
    Dim ext As String

    ' FileSystemObject を作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' パス末尾の"\"を保証 (プラットフォーム依存文字を使用)
    If Right$(source_folder, 1) <> Application.PathSeparator Then source_folder = source_folder & Application.PathSeparator
    If Right$(export_folder, 1) <> Application.PathSeparator Then export_folder = export_folder & Application.PathSeparator

    ' エクスポート先フォルダ作成（なければ）
    If Not fso.FolderExists(export_folder) Then fso.CreateFolder export_folder

    ' 既存サブフォルダがあればエラー終了
    If fso.GetFolder(export_folder).SubFolders.Count > 0 Then
        MsgBox "エクスポート先フォルダに既存のサブフォルダが存在します。処理を中止します。", vbCritical
        GoTo CleanExit
    End If

    ' 指定フォルダ内のファイルをループ
    Set folderItem = fso.GetFolder(source_folder)
    For Each fileItem In folderItem.Files
        Select Case LCase(fso.GetExtensionName(fileItem.Name))
            Case "xls", "xlsx", "xlsm"
                wbPath = fileItem.Path
                Set targetWb = Application.Workbooks.Open(fileName:=wbPath, ReadOnly:=True)

                ' サブフォルダ名＝ブック名（拡張子除く）
                baseName = Left$(fileItem.Name, InStrRev(fileItem.Name, ".") - 1)
                outFolder = export_folder & baseName & Application.PathSeparator
                If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder

                ' 各コンポーネントをエクスポート
                For Each comp In targetWb.VBProject.VBComponents
                    Select Case comp.Type
                        Case vbext_ct_StdModule:     ext = ".bas"
                        Case vbext_ct_ClassModule:   ext = ".cls"
                        Case vbext_ct_MSForm:        ext = ".frm"
                        Case Else:                   ext = ""
                    End Select
                    If ext <> "" Then comp.Export outFolder & comp.Name & ext
                Next comp

                targetWb.Close SaveChanges:=False
        End Select
    Next fileItem

    MsgBox "モジュールのエクスポートが完了しました。", vbInformation

CleanExit:
    Exit Sub

ErrHandler:
    MsgBox "[ExportAllModulesFromFolder] エラー " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'--------------------------------------------------
' Sub  : SearchExportedModulesAndSummary
' 機能  : エクスポート済みモジュールを対象にキーワード検索し
'         ・1行目に検索文字列と実行日時を出力
'         ・2行目にヘッダー、3行目以降に検索結果を出力
' 引数  : ByVal export_folder   - モジュール出力基底フォルダ
'        : ByVal search_keyword  - 検索文字列
'        : Optional ByVal case_sensitive As Boolean = False
' 戻値  : なし
'--------------------------------------------------
Public Sub SearchExportedModulesAndSummary(ByVal export_folder As String, ByVal search_keyword As String, Optional ByVal case_sensitive As Boolean = False)
    On Error GoTo ErrHandler
    Dim fso As Object
    Dim rootFolder As Object
    Dim folderItem As Object
    Dim fileItem As Object
    Dim wsDetail As Worksheet
    Dim wsSummary As Worksheet
    Dim summaryColl As Collection: Set summaryColl = New Collection
    Dim ts As Object
    Dim lineText As String
    Dim lineNum As Long
    Dim compareMode As VbCompareMethod
    Dim currentProc As String
    Dim currentRow As Long
    Dim hitCount As Long
    Dim currentTs As String
    Dim fileExt As String
    Dim procPattern As Object

    ' 比較モード設定
    If case_sensitive Then
        compareMode = vbBinaryCompare
    Else
        compareMode = vbTextCompare
    End If

    ' FileSystemObject 作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' パス末尾の"\"を保証
    If Right$(export_folder, 1) <> Application.PathSeparator Then export_folder = export_folder & Application.PathSeparator
    Set rootFolder = fso.GetFolder(export_folder)

    ' 正規表現パターン準備
    Set procPattern = CreateObject("VBScript.RegExp")
    With procPattern
        .Pattern = "^(?:Public|Private)?\s*(?:Sub|Function)\s+(\w+)"
        .IgnoreCase = True
    End With

    ' シート準備
    On Error Resume Next
    Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("SearchResults").Delete
        ThisWorkbook.Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    ' Detail シート作成
    Set wsDetail = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDetail.Name = "SearchResults"

    ' 1行目に検索文字列と実行日時
    currentTs = Format(Now, "yyyy-mm-dd hh:nn:ss")
    wsDetail.Cells(1, 1).Value = "検索文字列: " & search_keyword & "    実行日時: " & currentTs

    ' 2行目にヘッダー
    wsDetail.Range("A2:D2").Value = Array("モジュールファイル", "プロシージャ名", "行番号", "コード内容")
    currentRow = 3

    ' 各サブフォルダ（ブック名）を走査
    For Each folderItem In rootFolder.SubFolders
        For Each fileItem In folderItem.Files
            fileExt = LCase(fso.GetExtensionName(fileItem.Name))
            If fileExt = "bas" Or fileExt = "cls" Or fileExt = "frm" Then
                hitCount = 0
                Set ts = fso.OpenTextFile(fileItem.Path, 1)
                lineNum = 0
                currentProc = "(モジュールレベル)"
                Do While Not ts.AtEndOfStream
                    lineText = ts.ReadLine
                    lineNum = lineNum + 1
                    ' プロシージャ名を検出して保持
                    If procPattern.Test(lineText) Then
                        currentProc = procPattern.Execute(lineText)(0).SubMatches(0)
                    End If
                    ' キーワード検索
                    If InStr(1, lineText, search_keyword, compareMode) > 0 Then
                        wsDetail.Cells(currentRow, 1).Value = folderItem.Name & ":" & fileItem.Name
                        wsDetail.Cells(currentRow, 2).Value = currentProc
                        wsDetail.Cells(currentRow, 3).Value = lineNum
                        wsDetail.Cells(currentRow, 4).Value = Trim$(lineText)
                        currentRow = currentRow + 1
                        hitCount = hitCount + 1
                    End If
                Loop
                ts.Close
                ' Summary 登録
                summaryColl.Add Array(folderItem.Name, fileItem.Name, hitCount, "")
            End If
        Next fileItem
    Next folderItem

    ' Summary シート作成
    Set wsSummary = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsSummary.Name = "Summary"
    wsSummary.Range("A1:D1").Value = Array("ブック名", "モジュールファイル", "ヒット件数", "エラー内容")
    Dim i As Long
    For i = 1 To summaryColl.Count
        wsSummary.Cells(i + 1, 1).Resize(1, 4).Value = summaryColl(i)
    Next i
    wsSummary.Columns("A:D").AutoFit
    wsDetail.Columns("A:D").AutoFit

    MsgBox "検索結果の出力が完了しました。", vbInformation

CleanExit:
    Exit Sub

ErrHandler:
    MsgBox "[SearchExportedModulesAndSummary] エラー " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'--------------------------------------------------
' Function: ExtractProcedureName
' （使用せず ProcPattern で行内検出に統一）
'--------------------------------------------------
Private Function ExtractProcedureName(ByVal lineText As String) As String
    ExtractProcedureName = ""
End Function

