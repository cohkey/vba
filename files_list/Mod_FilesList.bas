Option Explicit

'--------------------------------------------------
' 機能: "指定拡張子のファイルを再帰的に収集し、ログ出力とコピーを行う"
' 引数: "srcRoot - 探索元ルートフォルダのフルパス; dstRoot - コピー先ルートフォルダのフルパス; fileExts - 拡張子配列(例: Array(\"pdf\",\"txt\")、小文字、ドット無し); overwrite - 既存ファイルを上書きするか(既定 True)"
' 返値: "なし"
'--------------------------------------------------
Public Sub CopyFilesByExt(
    ByVal srcRoot As String,
    ByVal dstRoot As String,
    ByVal fileExts As Variant,
    Optional ByVal overwrite As Boolean = True)

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    If Not fso.FolderExists(srcRoot) Then
        Err.Raise ERR_SOURCE_NOT_FOUND, "CopyFilesByExt", "Source folder not found: " & srcRoot
    End If

    Dim fileList As Collection
    Set fileList = New Collection
    CollectFilesRec fso, srcRoot, dstRoot, srcRoot, fileExts, fileList

    LogFileList fileList
    CopyFromList fso, fileList, overwrite

    Set fso = Nothing
End Sub

'--------------------------------------------------
' 機能: "currentFolder以下を再帰的にスキャンし、対象拡張子のファイル情報を fileList に追加する"
' 引数: "fso - FileSystemObject; currentFolder - 現在のフォルダパス; dstRoot - コピー先ルートフォルダ; srcRoot - 探索元ルートフォルダ; fileExts - 拡張子配列; fileList - Collection(Array(sourceFolder, targetFolder, fileName))"
' 返値: "なし"
'--------------------------------------------------
Private Sub CollectFilesRec(
    ByVal fso As FileSystemObject,
    ByVal currentFolder As String,
    ByVal dstRoot As String,
    ByVal srcRoot As String,
    ByVal fileExts As Variant,
    ByRef fileList As Collection)

    Dim folderItem As Folder
    Set folderItem = fso.GetFolder(currentFolder)

    Dim fileItem As File, subFolder As Folder
    Dim relPath As String, targetFolder As String, ext As String
    Dim i As Long

    For Each fileItem In folderItem.Files
        ext = LCase$(fso.GetExtensionName(fileItem.Name))
        For i = LBound(fileExts) To UBound(fileExts)
            If ext = fileExts(i) Then
                relPath = Replace(folderItem.Path, srcRoot & "\", "")
                If relPath <> "" Then
                    targetFolder = dstRoot & "\" & relPath
                Else
                    targetFolder = dstRoot
                End If
                fileList.Add Array(folderItem.Path, targetFolder, fileItem.Name)
                Exit For
            End If
        Next i
    Next fileItem

    For Each subFolder In folderItem.SubFolders
        CollectFilesRec fso, subFolder.Path, dstRoot, srcRoot, fileExts, fileList
    Next subFolder
End Sub

'--------------------------------------------------
' 機能: "fileList の内容を 'Log' シートにヘッダー行と共に一括出力する(すべて文字列として貼り付け)"
' 引数: "fileList - Collection(Array(sourceFolder, targetFolder, fileName))"
' 返値: "なし"
'--------------------------------------------------
Private Sub LogFileList(
    ByVal fileList As Collection)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Log")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Log"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' ヘッダー出力
    ws.Range("A1:C1").Value = Array("InputFolder", "OutputFolder", "FileName")
    ws.Range("A1:C1").Font.Bold = True

    Dim countRows As Long
    countRows = fileList.Count
    If countRows > 0 Then
        Dim dataArr As Variant
        ReDim dataArr(1 To countRows, 1 To 3)
        Dim i As Long
        For i = 1 To countRows
            Dim item As Variant
            item = fileList(i)
            dataArr(i, 1) = CStr(item(0))
            dataArr(i, 2) = CStr(item(1))
            dataArr(i, 3) = CStr(item(2))
        Next i

        With ws.Range("A2").Resize(countRows, 3)
            .NumberFormat = "@"
            .Value = dataArr
        End With
    End If

    ws.Columns("A:C").AutoFit
End Sub

'--------------------------------------------------
' 機能: "fileList をもとにファイルをコピーし、必要に応じてフォルダを作成する"
' 引数: "fso - FileSystemObject; fileList - Collection(Array(sourceFolder, targetFolder, fileName)); overwrite - 上書き可否"
' 返値: "なし"
'--------------------------------------------------
Private Sub CopyFromList(
    ByVal fso As FileSystemObject,
    ByVal fileList As Collection,
    ByVal overwrite As Boolean)

    Dim item As Variant
    Dim sourceFolder As String, outputFolder As String, fileName As String

    For Each item In fileList
        sourceFolder = item(0)
        outputFolder = item(1)
        fileName = item(2)

        EnsureFolderExists fso, outputFolder
        fso.CopyFile Source:= sourceFolder & "\" & fileName, _
                     Destination:= outputFolder & "\" & fileName, _
                     OverWriteFiles:= overwrite
    Next item
End Sub

'--------------------------------------------------
' 機能: "指定フォルダとその親を再帰的に作成する"
' 引数: "fso - FileSystemObject; folderPath - 作成対象フォルダのフルパス"
' 返値: "なし"
'--------------------------------------------------
Private Sub EnsureFolderExists(
    ByVal fso As FileSystemObject,
    ByVal folderPath As String)

    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> "" Then EnsureFolderExists fso, parentPath
    fso.CreateFolder folderPath
End Sub

'--------------------------------------------------
' 機能: "CopyFilesByExt の使用例"
' 引数: "なし"
' 返値: "なし"
'--------------------------------------------------
Private Sub TestCopyByExt()
    Dim src As String, dst As String, exts As Variant
    src = "C:\Source": dst = "D:\Dest"
    exts = Array("pdf", "txt", "xlsx")
    CopyFilesByExt src, dst, exts, True
End Sub
