Option Explicit

'--------------------------------------------------
' 機能: "指定拡張子のファイルを再帰的に収集し、ログ出力とコピーを高速に行う"
' 引数: "srcRoot - 探索元ルートフォルダのフルパス; dstRoot - コピー先ルートフォルダのフルパス; fileExts - 拡張子配列(例: Array(\"pdf\",\"txt\")、小文字、ドット無し); overwrite - 既存ファイルを上書きするか(既定 True)"
' 返値: "なし"
'--------------------------------------------------
Public Sub CopyFilesByExt(
    ByVal srcRoot As String,
    ByVal dstRoot As String,
    ByVal fileExts As Variant,
    Optional ByVal overwrite As Boolean = True)

    Dim fso As FileSystemObject
    Dim extDict As Dictionary
    Dim fileList As Collection
    Dim rootLen As Long

    ' パフォーマンス最適化: 画面更新・計算を停止
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    ' FSO と拡張子辞書の初期化
    Set fso = New FileSystemObject
    Set extDict = New Dictionary
    Dim ext As Variant
    For Each ext In fileExts
        extDict.Add ext, True
    Next ext

    ' ルートフォルダ存在チェック
    If Not fso.FolderExists(srcRoot) Then
        Err.Raise vbObjectError + 513, "CopyFilesByExt", _
                  "Source folder not found: " & srcRoot
    End If

    ' 相対パス算出用の長さ
    rootLen = Len(srcRoot) + 2   ' パス区切り文字分

    Set fileList = New Collection
    CollectFilesRec fso, srcRoot, dstRoot, rootLen, extDict, fileList

    LogFileList fileList
    CopyFromList fso, fileList, overwrite

    ' 後処理: パフォーマンス設定復元
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set extDict = Nothing
    Set fso = Nothing
End Sub

'--------------------------------------------------
' 機能: "currentFolder以下を再帰的にスキャンし、対象拡張子のファイル情報を fileList に追加する"
' 引数: "fso - FileSystemObject; currentFolder - 現在のフォルダパス; dstRoot - コピー先ルートフォルダ; rootLen - srcRoot を除去する文字数; extDict - 対象拡張子の辞書; fileList - Collection(要素: Array(sourceFolder, targetFolder, fileName))"
' 返値: "なし"
'--------------------------------------------------
Private Sub CollectFilesRec(
    ByVal fso As FileSystemObject,
    ByVal currentFolder As String,
    ByVal dstRoot As String,
    ByVal rootLen As Long,
    ByVal extDict As Dictionary,
    ByRef fileList As Collection)

    Dim folderItem As Folder
    Set folderItem = fso.GetFolder(currentFolder)

    Dim fileItem As File, subFolder As Folder
    Dim relPath As String, targetFolder As String, extName As String

    ' ファイルのスキャン
    For Each fileItem In folderItem.Files
        extName = LCase$(fso.GetExtensionName(fileItem.Name))
        If extDict.Exists(extName) Then
            relPath = Mid$(folderItem.Path, rootLen)
            If Len(relPath) > 0 Then
                targetFolder = dstRoot & "\" & relPath
            Else
                targetFolder = dstRoot
            End If
            fileList.Add Array(folderItem.Path, targetFolder, fileItem.Name)
        End If
    Next fileItem

    ' サブフォルダの再帰
    For Each subFolder In folderItem.SubFolders
        CollectFilesRec fso, subFolder.Path, dstRoot, rootLen, extDict, fileList
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
    Set ws = ThisWorkbook.Sheets("Log")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Log"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' ヘッダー
    ws.Range("A1:C1").Value = Array("InputFolder", "OutputFolder", "FileName")
    ws.Range("A1:C1").Font.Bold = True

    Dim n As Long: n = fileList.Count
    If n > 0 Then
        Dim dataArr() As String
        ReDim dataArr(1 To n, 1 To 3)
        Dim i As Long, itm As Variant
        For i = 1 To n
            itm = fileList(i)
            dataArr(i, 1) = itm(0)
            dataArr(i, 2) = itm(1)
            dataArr(i, 3) = itm(2)
        Next i
        With ws.Range("A2").Resize(n, 3)
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
    For Each item In fileList
        EnsureFolderExists fso, item(1)
        fso.CopyFile Source:=item(0) & "\" & item(2), _
                     Destination:=item(1) & "\" & item(2), _
                     OverWriteFiles:=overwrite
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
    Dim p As String
    p = fso.GetParentFolderName(folderPath)
    If Len(p) > 0 Then EnsureFolderExists fso, p
    fso.CreateFolder folderPath
End Sub

'--------------------------------------------------
' 機能: "CopyFilesByExt の使用例"
' 引数: "なし"
' 返値: "なし"
'--------------------------------------------------
Private Sub TestCopyByExt()
    Dim exts As Variant
    exts = Array("pdf", "txt", "xlsx")
    CopyFilesByExt "C:\Source", "D:\Dest", exts, True
End Sub
