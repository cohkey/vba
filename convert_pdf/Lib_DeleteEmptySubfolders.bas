Option Explicit

'============================================================
' 処理内容 : 指定フォルダ直下の各サブフォルダを再帰的に調査し、
'          配下にファイルが1つも無いフォルダ階層のみ削除する
' 引数     : basePath  対象の親フォルダ（このフォルダ自体は削除しない）
' 引数     : dryRun    True=削除せず件数だけ数える / False=実際に削除
' 戻り値   : 削除した（または削除予定の）フォルダ数
' 仕様     : サブフォルダ配下のどこかに1つでもファイルがあれば、その階層は残す
' 例外     : フォルダ未存在など致命的エラー時は Err.Raise 999 で停止
'============================================================
Public Function DeleteEmptySubfolders( _
    ByVal basePath As String, _
    ByVal dryRun As Boolean _
) As Long
    Const ERR_NUM As Long = 999

    Dim fso As Scripting.FileSystemObject
    Dim baseFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    Dim deletedCount As Long

    Set fso = New Scripting.FileSystemObject

    If fso.FolderExists(basePath) = False Then
        Err.Raise ERR_NUM, , "Folder not found: " & basePath
    End If

    Set baseFolder = fso.GetFolder(basePath)

    ' 親フォルダ自身は削除対象外。直下サブフォルダを個別に枝刈り。
    For Each subFolder In baseFolder.SubFolders
        deletedCount = deletedCount + _
                       PruneFolderIfEmpty(subFolder, fso, dryRun)
    Next subFolder

    DeleteEmptySubfolders = deletedCount
End Function

'------------------------------------------------------------
' 処理内容 : フォルダ配下を再帰的に調査し、配下にファイルが無ければ
'          「このフォルダ自体」を削除する（ポストオーダー）
' 引数     : targetFolder 対象サブフォルダ
' 引数     : fso          共有のFSOインスタンス
' 引数     : dryRun       True=削除せずカウントのみ
' 戻り値   : 実際に削除した（またはDryRunで削除予定の）フォルダ数
' 注意     : 配下に1つでもFileがあれば、当該フォルダは残す
'------------------------------------------------------------
Private Function PruneFolderIfEmpty( _
    ByVal targetFolder As Scripting.Folder, _
    ByVal fso As Scripting.FileSystemObject, _
    ByVal dryRun As Boolean _
) As Long
    Dim child As Scripting.Folder
    Dim deletedCount As Long

    ' 1) まず子から再帰的に枝刈り（ポストオーダー）
    For Each child In targetFolder.SubFolders
        deletedCount = deletedCount + _
                       PruneFolderIfEmpty(child, fso, dryRun)
    Next child

    ' 2) 自分の直下にファイルがあるか判定（ここまでで空の子は消えている）
    If targetFolder.Files.Count = 0 And targetFolder.SubFolders.Count = 0 Then
        ' 配下にファイルが一切無い＝空ツリー → 自分を削除
        If dryRun = False Then
            ' 第2引数 True は読み取り専用でも強制削除のため
            fso.DeleteFolder targetFolder.Path, True
        End If
        deletedCount = deletedCount + 1
    End If

    PruneFolderIfEmpty = deletedCount
End Function

'============================================================
' 使い方（テスト用）
'============================================================
Public Sub Run_DeleteEmptySubfolders_Demo()
    Dim target As String
    Dim cnt As Long

    target = "C:\Temp\Work"  ' ← 対象フォルダに変更

    ' まずはドライランで安全確認
    cnt = DeleteEmptySubfolders(target, True)
    Debug.Print "[DRY-RUN] Would delete folders: "; cnt

    ' 問題なければ実行
    cnt = DeleteEmptySubfolders(target, False)
    Debug.Print "[EXEC] Deleted folders: "; cnt
End Sub
