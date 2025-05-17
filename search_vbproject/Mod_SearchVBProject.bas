Option Explicit


Sub TestFindInFile()
    Dim foundColl  As Collection
    Dim moduleInfo As Variant

    Set foundColl = FindModulesByKeywordInFile( _
        "C:\path\to\YourMacroBook.xlsm", _
        "検索したいキーワード" _
    )

    If foundColl.Count = 0 Then
        MsgBox "どのモジュールにもキーワードは見つかりませんでした。"
    Else
        For Each moduleInfo In foundColl
            MsgBox "キーワードを含むモジュール: " & moduleInfo
        Next
    End If
End Sub





'--------------------------------------------------
'- 機能: 指定したマクロファイルを開き、
'       全モジュールから指定キーワードを検索し、
'       該当モジュール名＋拡張子を返す
'- 引数: file_path      - マクロファイルのフルパス
'-       search_keyword - 検索する文字列
'- 返値: 検索キーワードを含むモジュール名＋拡張子のCollection
'        （例: "Module1.bas", "MyClass.cls", "UserForm1.frm"）
'--------------------------------------------------
Public Function FindModulesByKeywordInFile( _
    ByVal file_path As String, _
    ByVal search_keyword As String _
) As Collection

    Dim resultsColl    As Collection
    Set resultsColl = New Collection

    Dim targetWb       As Workbook
    Dim vbComp         As VBIDE.VBComponent
    Dim codeMod        As VBIDE.CodeModule
    Dim totalLines     As Long
    Dim moduleText     As String
    Dim ext            As String

    ' 1) マクロファイルを読み取り専用で開く
    Set targetWb = Application.Workbooks.Open( _
        Filename:=file_path, _
        ReadOnly:=True _
    )

    ' 2) 各モジュールをループ
    For Each vbComp In targetWb.VBProject.VBComponents

        ' ── 拡張子セット (.bas/.cls/.frm)
        Select Case vbComp.Type
        Case vbext_ct_StdModule
            ext = ".bas"
        Case vbext_ct_ClassModule
            ext = ".cls"
        Case vbext_ct_MSForm
            ext = ".frm"
        Case Else
            ext = ""
        End Select

        ' ── コード行を取得
        Set codeMod = vbComp.CodeModule
        totalLines = codeMod.CountOfLines

        If totalLines > 0 Then
            moduleText = codeMod.Lines(1, totalLines)
            ' 大文字小文字を無視して検索
            If InStr(1, moduleText, search_keyword, vbTextCompare) > 0 Then
                Call resultsColl.Add(vbComp.Name & ext)
            End If
        End If
    Next vbComp

    ' 3) ブックを閉じる
    Call targetWb.Close(SaveChanges:=False)

    Set FindModulesByKeywordInFile = resultsColl
End Function


