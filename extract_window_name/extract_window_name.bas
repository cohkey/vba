Option Explicit


Function ExtractEdgeTitle(ByVal windowTitle As String) As String
    Dim posYobi As Long, posShokuba As Long, posCut As Long

    ' 右端から「および他」の開始位置を取得
    posYobi = InStrRev(windowTitle, "および他")
    ' 右端から「- 職場」の開始位置を取得
    posShokuba = InStrRev(windowTitle, "- 職場")

    ' 両方のマーカーが見つからない場合は、そのまま返す
    If posYobi = 0 And posShokuba = 0 Then
        ExtractEdgeTitle = windowTitle
        Exit Function
    End If

    ' 両方見つかった場合、より左側（小さい位置）の方を採用
    If posYobi > 0 And posShokuba > 0 Then
        posCut = IIf(posYobi < posShokuba, posYobi, posShokuba)
    ElseIf posYobi > 0 Then
        posCut = posYobi
    Else
        posCut = posShokuba
    End If

    ' 先頭から posCut - 1 までを抽出し、余計な空白を除去
    ExtractEdgeTitle = Trim(Left(windowTitle, posCut - 1))
End Function

Sub TestExtractEdgeTitle()
    Dim s1 As String, s2 As String, s3 As String, s4 As String

    s1 = "EdgeOpen - タイトル および他１ページ - 職場 - Microsoft Edge"
    s2 = "EdgeOpen - タイトル および他３ページ - 職場 - Microsoft Edge"
    s3 = "EdgeOpen - タイトル - 職場 - Microsoft Edge"
    s4 = "EdgeOpen - タイトル１ - タイトル２ および他３ページ - 職場 - Microsoft Edge"

    Debug.Print ExtractEdgeTitle(s1) ' 出力例: "EdgeOpen - タイトル"
    Debug.Print ExtractEdgeTitle(s2) ' 出力例: "EdgeOpen - タイトル"
    Debug.Print ExtractEdgeTitle(s3) ' 出力例: "EdgeOpen - タイトル"
    Debug.Print ExtractEdgeTitle(s4) ' 出力例: "EdgeOpen - タイトル１ - タイトル２"
End Sub

