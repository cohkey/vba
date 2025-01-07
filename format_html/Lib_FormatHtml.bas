Option Explicit

' テスト実行
Sub TestFormatHtmlFile()
    Call FormatHtmlFile("path\to\your.html")
End Sub


' ===========================
' 1) HTMLファイルを読み込む関数
'   - 概要: ADODB.Streamを使って指定パスのHTMLファイルを読み込み、文字列を返す
'   - 引数:
'       ByVal file_path As String : 読み込むHTMLファイルのパス
'   - 戻り値:
'       String : ファイルの中身を文字列として返す
' ===========================
Private Function LoadHtmlFile(ByVal file_path As String) As String

    Dim streamObj As Object
    Set streamObj = CreateObject("ADODB.Stream")

    ' ストリームをバイナリモードで開く
    streamObj.Type = 1  ' adTypeBinary
    streamObj.Open
    streamObj.LoadFromFile file_path

    ' バイナリ → テキスト変換
    streamObj.Type = 2  ' adTypeText
    streamObj.Charset = "utf-8"  ' 必要に応じて変更
    streamObj.Position = 0

    LoadHtmlFile = streamObj.ReadText(-1)

    streamObj.Close
    Set streamObj = Nothing

End Function

' ===========================
' 2) 改行や余分な空白を削除する関数
'   - 概要: 文字列から改行(\r,\n)を除去し、連続スペースを1つにまとめる
'   - 引数:
'       ByVal html_str As String : 対象のHTML文字列
'   - 戻り値:
'       String : 不要な改行や空白を取り除いたHTML文字列
' ===========================
Private Function RemoveLineBreaksAndExtraSpaces(ByVal html_str As String) As String

    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")

    ' 改行を削除
    regExObj.pattern = "[\r\n]+"
    regExObj.Global = True
    Dim tempStr As String
    tempStr = regExObj.Replace(html_str, "")

    ' 連続空白を単一空白に
    regExObj.pattern = "\s+"
    regExObj.Global = True
    RemoveLineBreaksAndExtraSpaces = regExObj.Replace(tempStr, " ")

    Set regExObj = Nothing

End Function

' ===========================
' 3) 自己完結タグ判定
'   - 概要: タグ名から自己完結タグかどうか判定する
'   - 引数:
'       ByVal tag_name_str As String : タグ名
'   - 戻り値:
'       Boolean : 自己完結タグならTrue、そうでなければFalse
' ===========================
Private Function IsSelfClosingTag(ByVal tag_name_str As String) As Boolean

    Dim selfClosingTagsArr As Variant
    ' ここで「!doctype」「base」を追加
    selfClosingTagsArr = Array("img", "input", "br", "hr", "meta", _
                               "link", "!doctype", "base")

    Dim t As Variant
    For Each t In selfClosingTagsArr
        If LCase(tag_name_str) = t Then
            IsSelfClosingTag = True
            Exit Function
        End If
    Next t

    IsSelfClosingTag = False

End Function

' ===========================
' 4) コメント判定
'   - 概要: トークンがコメント(<!-- -->)かどうかを判定する
'   - 引数:
'       ByVal token_str As String : 判定対象文字列
'   - 戻り値:
'       Boolean : コメントならTrue、そうでなければFalse
' ===========================
Private Function IsCommentToken(ByVal token_str As String) As Boolean

    Dim tmpStr As String
    tmpStr = LCase(Trim(token_str))

    If Left(tmpStr, 4) = "<!--" And Right(tmpStr, 3) = "-->" Then
        IsCommentToken = True
    Else
        IsCommentToken = False
    End If

End Function

' ===========================
' 5) タグ名のみ小文字に書き換える関数
'   - 概要: <HTML class="x"> を <html class="x"> のように、タグ名だけ小文字化する
'   - 引数:
'       ByVal original_tag As String : 小文字化したいタグ文字列(例: <HTML>, </DIV>, <IMG src="...")
'   - 戻り値:
'       String : タグ名だけ小文字になったタグ文字列
' ===========================
Private Function RewriteTagNameToLowerCase(ByVal original_tag As String) As String

    Dim tagBodyStr As String
    Dim trimmedTag As String
    trimmedTag = Trim(original_tag)

    ' 角括弧を除去
    tagBodyStr = Replace(Replace(trimmedTag, "<", ""), ">", "")
    tagBodyStr = Trim(tagBodyStr)

    ' 終了タグかどうかチェック
    Dim isEnd As Boolean
    isEnd = False

    If Left(tagBodyStr, 1) = "/" Then
        isEnd = True
        tagBodyStr = Mid(tagBodyStr, 2) ' 先頭の"/"を取り除く
        tagBodyStr = Trim(tagBodyStr)
    End If

    ' タグ名を抜き出す（最初の空白まで）
    Dim spacePos As Long
    spacePos = InStr(tagBodyStr, " ")

    Dim tagNameStr As String
    Dim attributesStr As String

    If spacePos > 0 Then
        tagNameStr = Left(tagBodyStr, spacePos - 1)
        attributesStr = Mid(tagBodyStr, spacePos)
    Else
        ' 属性が無ければ全体がタグ名
        tagNameStr = tagBodyStr
        attributesStr = ""
    End If

    ' タグ名だけを小文字化
    Dim lowerTagNameStr As String
    lowerTagNameStr = LCase(tagNameStr)

    ' 再構築
    Dim newTagBodyStr As String
    If isEnd Then
        newTagBodyStr = "/" & lowerTagNameStr & attributesStr
    Else
        newTagBodyStr = lowerTagNameStr & attributesStr
    End If

    newTagBodyStr = Trim(newTagBodyStr)

    ' 角括弧を復元
    RewriteTagNameToLowerCase = "<" & newTagBodyStr & ">"

End Function

' ===========================
' 6) HTMLを整形する関数
'   - 概要: HTML文字列を正規表現でタグ/コメントに分割し、インデントを付けて整形する
'   - 引数:
'       ByVal html_str As String : 整形対象のHTML
'   - 戻り値:
'       String : インデント付きで整形されたHTML文字列
' ===========================
Private Function FormatHtml(ByVal html_str As String) As String

    ' コメントブロック と 通常のタグ をどちらも拾う正規表現
    Dim patternStr As String
    patternStr = "(<!--.*?-->)|(<[^>]+>)"

    Dim regExObj As Object
    Set regExObj = CreateObject("VBScript.RegExp")
    regExObj.pattern = patternStr
    regExObj.Global = True
    regExObj.IgnoreCase = True
    regExObj.MultiLine = True

    Dim tokensArr As Variant
    tokensArr = SplitWithRegex(html_str, regExObj)

    Dim resultStr As String
    resultStr = ""

    Dim indentLevel As Long
    indentLevel = 0

    Dim i As Long
    For i = LBound(tokensArr) To UBound(tokensArr)

        Dim currentTokenStr As String
        currentTokenStr = tokensArr(i)

        ' ------------------------------------------------
        ' 「タグ or コメント」かどうかの判定
        ' ------------------------------------------------
        If IsTag(currentTokenStr) Then

            ' コメントかどうか
            If IsCommentToken(currentTokenStr) Then
                ' コメントならインデント変更せず、そのまま出力
                resultStr = resultStr & _
                            GetIndent(indentLevel) & currentTokenStr & vbCrLf
            Else
                ' 通常のタグ → タグ名だけ小文字へ変換
                currentTokenStr = RewriteTagNameToLowerCase(currentTokenStr)

                Dim tagNameStr As String
                Dim isEndTag As Boolean

                ' タグ名取得
                tagNameStr = GetTagName(currentTokenStr, isEndTag)

                If isEndTag Then
                    ' 終了タグ
                    indentLevel = indentLevel - 1
                    resultStr = resultStr & _
                                GetIndent(indentLevel) & currentTokenStr & vbCrLf
                ElseIf IsSelfClosingTag(tagNameStr) Then
                    ' 自己完結タグ
                    resultStr = resultStr & _
                                GetIndent(indentLevel) & currentTokenStr & vbCrLf
                Else
                    ' 開始タグ
                    resultStr = resultStr & _
                                GetIndent(indentLevel) & currentTokenStr & vbCrLf
                    indentLevel = indentLevel + 1
                End If
            End If

        Else
            ' ------------------------------------------------
            ' テキスト部分の場合
            ' ------------------------------------------------
            Dim trimmedTextStr As String
            trimmedTextStr = Trim(currentTokenStr)

            If trimmedTextStr <> "" Then
                resultStr = resultStr & _
                            GetIndent(indentLevel) & trimmedTextStr & vbCrLf
            End If
        End If

    Next i

    Set regExObj = Nothing
    FormatHtml = resultStr

End Function

' ===========================
' (補助) 正規表現でタグ/コメントを抽出し、配列に格納する関数
'   - 概要: 指定した正規表現でタグやコメントを抽出し、
'           それ以外のテキストと交互に配列に格納する
'   - 引数:
'       ByVal target_str As String : 処理対象文字列
'       ByVal regEx_obj As Object  : 事前設定したRegExpオブジェクト
'   - 戻り値:
'       Variant : 結果を格納したVariant配列
' ===========================
Private Function SplitWithRegex(ByVal target_str As String, _
                               ByVal regEx_obj As Object) As Variant

    Dim resultArr() As Variant
    ReDim resultArr(0)

    Dim matchesObj As Object
    Set matchesObj = regEx_obj.Execute(target_str)

    Dim startPos As Long
    startPos = 1

    Dim matchItemObj As Object
    Dim idx As Long
    idx = 0

    For Each matchItemObj In matchesObj

        ' マッチ開始位置 (VBAのMid関数は1-basedなので+1調整)
        Dim matchStart As Long
        matchStart = matchItemObj.FirstIndex + 1

        ' マッチ前のテキスト部分
        Dim preTextStr As String
        preTextStr = Mid(target_str, startPos, matchStart - startPos)

        If Trim(preTextStr) <> "" Then
            resultArr(idx) = preTextStr
            idx = idx + 1
            ReDim Preserve resultArr(idx)
        End If

        ' マッチした文字列 (タグ or コメント)
        resultArr(idx) = matchItemObj.Value
        idx = idx + 1
        ReDim Preserve resultArr(idx)

        ' 次に探索開始する位置
        startPos = matchStart + Len(matchItemObj.Value)
    Next matchItemObj

    ' 最後に残ったテキスト
    If startPos <= Len(target_str) Then
        Dim tailTextStr As String
        tailTextStr = Mid(target_str, startPos)
        If Trim(tailTextStr) <> "" Then
            resultArr(idx) = tailTextStr
        End If
    End If

    SplitWithRegex = resultArr

End Function

' ===========================
' (補助) タグ判定
'   - 概要: 先頭文字が"<"であり、末尾文字が">"ならタグとみなす
'   - 引数:
'       ByVal str_data As String : 判定対象文字列
'   - 戻り値:
'       Boolean : タグならTrue、そうでなければFalse
' ===========================
Private Function IsTag(ByVal str_data As String) As Boolean

    Dim trimmedStr As String
    trimmedStr = Trim(str_data)

    If Left(trimmedStr, 1) = "<" And Right(trimmedStr, 1) = ">" Then
        IsTag = True
    Else
        IsTag = False
    End If

End Function

' ===========================
' (補助) タグ名取得
'   - 概要: <xxx>や</xxx>などの文字列からタグ名を切り出し、終了タグかどうかを判定する
'   - 引数:
'       ByVal tag_str As String      : タグ文字列
'       ByRef is_end_tag As Boolean  : 終了タグならTrueが代入される
'   - 戻り値:
'       String : タグ名(小文字化済み)を返す
' ===========================
Private Function GetTagName(ByVal tag_str As String, _
                            ByRef is_end_tag As Boolean) As String

    Dim tmpStr As String
    tmpStr = Replace(tag_str, "<", "")
    tmpStr = Replace(tmpStr, ">", "")
    tmpStr = Trim(tmpStr)

    If Left(tmpStr, 1) = "/" Then
        is_end_tag = True
        tmpStr = Mid(tmpStr, 2)
    Else
        is_end_tag = False
    End If

    ' 属性があれば最初の空白までがタグ名
    If InStr(tmpStr, " ") > 0 Then
        tmpStr = Left(tmpStr, InStr(tmpStr, " ") - 1)
    End If

    GetTagName = LCase(tmpStr)

End Function

' ===========================
' (補助) インデント文字列取得
'   - 概要: インデントレベルに応じて半角スペースを生成
'   - 引数:
'       ByVal level As Long : インデントレベル
'   - 戻り値:
'       String : インデント用の半角スペース文字列
' ===========================
Private Function GetIndent(ByVal level As Long) As String

    Dim i As Long
    Dim indentStr As String
    indentStr = ""

    For i = 1 To level
        indentStr = indentStr & "  "
    Next i

    GetIndent = indentStr

End Function

' ===========================
' 7) メイン処理
'   - 概要: 指定HTMLファイルを読み込み、改行/空白除去 → 整形 → 出力
'   - 引数:
'       ByVal file_path As String : 対象HTMLファイルのパス
' ===========================
Public Sub FormatHtmlFile(ByVal file_path As String)

    ' 1. HTMLファイル読み込み
    Dim originalHtml As String
    originalHtml = LoadHtmlFile(file_path)

    ' 2. 改行や余分な空白を削除
    Dim compactHtml As String
    compactHtml = RemoveLineBreaksAndExtraSpaces(originalHtml)

    ' 3. 整形
    Dim formattedHtml As String
    formattedHtml = FormatHtml(compactHtml)

    ' 4. 出力
    Debug.Print formattedHtml

    Dim outputPath As String
    outputPath = ThisWorkbook.Path & "\output.html"

    Dim stmObj As Object
    Set stmObj = CreateObject("ADODB.Stream")
    stmObj.Type = 2  ' adTypeText
    stmObj.Charset = "utf-8"
    stmObj.Open
    stmObj.WriteText formattedHtml
    stmObj.SaveToFile outputPath, 2  ' adSaveCreateOverWrite
    stmObj.Close
    Set stmObj = Nothing

    Call MsgBox("整形完了: " & outputPath, vbInformation)

End Sub


