Option Explicit

'============================================================
' 概要   : セル値の等価判定を型感度高く行う（数値/日付/文字列）
' 規約   : 事前バインディング不要
'============================================================

'------------------------------------------------------------
' 処理内容 : 値の等価判定
' ルール   : 両方数値→CDbl比較、両方日付→シリアル比較、その他→Trim文字列（区別）
' 引数     : ByVal v1, ByVal v2
' 戻り値   : Boolean 一致=True
'------------------------------------------------------------
Public Function AreCellValuesEqual( _
    ByVal v1 As Variant, _
    ByVal v2 As Variant _
) As Boolean
    If IsError(v1) Or IsError(v2) Then
        AreCellValuesEqual = False
        Exit Function
    End If

    If (IsNull(v1) Or v1 = vbNullString) And (IsNull(v2) Or v2 = vbNullString) Then
        AreCellValuesEqual = True
        Exit Function
    End If

    If IsNumeric(v1) And IsNumeric(v2) Then
        AreCellValuesEqual = (CDbl(v1) = CDbl(v2))
        Exit Function
    End If

    If IsDate(v1) And IsDate(v2) Then
        AreCellValuesEqual = (CDbl(CDate(v1)) = CDbl(CDate(v2)))
        Exit Function
    End If

    Dim s1 As String
    Dim s2 As String
    s1 = Trim$(CStr(v1))
    s2 = Trim$(CStr(v2))
    AreCellValuesEqual = (StrComp(s1, s2, vbBinaryCompare) = 0)
End Function


