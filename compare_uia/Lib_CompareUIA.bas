Attribute VB_Name = "Lib_CompareUIA"
Option Explicit



'*************************************
' サブ: CompareOldAndNew
' 概要:
'   Old シートと New シートを比較し、結果を Result シートに出力
' 引数: なし
' 戻り値: なし (結果を Result シートに書き込み)
'*************************************
Public Sub CompareOldAndNew()

    ' 定数 (修正版)
    Const RED_BG As Long = 13027071     ' 赤
    Const BLUE_BG As Long = 15123099   ' 青
    Const PURPLE_BG As Long = 16750280 ' 紫


    ' ローカル変数
    Dim wsOld As Worksheet, wsNew As Worksheet, wsResult As Worksheet
    Dim lastRowOld As Long, lastRowNew As Long, lastCol As Long
    Dim oldDataArr As Variant, newDataArr As Variant
    Dim oldDic As Object, newDic As Object
    Dim i As Long, j As Long, resultRow As Long

    ' ワークシート取得
    Set wsOld = ThisWorkbook.Worksheets("Old")
    Set wsNew = ThisWorkbook.Worksheets("New")

    ' "Result" シート作成
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Result").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "Result"

    ' 最終行・列取得
    lastRowOld = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).row
    lastRowNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).row
    lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column

    ' ヘッダ行コピー
    Call CopyHeader(wsOld, wsNew, wsResult, lastCol)

    ' Old/New のデータを配列化
    If lastRowOld > 1 Then
        oldDataArr = wsOld.Range(wsOld.Cells(2, 1), wsOld.Cells(lastRowOld, lastCol)).Value
    Else
        ReDim oldDataArr(1 To 1, 1 To lastCol)
    End If

    If lastRowNew > 1 Then
        newDataArr = wsNew.Range(wsNew.Cells(2, 1), wsNew.Cells(lastRowNew, lastCol)).Value
    Else
        ReDim newDataArr(1 To 1, 1 To lastCol)
    End If

    ' Dictionary 作成 (キー判定用)
    Set oldDic = CreateObject("Scripting.Dictionary")
    Set newDic = CreateObject("Scripting.Dictionary")

    Dim key As String
    For i = 1 To UBound(oldDataArr, 1)
        key = CreateCompareKey(oldDataArr, i, lastCol)
        If Not oldDic.Exists(key) Then oldDic.Add key, True
    Next i

    For j = 1 To UBound(newDataArr, 1)
        key = CreateCompareKey(newDataArr, j, lastCol)
        If Not newDic.Exists(key) Then newDic.Add key, True
    Next j

    ' 二重走査 (i: Old ポインタ, j: New ポインタ)
    i = 1: j = 1
    resultRow = 2

    While i <= UBound(oldDataArr, 1) Or j <= UBound(newDataArr, 1)

        If i > UBound(oldDataArr, 1) Then
            ' Old 側データなし → 追加
            Call WriteResultRow(wsResult, resultRow, EmptyArray(lastCol), _
                                GetRowArray(newDataArr, j, lastCol), "追加", lastCol)
            Call ColorRow(wsResult, resultRow, lastCol + 2, lastCol * 2 + 1, BLUE_BG)
            j = j + 1

        ElseIf j > UBound(newDataArr, 1) Then
            ' New 側データなし → 削除
            Call WriteResultRow(wsResult, resultRow, GetRowArray(oldDataArr, i, lastCol), _
                                EmptyArray(lastCol), "削除", lastCol)
            Call ColorRow(wsResult, resultRow, 1, lastCol, RED_BG)
            i = i + 1

        Else
            ' 両方データあり → キー比較
            Dim oldKey As String
            Dim newKey As String
            oldKey = CreateCompareKey(oldDataArr, i, lastCol)
            newKey = CreateCompareKey(newDataArr, j, lastCol)

            If oldKey = newKey Then
                ' キー一致 → 値比較
                If CompareArrays(GetRowArray(oldDataArr, i, lastCol), _
                                 GetRowArray(newDataArr, j, lastCol)) Then
                    ' 一致
                    Call WriteResultRow(wsResult, resultRow, GetRowArray(oldDataArr, i, lastCol), _
                                        GetRowArray(newDataArr, j, lastCol), "一致", lastCol)
                Else
                    ' 変更
                    Call WriteResultRow(wsResult, resultRow, GetRowArray(oldDataArr, i, lastCol), _
                                        GetRowArray(newDataArr, j, lastCol), "変更", lastCol)
                    Call HighlightDiffCells(wsResult, resultRow, GetRowArray(oldDataArr, i, lastCol), _
                                            GetRowArray(newDataArr, j, lastCol), PURPLE_BG, lastCol)
                End If
                i = i + 1
                j = j + 1

            ElseIf Not newDic.Exists(oldKey) Then
                ' Old にあるが New にない → 削除
                Call WriteResultRow(wsResult, resultRow, GetRowArray(oldDataArr, i, lastCol), _
                                    EmptyArray(lastCol), "削除", lastCol)
                Call ColorRow(wsResult, resultRow, 1, lastCol, RED_BG)
                i = i + 1

            Else
                ' New にあるが Old にない → 追加
                Call WriteResultRow(wsResult, resultRow, EmptyArray(lastCol), _
                                    GetRowArray(newDataArr, j, lastCol), "追加", lastCol)
                Call ColorRow(wsResult, resultRow, lastCol + 2, lastCol * 2 + 1, BLUE_BG)
                j = j + 1
            End If
        End If

        ' 結果行を進める
        resultRow = resultRow + 1

    Wend


    wsResult.Columns.AutoFit
    MsgBox "比較が完了しました。", vbInformation

End Sub


'----------------------------------------------
' 関数: CreateCompareKey
' 概要:
'   渡された2次元配列の rowNum 行目から、比較用のキー文字列を生成する
' 引数:
'   - data_array (ByVal) : 元データ(2次元配列)
'   - row_num (ByVal)    : 行番号(配列内の1-based)
'   - last_col (ByVal)   : 列数
' 戻り値:
'   - String: 生成されたキー文字列
'----------------------------------------------
Private Function CreateCompareKey(ByVal data_array As Variant, _
                                  ByVal row_num As Long, _
                                  ByVal last_col As Long) As String

    ' ローカル変数 (キャメルケース)
    Dim nameVal As String
    Dim controlVal As String

    ' 例: Name列=2, ControlType列=4
    nameVal = CStr(data_array(row_num, 2))
    controlVal = CStr(data_array(row_num, 3))

    If (nameVal <> "") And (controlVal <> "") Then
        CreateCompareKey = nameVal & "-" & controlVal
    Else
        ' どちらか空なら行全体を連結
        Dim colIndex As Long
        Dim tempKey As String
        tempKey = ""
        For colIndex = 1 To last_col
            tempKey = tempKey & "|" & CStr(data_array(row_num, colIndex))
        Next colIndex
        CreateCompareKey = tempKey
    End If

End Function

'----------------------------------------------
' 関数: GetRowArray
' 概要:
'   2次元配列 data_array の row_num 行目を1次元配列として返す
' 引数:
'   - data_array (ByVal) : 2次元配列
'   - row_num (ByVal)    : 配列内行(1-based)
'   - last_col (ByVal)   : 列数
' 戻り値:
'   - Variant() : 該当行の1次元配列
'----------------------------------------------
Private Function GetRowArray(ByVal data_array As Variant, _
                             ByVal row_num As Long, _
                             ByVal last_col As Long) As Variant

    Dim arr() As Variant
    ReDim arr(0 To last_col - 1)

    Dim c As Long
    For c = 1 To last_col
        arr(c - 1) = data_array(row_num, c)
    Next c

    GetRowArray = arr

End Function

'----------------------------------------------
' 関数: CompareArrays
' 概要:
'   2つの1次元配列の要素が全て等しいかどうか判定する
' 引数:
'   - arr1 (ByVal) : 1次元配列
'   - arr2 (ByVal) : 1次元配列
' 戻り値:
'   - Boolean : True=一致 / False=不一致
'----------------------------------------------
Private Function CompareArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Boolean

    If UBound(arr1) <> UBound(arr2) Then
        CompareArrays = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) <> arr2(i) Then
            CompareArrays = False
            Exit Function
        End If
    Next i

    CompareArrays = True

End Function

'----------------------------------------------
' サブ: WriteResultRow
' 概要:
'   Resultシートの指定行(resultRow)に、Old/Status/New を書き込む
' 引数:
'   - ws_result (ByVal) : 出力先シート
'   - result_row (ByVal): 書き込み先行
'   - arr_old (ByVal)   : Old側1次元配列
'   - arr_new (ByVal)   : New側1次元配列
'   - status_str (ByVal): "一致" "変更" "追加" "削除" 等
'   - last_col (ByVal)  : 列数
' 戻り値: なし
'----------------------------------------------
Private Sub WriteResultRow(ByVal ws_result As Worksheet, _
                           ByVal result_row As Long, _
                           ByVal arr_old As Variant, _
                           ByVal arr_new As Variant, _
                           ByVal status_str As String, _
                           ByVal last_col As Long)

    Dim c As Long

    ' Old 側
    For c = 0 To last_col - 1
        ws_result.Cells(result_row, 1 + c).Value = arr_old(c)
    Next c

    ' Status
    ws_result.Cells(result_row, last_col + 1).Value = status_str

    ' New 側
    For c = 0 To last_col - 1
        ws_result.Cells(result_row, last_col + 2 + c).Value = arr_new(c)
    Next c

End Sub

'----------------------------------------------
' サブ: HighlightDiffCells
' 概要:
'   Old/New 1次元配列の不一致セルを紫色に着色
' 引数:
'   - ws_result (ByVal)   : 着色先シート
'   - result_row (ByVal)  : 着色行
'   - arr_old (ByVal)     : Old側1次元配列
'   - arr_new (ByVal)     : New側1次元配列
'   - color_code (ByVal)  : 着色カラー
'   - last_col (ByVal)    : 列数
'----------------------------------------------
Private Sub HighlightDiffCells(ByVal ws_result As Worksheet, _
                               ByVal result_row As Long, _
                               ByVal arr_old As Variant, _
                               ByVal arr_new As Variant, _
                               ByVal color_code As Long, _
                               ByVal last_col As Long)

    Dim c As Long
    For c = 0 To UBound(arr_old)
        If arr_old(c) <> arr_new(c) Then
            ' Old 側
            ws_result.Cells(result_row, 1 + c).Interior.Color = color_code
            ' New 側
            ws_result.Cells(result_row, last_col + 2 + c).Interior.Color = color_code
        End If
    Next c

End Sub

'----------------------------------------------
' サブ: ColorRow
' 概要:
'   指定行(resultRow)の、startCol～endCol 列を colorCode で塗りつぶす
' 引数:
'   - ws_result (ByVal) : 対象シート
'   - result_row (ByVal): 行番号
'   - start_col (ByVal) : 開始列
'   - end_col (ByVal)   : 終了列
'   - color_code (ByVal): 塗りつぶしカラー
'----------------------------------------------
Private Sub ColorRow(ByVal ws_result As Worksheet, _
                     ByVal result_row As Long, _
                     ByVal start_col As Long, _
                     ByVal end_col As Long, _
                     ByVal color_code As Long)

    Dim c As Long
    For c = start_col To end_col
        ws_result.Cells(result_row, c).Interior.Color = color_code
    Next c

End Sub

'----------------------------------------------
' 関数: EmptyArray
' 概要:
'   指定列数 colCount の要素をすべて "" とした1次元配列を返す
' 引数:
'   - col_count (ByVal): 列数
' 戻り値:
'   - Variant() : 要素がすべて空文字の1次元配列
'----------------------------------------------
Private Function EmptyArray(ByVal col_count As Long) As Variant

    Dim arr() As Variant
    ReDim arr(0 To col_count - 1)

    Dim i As Long
    For i = 0 To col_count - 1
        arr(i) = ""
    Next i

    EmptyArray = arr

End Function

'----------------------------------------------
' サブ: CopyHeader
' 概要:
'   Old/New シートのヘッダ1行目をResultにコピーし、
'   その間に "Status" 列を挟む
' 引数:
'   - ws_old (ByVal)    : Oldシート
'   - ws_new (ByVal)    : Newシート
'   - ws_result (ByVal) : 出力先シート
'   - last_col (ByVal)  : 列数
' 戻り値: なし
'----------------------------------------------
Private Sub CopyHeader(ByVal ws_old As Worksheet, _
                       ByVal ws_new As Worksheet, _
                       ByVal ws_result As Worksheet, _
                       ByVal last_col As Long)

    Dim oldHeader As Variant
    Dim newHeader As Variant
    Dim c As Long

    oldHeader = ws_old.Range(ws_old.Cells(1, 1), ws_old.Cells(1, last_col)).Value
    For c = 1 To last_col
        ws_result.Cells(1, c).Value = oldHeader(1, c)
    Next c

    ws_result.Cells(1, last_col + 1).Value = "Status"

    newHeader = ws_new.Range(ws_new.Cells(1, 1), ws_new.Cells(1, last_col)).Value
    For c = 1 To last_col
        ws_result.Cells(1, last_col + 1 + c).Value = newHeader(1, c)
    Next c

End Sub


