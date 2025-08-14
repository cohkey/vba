Option Explicit

'===============================================================================
' 2つのシートをID列で突き合わせ、差異セルを黄色でハイライト
' 【処理内容】
'   1) 両シートをID列でソート
'   2) 値を配列に取り込み、右シートのID→行Indexの辞書を構築
'   3) 左シートの各レコードについて一致IDの行を見つけ、列ごと比較
'      不一致セルだけを両シート側で黄色ハイライト（バッチ塗り）
' 【前提】
'   - 両シートとも1行目がヘッダー行、2行目以降データ
'   - ID列はヘッダー名で指定（列順が違ってもOK）
'   - 値は文字列で貼り付け済み（自動変換なし前提）
' 【引数】
'   ws_left        : 左側（基準）のシート
'   ws_right       : 右側（比較対象）のシート
'   id_col_name    : ID 列のヘッダー名（例: "ID"）
'   header_row     : ヘッダー行番号（通常は 1）
' 【戻り値】なし（セル背景を直接変更）
'===============================================================================
Public Sub CompareTwoSheetsByIdHighlight( _
    ByVal ws_left As Worksheet, _
    ByVal ws_right As Worksheet, _
    ByVal id_col_name As String, _
    ByVal header_row As Long)

    '--- 最適化
    Dim prevScreenUpdating As Boolean
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    Const YELLOW_BG As Long = &H99FFFF    ' 薄い黄色 (RGB 255,255,153)
    Const BATCH_SIZE As Long = 1024       ' 一度にUnionせず、アドレス結合でまとめ塗り

    '--- 対象範囲（UsedRangeベースでヘッダーからの矩形に限定）
    Dim rangeLeft As Range
    Dim rangeRight As Range
    Set rangeLeft = GetTableRange(ws_left, header_row)
    Set rangeRight = GetTableRange(ws_right, header_row)

    If rangeLeft Is Nothing Or rangeRight Is Nothing Then
        Err.Raise vbObjectError + 1, , "データ範囲（ヘッダー行含む）を取得できませんでした。"
    End If

    '--- ヘッダーとID列位置
    Dim headersLeftArr As Variant
    Dim headersRightArr As Variant
    headersLeftArr = GetHeaderArray(rangeLeft.Rows(1))
    headersRightArr = GetHeaderArray(rangeRight.Rows(1))

    Dim idColLeft As Long
    Dim idColRight As Long
    idColLeft = FindHeaderIndex(headersLeftArr, id_col_name)
    idColRight = FindHeaderIndex(headersRightArr, id_col_name)
    If idColLeft = 0 Or idColRight = 0 Then
        Err.Raise vbObjectError + 2, , "ID列（" & id_col_name & "）が見つかりません。"
    End If

    '--- ID列でソート（昇順）
    Call SortByColumn rangeLeft, idColLeft
    Call SortByColumn rangeRight, idColRight

    '--- データ部（2行目以降）を配列化
    Dim dataLeftArr As Variant
    Dim dataRightArr As Variant
    dataLeftArr = RangeDataToArray(rangeLeft, True)   ' True => データ部のみ
    dataRightArr = RangeDataToArray(rangeRight, True)

    ' 行数/列数
    Dim leftRows As Long, leftCols As Long
    Dim rightRows As Long, rightCols As Long
    leftRows = IIf(IsEmpty(dataLeftArr), 0, UBound(dataLeftArr, 1))
    leftCols = IIf(IsEmpty(dataLeftArr), 0, IIf(leftRows = 0, 0, UBound(dataLeftArr, 2)))
    rightRows = IIf(IsEmpty(dataRightArr), 0, UBound(dataRightArr, 1))
    rightCols = IIf(IsEmpty(dataRightArr), 0, IIf(rightRows = 0, 0, UBound(dataRightArr, 2)))

    If leftRows = 0 Or rightRows = 0 Then
        ' 片方が空なら何もしない
        GoTo CleanExit
    End If

    ' 列数が違う場合でも、共通最小列まで比較（必要なら揃っていることをチェックしてエラーに）
    Dim compareCols As Long
    compareCols = Application.WorksheetFunction.Min(leftCols, rightCols)

    '--- 右シート: ID -> 行Index の辞書を作成（1-basedの配列行Index）
    Dim idToRowRightDic As Scripting.Dictionary
    Set idToRowRightDic = New Scripting.Dictionary
    idToRowRightDic.CompareMode = TextCompare

    Dim r As Long
    For r = 1 To rightRows
        Dim idVal As String
        idVal = CStr(Nz(dataRightArr(r, idColRight), ""))
        If idVal <> "" Then
            If Not idToRowRightDic.Exists(idVal) Then
                Call idToRowRightDic.Add(idVal, r)
            Else
                ' 重複IDは後勝ち。重複を不正扱いにするならここでRaise
                idToRowRightDic(idVal) = r
            End If
        End If
    Next r

    '--- 着色候補を列ごとにバッチング（アドレスを溜め、まとめ塗り）
    Dim leftAddrsDic As Scripting.Dictionary
    Dim rightAddrsDic As Scripting.Dictionary
    Set leftAddrsDic = New Scripting.Dictionary
    Set rightAddrsDic = New Scripting.Dictionary

    ' 列ごとに可変長の文字列配列（バッチ）を持たせる
    Dim c As Long
    For c = 1 To compareCols
        Call leftAddrsDic.Add(c, New Collection)
        Call rightAddrsDic.Add(c, New Collection)
    Next c

    '--- 左→右 突き合わせ
    Dim leftRow As Long
    For leftRow = 1 To leftRows
        Dim leftId As String
        leftId = CStr(Nz(dataLeftArr(leftRow, idColLeft), ""))
        If leftId <> "" Then
            If idToRowRightDic.Exists(leftId) Then
                Dim rightRow As Long
                rightRow = CLng(idToRowRightDic(leftId))

                For c = 1 To compareCols
                    ' ID列も比較対象に含めない（同じID同士の比較なのでスキップ推奨）
                    If c <> idColLeft Or c <> idColRight Then
                        Dim vL As String
                        Dim vR As String
                        vL = CStr(Nz(dataLeftArr(leftRow, c), ""))
                        vR = CStr(Nz(dataRightArr(rightRow, c), ""))

                        ' 完全一致で比較（必要ならTrimや大文字小文字無視に変更）
                        If StrComp(vL, vR, vbBinaryCompare) <> 0 Then
                            ' 左シートのセルアドレス
                            Call AddAddressToBatch(leftAddrsDic(c), _
                                GetCellAddress(rangeLeft, leftRow + 1, c)) ' +1 はヘッダー行オフセット

                            ' 右シートのセルアドレス
                            Call AddAddressToBatch(rightAddrsDic(c), _
                                GetCellAddress(rangeRight, rightRow + 1, c))
                        End If
                    End If
                Next c
            Else
                ' 右に存在しないID（今回の要件では無視。行ごと着色したいならここでA~lastColを追加）
            End If
        End If
    Next leftRow

    '--- 実際にまとめ塗り
    Call PaintBatches(ws_left, leftAddrsDic, YELLOW_BG, BATCH_SIZE)
    Call PaintBatches(ws_right, rightAddrsDic, YELLOW_BG, BATCH_SIZE)

CleanExit:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Exit Sub

ErrHandler:
    MsgBox "CompareTwoSheetsByIdHighlight エラー: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

'===============================================================================
' 範囲（ヘッダー含む表の矩形）を返す。A1起点でなくてもOK
' UsedRangeから、header_row を先頭行として右下端までの矩形を返却
'===============================================================================
Private Function GetTableRange(ByVal ws_sheet As Worksheet, ByVal header_row As Long) As Range
    Dim ur As Range
    Set ur = ws_sheet.UsedRange
    If ur Is Nothing Then
        Set GetTableRange = Nothing
        Exit Function
    End If

    Dim firstCol As Long, lastCol As Long, lastRow As Long
    firstCol = ur.Columns(1).Column
    lastCol = ur.Columns(ur.Columns.Count).Column
    lastRow = ws_sheet.Cells(ws_sheet.Rows.Count, firstCol).End(xlUp).Row

    If lastRow < header_row Then
        Set GetTableRange = Nothing
        Exit Function
    End If

    Set GetTableRange = ws_sheet.Range(ws_sheet.Cells(header_row, firstCol), ws_sheet.Cells(lastRow, lastCol))
End Function

'===============================================================================
' ヘッダー行（1行Range）を配列化（1-based配列）
'===============================================================================
Private Function GetHeaderArray(ByVal header_range As Range) As Variant
    Dim colCount As Long
    colCount = header_range.Columns.Count

    Dim headersArr() As String
    ReDim headersArr(1 To colCount)

    Dim c As Long
    For c = 1 To colCount
        headersArr(c) = CStr(header_range.Cells(1, c).Value)
    Next c

    GetHeaderArray = headersArr
End Function

'===============================================================================
' 指定ヘッダー名の1-based列Indexを返す（見つからなければ0）
'===============================================================================
Private Function FindHeaderIndex(ByVal headers_arr As Variant, ByVal header_name As String) As Long
    Dim c As Long
    For c = LBound(headers_arr) To UBound(headers_arr)
        If StrComp(CStr(headers_arr(c)), header_name, vbTextCompare) = 0 Then
            FindHeaderIndex = c
            Exit Function
        End If
    Next c
    FindHeaderIndex = 0
End Function

'===============================================================================
' 範囲をID列で昇順ソート
'===============================================================================
Private Sub SortByColumn(ByVal table_range As Range, ByVal col_index As Long)
    With table_range
        .Sort Key1:=.Columns(col_index), Order1:=xlAscending, Header:=xlYes
    End With
End Sub

'===============================================================================
' 表Rangeからデータ部（ヘッダー除く）を2次元配列で取得
' 空なら Empty を返す
'===============================================================================
Private Function RangeDataToArray(ByVal table_range As Range, ByVal data_only As Boolean) As Variant
    Dim dataRange As Range
    If data_only Then
        If table_range.Rows.Count <= 1 Then
            RangeDataToArray = Empty
            Exit Function
        End If
        Set dataRange = table_range.Offset(1, 0).Resize(table_range.Rows.Count - 1, table_range.Columns.Count)
    Else
        Set dataRange = table_range
    End If
    RangeDataToArray = dataRange.Value2
End Function

'===============================================================================
' 文字列Null/空対策
'===============================================================================
Private Function Nz(ByVal v As Variant, ByVal default_value As String) As String
    If IsError(v) Then
        Nz = default_value
    ElseIf IsNull(v) Then
        Nz = default_value
    Else
        Nz = CStr(v)
    End If
End Function

'===============================================================================
' セルアドレス（A1形式）を返す：table_rangeの左上を(1,1)として row_index/col_index を指定
' row_index/col_index は表全体基準（ヘッダー含む）
'===============================================================================
Private Function GetCellAddress(ByVal table_range As Range, ByVal row_index As Long, ByVal col_index As Long) As String
    GetCellAddress = table_range.Cells(row_index, col_index).Address(False, False)
End Function

'===============================================================================
' アドレスをコレクションに追加（単純追加）
'===============================================================================
Private Sub AddAddressToBatch(ByVal addr_coll As Collection, ByVal addr As String)
    Call addr_coll.Add(addr)
End Sub

'===============================================================================
' 溜めたアドレスをバッチ毎に塗る（列ごと）
' addrDic: key=列Index (Long), val=Collection of "A1" addresses
'===============================================================================
Private Sub PaintBatches( _
    ByVal ws As Worksheet, _
    ByVal addrDic As Scripting.Dictionary, _
    ByVal color_value As Long, _
    ByVal batch_size As Long)

    Dim c As Variant
    For Each c In addrDic.Keys
        Dim addrColl As Collection
        Set addrColl = addrDic(c)

        Dim i As Long
        Dim batchArr() As String
        Dim batchCount As Long
        batchCount = 0

        If addrColl.Count = 0 Then
            GoTo ContinueNext
        End If

        ReDim batchArr(1 To batch_size)

        For i = 1 To addrColl.Count
            batchCount = batchCount + 1
            batchArr(batchCount) = CStr(addrColl(i))

            If batchCount = batch_size Then
                Call ApplyColorToAddresses(ws, batchArr, batchCount, color_value)
                batchCount = 0
            End If
        Next i

        If batchCount > 0 Then
            Call ApplyColorToAddresses(ws, batchArr, batchCount, color_value)
        End If

ContinueNext:
    Next c
End Sub

'===============================================================================
' 連結アドレスをまとめて着色
'===============================================================================
Private Sub ApplyColorToAddresses( _
    ByVal ws As Worksheet, _
    ByRef addr_arr() As String, _
    ByVal count As Long, _
    ByVal color_value As Long)

    Dim sliceArr() As String
    ReDim sliceArr(1 To count)

    Dim i As Long
    For i = 1 To count
        sliceArr(i) = addr_arr(i)
    Next i

    Dim addrJoined As String
    addrJoined = Join(sliceArr, ",")

    With ws.Range(addrJoined).Interior
        .Pattern = xlSolid
        .Color = color_value
    End With
End Sub