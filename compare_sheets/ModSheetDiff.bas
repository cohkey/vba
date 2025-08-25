'==================== Module: ModSheetDiff ====================
Option Explicit

'============================================================
' 概要   : 2つのシート（Tool_Data, CSV_Data）をIDで突合し、
'          値が異なるCSV側セルのみ黄色背景にする。
'          Tool_Dataに存在しないID行はCSV側データ列をすべて黄色。
' 前提   : 見出しは1行目（HEADER_ROW）、両シートにID列が存在。
' 参照   : Microsoft Scripting Runtime
' 規約   : 事前バインディング / 1ステートメント1行 / 自作SubはCall
'============================================================

Private Const SHEET_TOOL As String = "Sheet1"
Private Const SHEET_CSV As String = "Sheet2"
Private Const ID_HEADER As String = "ID"
Private Const HEADER_ROW As Long = 1
Private Const YELLOW_BG As Long = vbYellow

'------------------------------------------------------------
' 処理内容 : Tool_Data と CSV_Data をIDで比較し、差分をCSV側に着色
' 引数     : なし（モジュール定数を参照）
' 戻り値   : なし
' 例外     : 必須ヘッダ欠如などで Err.Raise 999
'------------------------------------------------------------
Public Sub CompareAndHighlightCsvDifferences()
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean

    prevCalc = Application.Calculation
    prevScreenUpdating = Application.screenUpdating
    prevEnableEvents = Application.enableEvents

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    Dim wsTool As Worksheet
    Dim wsCsv As Worksheet
    Set wsTool = ThisWorkbook.Worksheets(SHEET_TOOL)
    Set wsCsv = ThisWorkbook.Worksheets(SHEET_CSV)

    '--- 配列取得（ヘッダ含む） ---
    Dim toolArr As Variant
    Dim csvArr As Variant
    Dim toolLastRow As Long
    Dim toolLastCol As Long
    Dim csvLastRow As Long
    Dim csvLastCol As Long

    Call ReadUsedBodyAsArray( _
        wsTool, _
        HEADER_ROW, _
        toolArr, _
        toolLastRow, _
        toolLastCol _
    )

    Call ReadUsedBodyAsArray( _
        wsCsv, _
        HEADER_ROW, _
        csvArr, _
        csvLastRow, _
        csvLastCol _
    )

    If toolLastRow < HEADER_ROW + 1 Or csvLastRow < HEADER_ROW + 1 Then
        GoTo CleanUpExit
    End If

    '--- ヘッダ行取得 ---
    Dim toolHeaderArr As Variant
    Dim csvHeaderArr As Variant
    toolHeaderArr = GetHeaderRow(toolArr)
    csvHeaderArr = GetHeaderRow(csvArr)

    '--- ID列インデックス（両シート） ---
    Dim toolIdCol As Long
    Dim csvIdCol As Long
    toolIdCol = GetColumnIndex(toolHeaderArr, ID_HEADER)
    csvIdCol = GetColumnIndex(csvHeaderArr, ID_HEADER)

    If toolIdCol < 1 Or csvIdCol < 1 Then
        Err.Raise 999, , "ID列が見つかりません（ID_HEADER=" & ID_HEADER & "）"
    End If

    '--- Tool側: ID→行Index の辞書（O(1)参照） ---
    Dim idToRowDic As Scripting.Dictionary
    Set idToRowDic = BuildIdToRowIndex(toolArr, toolIdCol)

    '--- 事前に「CSV列 → Tool列」の対応表を作る（ID列は0でスキップ） ---
    Dim csvToToolColArr() As Long
    ReDim csvToToolColArr(1 To csvLastCol)

    Dim c As Long
    For c = 1 To csvLastCol
        If c = csvIdCol Then
            csvToToolColArr(c) = 0
        Else
            Dim head As String
            head = CStr(csvHeaderArr(c))

            If LenB(head) = 0 Then
                csvToToolColArr(c) = 0
            Else
                Dim mapped As Long
                mapped = GetColumnIndex(toolHeaderArr, head)
                If mapped < 1 Then
                    Err.Raise 999, , "Tool_Dataに存在しない列名です: " & head
                End If
                csvToToolColArr(c) = mapped
            End If
        End If
    Next c

    '--- 差分フラグ配列（データ部のみ 1-based） ---
    Dim rowCount As Long
    Dim colCount As Long
    rowCount = csvLastRow - HEADER_ROW
    colCount = csvLastCol

    Dim diffFlagArr() As Boolean
    ReDim diffFlagArr(1 To rowCount, 1 To colCount)

    Dim rowIndex As Long
    For rowIndex = HEADER_ROW + 1 To csvLastRow
        Dim csvKey As String
        csvKey = NormalizeKey(csvArr(rowIndex, csvIdCol))

        Dim hasTool As Boolean
        hasTool = idToRowDic.Exists(csvKey)

        Dim toolRow As Long
        If hasTool Then
            toolRow = CLng(idToRowDic(csvKey))
        End If

        For c = 1 To csvLastCol
            If c <> csvIdCol Then
                If csvToToolColArr(c) > 0 Then
                    If hasTool Then
                        Dim vCsv As Variant
                        Dim vTool As Variant
                        vCsv = csvArr(rowIndex, c)
                        vTool = toolArr(toolRow, csvToToolColArr(c))

                        If Not AreCellValuesEqual(vCsv, vTool) Then
                            diffFlagArr(rowIndex - HEADER_ROW, c) = True
                        End If
                    Else
                        diffFlagArr(rowIndex - HEADER_ROW, c) = True
                    End If
                End If
            End If
        Next c
    Next rowIndex

    '--- 一括でCSV側へ着色 ---
    Call ApplyHighlightByFlags( _
        wsCsv, _
        HEADER_ROW, _
        csvLastRow, _
        csvLastCol, _
        diffFlagArr, _
        YELLOW_BG _
    )

CleanUpExit:
    Application.Calculation = prevCalc
    Application.screenUpdating = prevScreenUpdating
    Application.enableEvents = prevEnableEvents
    Exit Sub

CleanFail:
    Application.Calculation = prevCalc
    Application.screenUpdating = prevScreenUpdating
    Application.enableEvents = prevEnableEvents
    Err.Raise 999, , "CompareAndHighlightCsvDifferences 失敗: " & Err.Description
End Sub


