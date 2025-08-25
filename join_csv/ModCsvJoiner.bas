Option Explicit

'============================================================
' 概要   : 親CSVと複数子CSVを親IDで結合し、1枚のシートに出力する
' 参照   : Microsoft ActiveX Data Objects 2.x Library
'          Microsoft Scripting Runtime
' 規約   : 変数=camelCase / 配列=xxxArr / Dictionary=xxxDic / Collection=xxxColl
'          引数はByVal/ByRef必須、定数はUPPER_SNAKE_CASE
'          Public/Private明示、関数呼び出しは自作SubにCallを付与
'          エラー停止はErr.Raise 999（MsgBox禁止）
'============================================================

'==================== 設定値 ====================
Private Const CONCAT_DELIM As String = " | "

'==================== エントリ ====================
'------------------------------------------------------------
' 処理内容 : 親+子CSVの結合を実行し、結果を指定シートへ出力する
' 引数     : ByVal folderPath   : CSVフォルダパス（末尾\ 可/不可）
'          : ByVal parentCsv    : 親CSVファイル名（例: "Parent.csv"）
'          : ByVal parentIdCol  : 親ID列名（例: "ID"）
'          : ByVal childIdCol   : 子側の親ID列名（例: "ParentID"）※全子共通想定
'          : ByVal targetSheet  : 出力先シート
' 戻り値   : なし
' 例外     : エラー時 Err.Raise 999
'------------------------------------------------------------
Public Sub JoinAllCsvs( _
    ByVal folderPath As String, _
    ByVal parentCsv As String, _
    ByVal parentIdCol As String, _
    ByVal childIdCol As String, _
    ByVal targetSheet As Worksheet _
)
    Dim t0 As Double
    t0 = Timer

    Dim normFolder As String
    normFolder = NormalizeFolderPath(folderPath)

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    prevScreenUpdating = Application.screenUpdating
    prevEnableEvents = Application.enableEvents
    prevCalc = Application.Calculation

    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.enableEvents = False

    On Error GoTo CleanFail

    ' 親読み込み
    Dim parentArr As Variant
    parentArr = LoadCsvToArray(normFolder, parentCsv)
    If IsEmpty(parentArr) Then
        Err.Raise 999, , "親CSVの読込に失敗: " & parentCsv
    End If

    Dim parentHeaderArr As Variant
    parentHeaderArr = GetHeaderRow(parentArr)

    Dim parentIdColIndex As Long
    parentIdColIndex = GetColumnIndex(parentHeaderArr, parentIdCol)
    If parentIdColIndex < 1 Then
        Err.Raise 999, , "親ID列が見つかりません: " & parentIdCol
    End If

    ' 結果配列は親のクローンから開始
    Dim resultArr As Variant
    resultArr = CloneArray2D(parentArr)

    ' 親ID→行Index
    Dim idToRowDic As Scripting.Dictionary
    Set idToRowDic = BuildIdToRowIndex(parentArr, parentIdColIndex)

    ' 子CSV列挙（親CSVを除外）
    Dim childFilesColl As Collection
    Set childFilesColl = ListCsvFiles(normFolder, parentCsv)

    Dim i As Long
    For i = 1 To childFilesColl.Count
        Dim childName As String
        childName = CStr(childFilesColl(i))

        Dim childArr As Variant
        childArr = LoadCsvToArray(normFolder, childName)
        If IsEmpty(childArr) Then
            GoTo ContinueNextChild
        End If

        Dim childHeaderArr As Variant
        childHeaderArr = GetHeaderRow(childArr)

        Dim childKeyColIndex As Long
        childKeyColIndex = GetColumnIndex(childHeaderArr, childIdCol)
        If childKeyColIndex < 1 Then
            GoTo ContinueNextChild
        End If

        Dim childValueColsArr As Variant
        childValueColsArr = ListNonKeyColumns(childHeaderArr, childKeyColIndex)
        If Not ArrayHasElements(childValueColsArr) Then
            GoTo ContinueNextChild
        End If

        Dim childBase As String
        childBase = GetFileBaseName(childName)

        ' 「子名_元列名」で列追加し、子列→結果列のマップを得る
        Dim addedMapDic As Scripting.Dictionary
        Set addedMapDic = EnsureChildColumns( _
            resultArr, _
            childValueColsArr, _
            childHeaderArr, _
            childBase _
        )

        Call AppendChildIntoResult( _
            resultArr, _
            childArr, _
            idToRowDic, _
            childKeyColIndex, _
            childValueColsArr, _
            addedMapDic _
        )
ContinueNextChild:
    Next i

    Call DumpArrayToSheet(targetSheet, resultArr)

    Application.screenUpdating = prevScreenUpdating
    Application.Calculation = prevCalc
    Application.enableEvents = prevEnableEvents

    Debug.Print "JoinAllCsvs done in " & Format$(Timer - t0, "0.000") & " sec"
    Exit Sub

CleanFail:
    Application.screenUpdating = prevScreenUpdating
    Application.Calculation = prevCalc
    Application.enableEvents = prevEnableEvents
    Err.Raise 999, , "JoinAllCsvs 失敗: " & Err.Description
End Sub

'------------------------------------------------------------
' 処理内容 : 結果配列に子列（子名_元列名）を追加し、子→結果列のマッピング辞書を返す
' 引数     : ByRef resultArr         : 結果2次元配列（ヘッダ含む）
'          : ByVal childValueColsArr : 子のキー以外の列インデックス配列(1-based)
'          : ByVal childHeaderArr    : 子のヘッダ配列
'          : ByVal childBaseName     : 子ファイルのベース名
' 戻り値   : Scripting.Dictionary    : Key=子元列Index(String) / Item=結果列Index(Long)
' 例外     : なし
'------------------------------------------------------------
Private Function EnsureChildColumns( _
    ByRef resultArr As Variant, _
    ByVal childValueColsArr As Variant, _
    ByVal childHeaderArr As Variant, _
    ByVal childBaseName As String _
) As Scripting.Dictionary
    Dim mapDic As Scripting.Dictionary
    Set mapDic = New Scripting.Dictionary

    Dim existingHeaderArr As Variant
    existingHeaderArr = GetHeaderRow(resultArr)

    Dim c As Long
    For c = 1 To UBound(childValueColsArr)
        Dim srcCol As Long
        srcCol = CLng(childValueColsArr(c))

        Dim newHeader As String
        newHeader = childBaseName & "_" & CStr(childHeaderArr(srcCol))

        Dim existCol As Long
        existCol = GetColumnIndex(existingHeaderArr, newHeader)

        If existCol < 1 Then
            Call AddColumnToResult(resultArr, newHeader)
            existCol = UBound(resultArr, 2)
            existingHeaderArr = GetHeaderRow(resultArr)
        End If

        Call mapDic.Add(CStr(srcCol), existCol)
    Next c

    Set EnsureChildColumns = mapDic
End Function

'------------------------------------------------------------
' 処理内容 : 子配列の値を結果配列へ反映（1:nはセル内連結、Null安全）
' 引数     : ByRef resultArr         : 結果2次元配列（ヘッダ含む）
'          : ByVal childArr          : 子2次元配列（ヘッダ含む）
'          : ByVal idToRowDic        : 親ID→親配列行Index辞書
'          : ByVal childKeyColIndex  : 子の親ID列Index
'          : ByVal childValueColsArr : 子の値列Index配列
'          : ByVal addedMapDic       : 子元列Index→結果列Indexの辞書
' 戻り値   : なし
' 例外     : なし
' 備考     : 1:n のとき CONCAT_DELIM で連結
'------------------------------------------------------------
Private Sub AppendChildIntoResult( _
    ByRef resultArr As Variant, _
    ByVal childArr As Variant, _
    ByVal idToRowDic As Scripting.Dictionary, _
    ByVal childKeyColIndex As Long, _
    ByVal childValueColsArr As Variant, _
    ByVal addedMapDic As Scripting.Dictionary _
)
    Dim lastChildRow As Long
    lastChildRow = UBound(childArr, 1)

    Dim r As Long
    For r = 2 To lastChildRow
        Dim key As String
        key = NzStr(childArr(r, childKeyColIndex))
        If LenB(key) = 0 Then
            GoTo ContinueNextRow
        End If
        If Not idToRowDic.Exists(key) Then
            GoTo ContinueNextRow
        End If

        Dim parentRow As Long
        parentRow = CLng(idToRowDic(key))

        Dim k As Long
        For k = 1 To UBound(childValueColsArr)
            Dim srcCol As Long
            srcCol = CLng(childValueColsArr(k))

            Dim dstCol As Long
            dstCol = CLng(addedMapDic(CStr(srcCol)))

            Dim vText As String
            vText = NzStr(childArr(r, srcCol))
            If LenB(vText) = 0 Then
                GoTo ContinueNextK
            End If

            Dim curText As String
            curText = NzStr(resultArr(parentRow, dstCol))

            If LenB(curText) = 0 Then
                resultArr(parentRow, dstCol) = vText
            Else
                resultArr(parentRow, dstCol) = curText & CONCAT_DELIM & vText
            End If
ContinueNextK:
        Next k
ContinueNextRow:
    Next r
End Sub


