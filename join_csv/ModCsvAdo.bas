Option Explicit

'============================================================
' 概要   : ADOでCSVを読み込み、ヘッダ付き2次元配列に変換する
' 参照   : Microsoft ActiveX Data Objects 2.x Library
' 規約   : 事前バインディング / エラー時はEmptyを返却（上位でErr.Raise）
'============================================================

Private Const ADO_PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
Private Const TEXT_EXTENDED As String = "text;HDR=YES;FMT=Delimited"

'------------------------------------------------------------
' 処理内容 : ADOでCSVを2次元配列化（ヘッダ行含む、Nullはそのまま）
' 引数     : ByVal folderPath : CSVフォルダパス（末尾\ 可/不可）
'          : ByVal fileName   : CSVファイル名
' 戻り値   : Variant(1 To rows, 1 To cols) / 失敗時 Empty
' 例外     : 内部で捕捉しEmpty返却（上位で判定）
'------------------------------------------------------------
Public Function LoadCsvToArray( _
    ByVal folderPath As String, _
    ByVal fileName As String _
) As Variant
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo ErrHandler

    Set cn = New ADODB.Connection
    cn.Open _
        "Provider=" & ADO_PROVIDER & ";" & _
        "Data Source=" & folderPath & ";" & _
        "Extended Properties='" & TEXT_EXTENDED & "';"

    sql = "SELECT * FROM [" & fileName & "]"

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF And rs.BOF Then
        LoadCsvToArray = Empty
        GoTo FinallyProc
    End If

    LoadCsvToArray = RecordsetToArray(rs)

FinallyProc:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

ErrHandler:
    LoadCsvToArray = Empty
    Resume FinallyProc
End Function

'------------------------------------------------------------
' 処理内容 : Recordset全体をヘッダ付き2次元配列へ変換する
' 引数     : ByVal rs : ADODB.Recordset
' 戻り値   : Variant(1 To rows, 1 To cols)
' 例外     : なし
'------------------------------------------------------------
Private Function RecordsetToArray( _
    ByVal rs As ADODB.Recordset _
) As Variant
    Dim fldCount As Long
    fldCount = rs.Fields.Count

    Dim dataArr As Variant
    dataArr = rs.GetRows()

    Dim rowCount As Long
    rowCount = UBound(dataArr, 2) + 1

    Dim outArr() As Variant
    ReDim outArr(1 To rowCount + 1, 1 To fldCount)

    Dim c As Long
    For c = 1 To fldCount
        outArr(1, c) = rs.Fields(c - 1).Name
    Next c

    Dim r As Long
    For r = 1 To rowCount
        For c = 1 To fldCount
            outArr(r + 1, c) = dataArr(c - 1, r - 1)
        Next c
    Next r

    RecordsetToArray = outArr
End Function


