Option Explicit

' API宣言 (64bit対応)
Private Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As LongPtr) As Long
Private Declare PtrSafe Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" ( _
    ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, _
    ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As LongPtr) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, lpiid As Any) As Long
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As LongPtr, riid As Any, _
    ByVal wParam As LongPtr, ppvObject As Object) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' 定数
Private Const GW_CHILD    As Long = 5
Private Const GW_HWNDNEXT As Long = 2
Private Const SMTO_ABORTIFHUNG As Long = &H2

' Edgeブラウザのプロセス名
Private Const PROCESS_NAME_ED As String = "msedge.exe"

' ===============================
'  1) EdgeのHTMLドキュメントを取得するメイン関数
' ===============================
Public Function GetHtmlDocument(ByVal title As String) As Object
    Dim con        As Object
    Dim items      As Object
    Dim htmlDoc    As Object
    Dim hwnd       As LongPtr
    Dim pid        As LongPtr
    Dim buf        As String * 255
    Dim className  As String

    ' WMI接続
    Set con = CreateObject("WbemScripting.SWbemLocator").ConnectServer

    ' 最上位ウィンドウのハンドルを取得
    hwnd = GetTopWindow(0)

    Do While hwnd <> 0
        ' クラス名を取得
        GetClassName hwnd, buf, Len(buf)
        className = Left$(buf, InStr(buf, vbNullChar) - 1)

        ' Chrome/Edge のメインウィンドウを特定
        If InStr(className, "Chrome_WidgetWin_") > 0 Then
            ' ウィンドウからプロセスIDを取得
            GetWindowThreadProcessId hwnd, pid

            ' プロセスが msedge.exe か確認
            Set items = con.ExecQuery( _
                "Select ProcessId From Win32_Process " & _
                "Where (ProcessId = '" & pid & "') And (Name = '" & PROCESS_NAME_ED & "')" _
            )

            If items.Count > 0 Then
                ' 再帰的に子ウィンドウを探して Internet Explorer_Server を見つける
                Dim hIES As LongPtr
                hIES = FindIESChildWindow(hwnd)

                If hIES <> 0 Then
                    Set htmlDoc = GetHTMLDocumentFromIES(hIES)

                    ' タイトルが一致すれば返却
                    If Not htmlDoc Is Nothing Then
                        If htmlDoc.title Like title Then
                            Set GetHtmlDocument = htmlDoc
                            Exit Do
                        End If
                    End If
                End If
            End If
        End If

        ' 次の兄弟ウィンドウ
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop

End Function

' ===============================
'  2) 再帰的に子ウィンドウを探す関数
'    "Internet Explorer_Server" クラスを見つけたらそのhWndを返す
' ===============================
Private Function FindIESChildWindow(ByVal hWndParent As LongPtr) As LongPtr
    Dim hChild    As LongPtr
    Dim buf       As String * 255
    Dim className As String
    Dim found     As LongPtr

    ' 最初の子ウィンドウ
    hChild = GetWindow(hWndParent, GW_CHILD)

    While hChild <> 0
        ' クラス名を取得
        GetClassName hChild, buf, Len(buf)
        className = Left$(buf, InStr(buf, vbNullChar) - 1)

        ' Internet Explorer_Server を見つけたら返す
        If className = "Internet Explorer_Server" Then
            FindIESChildWindow = hChild
            Exit Function
        End If

        ' 子ウィンドウにもさらに子がある場合、再帰で探す
        found = FindIESChildWindow(hChild)
        If found <> 0 Then
            FindIESChildWindow = found
            Exit Function
        End If

        ' 次の兄弟ウィンドウ
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Wend
End Function

' ===============================
'  3) IESハンドルからHTMLドキュメントを取得する関数
' ===============================
Private Function GetHTMLDocumentFromIES(ByVal hwnd As LongPtr) As Object
    Dim msg As Long
    Dim res As LongPtr
    Dim iid(0 To 3) As Long
    Dim ret As Object
    Dim obj As Object

    Const IID_IHTMLDocument2 As String = "{332C4425-26CB-11D0-B483-00C04FD90119}"

    Set ret = Nothing

    ' ドキュメントを取得するためのメッセージ
    msg = RegisterWindowMessage("WM_HTML_GETOBJECT")

    ' メッセージ送信
    SendMessageTimeout hwnd, msg, 0, 0, SMTO_ABORTIFHUNG, 1000, res

    If res <> 0 Then
        ' IHTMLDocument2 のIID
        IIDFromString StrPtr(IID_IHTMLDocument2), iid(0)

        ' ObjectFromLresultでHTMLDocument2を取得
        If ObjectFromLresult(res, iid(0), 0, obj) = 0 Then
            Set ret = obj
        End If
    End If

    Set GetHTMLDocumentFromIES = ret
End Function


