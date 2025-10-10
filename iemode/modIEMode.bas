' ============================================
'  Windows API宣言（64bit 対応）
' ============================================

' ウィンドウ検索
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

' ウィンドウ状態
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Private Declare PtrSafe Function DwmGetWindowAttribute Lib "dwmapi" _
    (ByVal hwnd As LongPtr, _
     ByVal dwAttribute As Long, _
     ByRef pvAttribute As Any, _
     ByVal cbAttribute As Long) As Long

' DwmGetWindowAttribute の定数
Private Const DWMWA_CLOAKED As Long = 14
Private Const S_OK As Long = 0

' GetWindow の定義／定数
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Const GW_CHILD As Long = 5
Private Const GW_HWNDNEXT As Long = 2

' 文字列取得など
Private Declare PtrSafe Function GetTopWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, ByRef lpdwProcessId As LongPtr) As Long

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long

' オブジェクト変換（Msg/ID 取得）
Private Declare PtrSafe Function IIDFromString Lib "ole32" _
    (ByVal lpsz As Any, ByVal lpiid As LongPtr) As Long

Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" _
    (ByVal lResult As LongPtr, ByRef riid As Any, ByVal wParam As LongPtr, ByRef ppvObject As Object) As Long

Private Declare PtrSafe Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" _
    (ByVal lpString As String) As Long

Private Declare PtrSafe Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" _
    (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, _
     ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As LongPtr) As Long

' SendMessageTimeout の定数
Private Const SMTO_ABORTIFHUNG As Long = &H2

' 待機用
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Declare Function timeGetTime Lib "winmm.dll" () As Long  ' TODO: 削除予定


' ============================================
'  メンバー変数
' ============================================
Private WithEvents m_htmlDoc As MSHTML.HTMLDocument
Private WithEvents m_parentWindow As MSHTML.HTMLWindow2
Private m_uia As CUIAutomation8
Private m_uiBrowser As IUIAutomationElement

Private m_isExplore As Boolean
Private m_hWnd As LongPtr
Private m_processId As LongPtr
Private m_hWndList As VBA.Collection

' 開発者モードを実現する変数
Private m_developerMode As Boolean
Private m_setTime As String
Private m_eventListener As Boolean
Private m_eventOutput As Boolean
Private m_eventItems As Collection
Private m_timeoutListener As Boolean
Private m_timeoutOutput As Boolean
Private m_timeoutItems As Collection

' イベント関連
Private m_onUnloadEvent As Boolean
Private m_onLoadEvent As Boolean
Private m_onBeforeDeactivateEvent As Boolean
Private m_onReadyStateChangeEvent As Boolean

Private Const TIMEOUT_MAXIMUM As Long = 120000     '// 一日分のタイムアウト時間（24hour×60minx60sec×1000milisec-2、これを超えたらオーバーフローとして0扱い、comment by：小野）
Private Const TIMEOUT_NORMAL  As Long = 10000      ' // Waitの前待機のタイムアウト時間（10秒）

Private Const TITLE_LOGIN_PAGE As String = "ＴＥＰＣＯ 共通ログイン"
Private Const TITLE_INTRA_MENU As String = "イントラネット・システムメニュー"
Private Const TITLE_IE_SITEMAP As String = "IEサイトマップ"

' 待機タイプ指定
Public Enum WaitAdvancedType
    BY_NEW_WINDOW = 1
    BY_TITLE = 2
End Enum


' HTML要素指定用
Public Enum HtmlUniqueElement
    TYPE_ID = 1
    TYPE_NAME = 2
    TYPE_CLASS_NAME = 3
    TYPE_TAG_NAME = 4
End Enum

' マッチパターン指定用
Public Enum MatchPattern
    LIKE_BOTH = 1
    LIKE_ATTRIBUTE = 2
    PERFECT_INNER = 3
    LIKE_INNER = 4
End Enum

' ウィンドウ情報格納用
Private Type WindowInfo
    wName As String
    wClassName As String
    wHandle As LongPtr
End Type


' ============================================
'  ライフサイクル
' ============================================
Private Sub Class_Initialize()
    Set m_uia = New CUIAutomation8
    Set m_eventItems = New Collection
    Set m_timeoutItems = New Collection
    m_setTime = Replace(CStr(Date), "/", "") & "_" & Replace(CStr(Time), ":", "")
End Sub

Private Sub Class_Terminate()
    Set m_uiBrowser = Nothing
    Set m_uia = Nothing
    Set m_htmlDoc = Nothing
    Set m_parentWindow = Nothing

    If m_eventOutput = True Or m_timeoutOutput = True Then
        ' TODO: 出力ロジック（OutputExporter?）が別途存在
        ' Call OutputExporter(...)
    End If

    Set m_eventItems = Nothing
    Set m_timeoutItems = Nothing
End Sub


' ============================================
'  プロパティ
' ============================================
' Internet Explorerで強制したい場合に使用（TrueでIE互換）
Public Property Let IsExplore(ByVal is_explore As Boolean)
    m_isExplore = is_explore
End Property

Public Property Get IsExplore() As Boolean
    IsExplore = m_isExplore
End Property

Public Property Get IEWindow() As MSHTML.HTMLWindow2
    Set IEWindow = m_htmlDoc.parentWindow   ' Document処理中も参照できる（title等は更新中も取得可能）
End Property

Public Property Get IEDocument() As MSHTML.HTMLDocument
    Set IEDocument = m_htmlDoc
End Property

Public Property Get FrameDocument(ByVal index As Long) As HTMLDocument
    Set FrameDocument = m_htmlDoc.frames.Item(index).document
End Property

Public Property Get FrameLength() As Long
    FrameLength = m_htmlDoc.frames.Length
End Property

Public Property Get LocationURL() As String
    LocationURL = m_htmlDoc.url
    ' 以前は: m_parentWindow.Location.href
End Property

Public Property Get hwnd() As LongPtr
    hwnd = m_hWnd
End Property

Public Property Get ReadyState() As String
    ReadyState = m_htmlDoc.ReadyState
End Property

Public Property Let EventVisible(ByVal visible_apply As Boolean)
    m_eventListener = visible_apply
End Property

Public Property Let EventExport(ByVal export_apply As Boolean)
    m_eventOutput = export_apply
End Property

Public Property Let TimeoutVisible(ByVal visible_apply As Boolean)
    m_timeoutListener = visible_apply
End Property

Public Property Let TimeoutExport(ByVal export_apply As Boolean)
    m_timeoutOutput = export_apply
End Property

Public Property Let DeveloperMode(ByVal mode_apply As Boolean)
    m_developerMode = mode_apply
    If m_developerMode Then
        m_eventListener = True
        m_eventOutput = True
        m_timeoutListener = True
        m_timeoutOutput = True
    End If
End Property


' ============================================
'  初期化（ハンドルから）
' ============================================
' 機能: ウィンドウハンドルによるメンバー変数の初期化
' 引数: target_hWnd 対象ウィンドウハンドル
' 返値: 成否(Boolean)
Private Function InitializeByWindowHandle(ByVal target_hWnd As LongPtr) As Boolean
    If IsWindow(target_hWnd) = 0 Then
        Exit Function
    End If

    m_hWnd = target_hWnd
    Set m_htmlDoc = Nothing
    Set m_parentWindow = Nothing

    Set m_htmlDoc = GetHtmlDocumentByHandle(m_hWnd)

    ' TODO: デバッグ用 呼び出し
    ' Debug.Print "----- TestGetWindowInfoArray"
    ' Call TestGetWindowInfoArray

    If m_htmlDoc Is Nothing Then
        Err.Raise 999, , "(InitializeByWindowHandle)" & vbCrLf & _
            "指定したハンドルからHTMLDocumentを取得できませんでした。" & vbCrLf & _
            "hWnd: " & target_hWnd
    End If

    Set m_parentWindow = m_htmlDoc.parentWindow
    Set m_uiBrowser = m_uia.ElementFromHandle(ByVal m_hWnd)

    InitializeByWindowHandle = True
End Function


' ============================================
'  起動＆ナビゲート
' ============================================
' 機能 : Internet Explorer/IEモード相当でオブジェクトをインスタンスし、ページ表示を行う。
' 引数 : navigate_url 表示する対象URL
'        do_wait      Navigate後の待機実行フラグ（既定 False）
'        visible_event 監視イベントをログに出すか（既定 False）
'        visible_wait_time 監視ログのタイムスタンプ可視化（既定 False）
'        ※ 1度だけ新規ウィンドウで開いてから対象ウィンドウへフォーカスを戻す（Class_IEでも同様に機能していない）
Public Function OpenIE( _
    ByVal navigate_url As String, _
    Optional ByVal do_wait As Boolean = False, _
    Optional ByVal visible_event As Boolean = False, _
    Optional ByVal visible_wait_time As Boolean = False) As Boolean

    Dim oldWinInfoArr() As WindowInfo
    oldWinInfoArr = GetWindowInfoArray()

    Dim errMsg As String
    On Error Resume Next
    Dim arrIndex As Long
    arrIndex = UBound(oldWinInfoArr)
    If Err.Number <> 0 Then
        errMsg = "ブラウザ情報が取得できていません。"
    End If
    On Error GoTo 0

    If errMsg <> "" Then
        Err.Raise 999, , errMsg
    End If

    ' 既定ブラウザ起動（IE or Edge-IEモード）
    If m_isExplore Then
        CreateObject("WScript.Shell").Run "iexplore.exe " & navigate_url
    Else
        ' EdgeのIEモードで新規ウィンドウ
        CreateObject("WScript.Shell").Run "msedge.exe --url " & navigate_url & " --new-window"
        ' TODO: 実行環境に合わせて起動オプション調整の可能性あり
    End If

    ' 既に起動しているIEモード・ドウインドウを取得
    Dim topTitle As String
    If m_isExplore Then
        topTitle = TITLE_IE_SITEMAP
    Else
        topTitle = TITLE_INTRA_MENU
    End If

    If LoadWindow(topTitle, False) = False Then
        Err.Raise 999, , "IEモードの " & topTitle & " が見つかりません。"
    End If

    ' jsのwindow.openでポップアップ生成→一時ハンドル → そこから本ウィンドウへ
    ' （直接openのポップアップで開けば影響ない？）
    m_parentWindow.execScript "window.open('" & navigate_url & "', 'newWindow', 'popup')"
    Dim newWin As IHTMLWindow2
    Set newWin = m_parentWindow.Open(navigate_url, "newWindow", "popup", False)
    Debug.Print newWin.Name, newWin.Location.href

    Dim newWinInfoArr() As WindowInfo
    Dim newHandle As LongPtr
    Dim startTime As Double
    Dim retryCount As Long

Lbl_Retry:
    startTime = Timer()
    Do
        If (Timer() - startTime) > 10 Then
            Err.Raise 999, , "新規ウィンドウハンドルの待機処理がタイムアウトしました。" & vbCrLf & _
                              "対象URL: " & navigate_url
        End If

        newWinInfoArr = GetWindowInfoArray()

        ' 新規ウィンドウが1つ増えたタイミングでハンドル抽出
        If UBound(newWinInfoArr) = UBound(oldWinInfoArr) + 1 Then
            newHandle = ExtractNewHandle(oldWinInfoArr, newWinInfoArr)
        End If

        DoEvents
        Sleep 100
    Loop Until newHandle > 0

    ' HTMLDocument準備ができるまで待機
    If WaitNewWindowByHandle(newHandle) = False Then
        If retryCount = 0 Then
            Debug.Print "WaitNewWindowByHandle 直後にリトライ"
            retryCount = retryCount + 1
            GoTo Lbl_Retry
        End If
    End If

    ' メンバー変数の初期化
    If InitializeByWindowHandle(newHandle) = False Then
        If retryCount = 0 Then
            Debug.Print "InitializeByWindowHandle 直後にリトライ"
            retryCount = retryCount + 1
            GoTo Lbl_Retry
        End If
    End If

    If do_wait Then
        Call Me.Wait
    End If

    Debug.Print "After InitializeByWindowHandle"
    Debug.Print m_parentWindow.Name, m_parentWindow.Location.href

    If m_htmlDoc.title = TITLE_LOGIN_PAGE Then
        Call SetForegroundWindow(Me.hwnd)
        Call SetForegroundWindow(Application.hwnd)
'        Call EndWithMessage("インタネットログインのタイムアウトが発生しました。", vbExclamation)
    End If

    OpenIE = True
End Function


' =====================================================
'  新規ウィンドウのハンドル抽出
'   old_wininfo_arr  : 新規ウィンドウ起動前の WindowInfo 配列
'   new_wininfo_arr  : 新規ウィンドウ起動後の WindowInfo 配列（ウィンドウが増えている）
'   戻り値            : 新規ウィンドウのハンドル／見つからないとき 0
' =====================================================
Private Function ExtractNewHandle(ByRef old_wininfo_arr() As WindowInfo, _
                                  ByRef new_wininfo_arr() As WindowInfo) As LongPtr
    Dim i As Long, j As Long
    Dim oldWinInfo As WindowInfo, newWinInfo As WindowInfo
    Dim isOldHandle As Boolean

    For i = LBound(new_wininfo_arr) To UBound(new_wininfo_arr)
        isOldHandle = False
        newWinInfo = new_wininfo_arr(i)

        For j = LBound(old_wininfo_arr) To UBound(old_wininfo_arr)
            oldWinInfo = old_wininfo_arr(j)
            If newWinInfo.wHandle = oldWinInfo.wHandle Then
                isOldHandle = True
                Exit For
            End If
        Next

        If isOldHandle = False Then
            ' 旧ウィンドウリストに存在しない => 新規ウィンドウ
            ExtractNewHandle = newWinInfo.wHandle
            Exit Function
        End If
    Next
End Function

' ' TODO: 削除予定（テスト用ダミー）
' Public Function TestExtractNewHandle()
'     Dim oldWinInfoArr(2) As WindowInfo, newWinInfoArr(3) As WindowInfo
'     oldWinInfoArr(0).wHandle = 10000
'     oldWinInfoArr(1).wHandle = 10001
'     oldWinInfoArr(2).wHandle = 10002
'     newWinInfoArr(0).wHandle = 10000
'     newWinInfoArr(1).wHandle = 10001
'     newWinInfoArr(2).wHandle = 10002
'     newWinInfoArr(3).wHandle = 10003
' End Function


' =====================================================
'  タイトル or URL から IE/IEモードの HTMLDocument を取得
'   title_or_url : 対象タイトルまたはURL
'   full_match   : True=完全一致 / False=部分一致
'   return_hWnd  : 戻り値として対象ウィンドウの hWnd を返す（既定=-1）
'   戻り値       : 見つかった HTMLDocument / 失敗=Nothing
' =====================================================
Private Function GetHtmlDocument(ByVal title_or_url As String, _
                                 Optional ByVal full_match As Boolean = False, _
                                 Optional ByRef return_hWnd As LongPtr = -1) As MSHTML.HTMLDocument
    Dim winInfoArr() As WindowInfo
    Dim hwnd As LongPtr
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim urlFlag As Boolean
    Dim i As Long

    ' 先頭がURLっぽいかどうかで判定
    If IsURL(title_or_url) Then
        urlFlag = True
    Else
        title_or_url = CleanHtmlChar(title_or_url)
    End If

    winInfoArr = GetWindowInfoArray()

    For i = 0 To UBound(winInfoArr)
        Set htmlDoc = GetHtmlDocumentByHandle(winInfoArr(i).wHandle)
        If Not htmlDoc Is Nothing Then
            ' 一致判定
            If IsMatchHtmlTitleOrURL(htmlDoc, title_or_url, full_match) Then
                Set GetHtmlDocument = htmlDoc
                return_hWnd = winInfoArr(i).wHandle
                Exit For
            End If
        End If
    Next
End Function


' =====================================================
'  ウィンドウハンドルから IE/Edge IEモード の HTMLDocument を取得
'   target_hWnd : 対象ウィンドウのハンドル
'   戻り値      : 成功=対象の HTMLDocument / 失敗=Nothing
' =====================================================
Private Function GetHtmlDocumentByHandle(ByVal target_hWnd As LongPtr) As MSHTML.HTMLDocument
    Dim con As Object
    Dim items As Object
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim pid As LongPtr
    Dim buf As String * 255
    Dim wClassName As String
    Dim targetClassName As String
    Dim processName As String

    If target_hWnd <= 0 Then
        Err.Raise 999, , "(GetHtmlDocumentByHandle)" & vbCrLf & _
                         "指定したハンドルが不正です。" & vbCrLf & _
                         "target_hWnd: " & target_hWnd
    End If

    If m_isExplore Then
        processName = "iexplore.exe"
        targetClassName = "IEFrame"
    Else
        processName = "msedge.exe"
        targetClassName = "Chrome_WidgetWin_1"
    End If

    ' WMI 接続
    Set con = CreateObject("WbemScripting.SWbemLocator").ConnectServer

    ' クラス名を取得
    GetClassName target_hWnd, buf, Len(buf)
    wClassName = Left$(buf, InStr(buf, vbNullChar) - 1)

    ' IE/Edge のメインウィンドウ判定
    If wClassName = targetClassName Then
        ' ウィンドウからプロセスIDを取得
        GetWindowThreadProcessId target_hWnd, pid
    End If

    ' プロセスが msedge.exe か確認
    Set items = con.ExecQuery( _
        "Select ProcessId From Win32_Process " & _
        "Where (ProcessId = " & pid & ") And (Name = '" & processName & "')")

    If items.Count > 0 Then
        ' 再帰的に子ウィンドウを探して Internet Explorer_Server を見つける
        Dim hIES As LongPtr
        hIES = FindIESChildWindow(target_hWnd)

        If hIES <> 0 Then
            Set htmlDoc = GetHTMLDocumentFromIES(hIES)
            If Not htmlDoc Is Nothing Then
                Set GetHtmlDocumentByHandle = htmlDoc

                ' メンバーを同期化（ハンドル／PID）
                m_hWnd = target_hWnd
                m_processId = pid
            End If
        End If
    End If
End Function


' =====================================================
'  タイトルまたはURLの一致判定
'   html_Doc    : 対象の HTMLDocument
'   title_or_url: 比較するタイトルまたはURL
'   full_match  : True=完全一致 / False=部分一致
'   戻り値      : True/False
' =====================================================
Private Function IsMatchHtmlTitleOrURL(ByVal html_Doc As MSHTML.HTMLDocument, _
                                       ByVal title_or_url As String, _
                                       Optional ByVal full_match As Boolean = False) As Boolean
    Dim htmlTitleOrURL As String

    If IsURL(title_or_url) Then
        htmlTitleOrURL = URLEncode(title_or_url)
        title_or_url = URLEncode(html_Doc.Location.href)
    Else
        htmlTitleOrURL = html_Doc.title
        title_or_url = CleanHtmlChar(title_or_url)
    End If

    If full_match Then
        If htmlTitleOrURL = title_or_url Then
            IsMatchHtmlTitleOrURL = True
        End If
    Else
        ' 文字種の違い等があるため、今回は InStr による部分一致
        If InStr(htmlTitleOrURL, title_or_url) > 0 Then
            IsMatchHtmlTitleOrURL = True
        End If
    End If
End Function


' =====================================================
'  再帰探索で "Internet Explorer_Server" クラスを見つけて hWnd を返す
'   parent_hWnd : 親ウィンドウのハンドル
'   戻り値      : 成功=>IES クラスの hWnd / 失敗=>0
' =====================================================
Private Function FindIESChildWindow(ByVal parent_hWnd As LongPtr) As LongPtr
    Dim hChild As LongPtr
    Dim buf As String * 255
    Dim wClassName As String
    Dim found As LongPtr

    ' 最初の子ウィンドウ
    hChild = GetWindow(parent_hWnd, GW_CHILD)

    While hChild <> 0
        GetClassName hChild, buf, Len(buf)
        wClassName = Left$(buf, InStr(buf, vbNullChar) - 1)

        If wClassName = "Internet Explorer_Server" Then
            FindIESChildWindow = hChild
            Exit Function
        End If

        ' 子ウィンドウにもさらに子がある場合、再帰探索
        found = FindIESChildWindow(hChild)
        If found <> 0 Then
            FindIESChildWindow = found
            Exit Function
        End If

        ' 次の兄弟ウィンドウ
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Wend
End Function


' =====================================================
'  IES(Internet Explorer Server) のウィンドウハンドルから HTMLDocument を取得
'   target_hWnd : IES のハンドル
'   戻り値      : 成功=HTMLDocument / 失敗=Nothing
' =====================================================
Private Function GetHTMLDocumentFromIES(ByVal target_hWnd As LongPtr) As Object
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
    SendMessageTimeout target_hWnd, msg, 0, 0, SMTO_ABORTIFHUNG, 1000, res

    If res <> 0 Then
        ' IHTMLDocument2 の IID
        IIDFromString StrPtr(IID_IHTMLDocument2), iid(0)

        ' ObjectFromLresult で HTMLDocument2 を取得
        If ObjectFromLresult(res, iid(0), 0, obj) = 0 Then
            Set ret = obj
        End If
    End If

    Set GetHTMLDocumentFromIES = ret
End Function


' =====================================================
'  指定URLにページ遷移する（オプションで待機）
'   navigate_url   : 遷移先URL
'   do_wait        : 待機実行フラグ（既定 True）
'   timeout_milli  : タイムアウト上限（ミリ秒。0/省略で内部既定）
'   wait_title     : 遷移完了の判定にタイトル一致を使う場合のタイトル（Class_IE同様に追加）
'   new_tab        : True=新規タブで開く
' =====================================================
Public Function NavigatePage(ByVal navigate_url As String, _
                             Optional do_wait As Boolean = True, _
                             Optional timeout_milli_sec As Long = 0, _
                             Optional wait_title As String = "", _
                             Optional new_tab As Boolean = False) As Boolean
    With m_parentWindow
        If new_tab Then
            ' HTMLWindow2 には新規タブ Navigate がないため javascript で代替
            .execScript "window.open('" & navigate_url & "', '_blank');"
        Else
            .navigate navigate_url
        End If

        If wait_title <> "" Then
            ' ReadyState だけでは不足するため、タイトル一致待機を追加
            Call WaitNewWindowByHandle(Me.hwnd)
            Set m_parentWindow = m_htmlDoc.parentWindow   ' 念のためメンバー再同期
            Call InitializeByWindowHandle(Me.hwnd)
            Call WaitByTitleChange(wait_title)
        End If

        If do_wait Then
            NavigatePage = Me.Wait(timeout_milli_sec)
        Else
            NavigatePage = True
        End If
    End With
End Function


' =====================================================
'  現在起動している IE と Edge(IEモード) のウィンドウ情報を列挙
'   戻り値 : WindowInfo 型の配列
' =====================================================
Private Function GetWindowInfoArray() As WindowInfo()
    Dim winInfoArr() As WindowInfo
    Dim arrIndex As Long
    Dim hwnd As LongPtr
    Dim wClassName As String
    Dim targetClassName As String
    Dim processName As String
    Dim uiWindow As IUIAutomationElement

    If m_isExplore Then
        processName = "iexplore.exe"
        targetClassName = "IEFrame"
    Else
        processName = "msedge.exe"
        targetClassName = "Chrome_WidgetWin_1"
    End If

    ' 最初のウィンドウ
    hwnd = FindWindow(vbNullString, vbNullString)

    While hwnd <> 0
        ' 見えないウィンドウ/生成直後などはスキップ
        If IsWindowVisible(hwnd) = False Then GoTo Continue
        ' 仮想デスクトップでの隠蔽ウィンドウ等はスキップ
        If IsWindowCloaked(hwnd) = True Then GoTo Continue

        wClassName = GetWindowClassName(hwnd)
        If wClassName = targetClassName Then
            ' ここから UIA か WMI で絞り込み
            If GetProcessName(hwnd) = processName Then
                Set uiWindow = m_uia.ElementFromHandle(ByVal hwnd)

                ' Edge の場合、msedge.dll を含む UIA Provider のみ抽出
                ' iexplore.exe の場合はこの条件を外す
                If m_isExplore = False Then
                    If InStr(StrConv(uiWindow.CurrentProviderDescription, vbLowerCase), _
                             "msedge.dll") = 0 Then GoTo Continue
                End If

                If arrIndex = 0 Then
                    ReDim winInfoArr(arrIndex)
                Else
                    ReDim Preserve winInfoArr(arrIndex)
                End If

                With winInfoArr(arrIndex)
                    .wHandle = hwnd
                    .wClassName = wClassName
                    .wName = GetWindowName(hwnd)
                End With

                arrIndex = UBound(winInfoArr) + 1
            End If
        End If

Continue:
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)   ' 次の兄弟ウィンドウ
    Wend

    GetWindowInfoArray = winInfoArr
End Function


' =====================================================
'  指定ハンドルのウィンドウテキストを取得
'   hwnd : ウィンドウハンドル
'   戻値 : ウィンドウテキストの文字列
' =====================================================
Private Function GetWindowName(ByVal hwnd As LongPtr) As String
    Dim buf As String, nByte As Long
    buf = String$(255, vbNullChar)
    GetWindowText hwnd, buf, Len(buf)
    GetWindowName = Left$(buf, InStr(1, buf, vbNullChar) - 1)
End Function


' =====================================================
'  指定ハンドルのウィンドウクラスを取得
'   hwnd : ウィンドウハンドル
'   戻値 : ウィンドウクラスの文字列
' =====================================================
Private Function GetWindowClassName(ByVal hwnd As LongPtr) As String
    Dim buf As String * 255
    Call GetClassName(hwnd, buf, Len(buf))
    GetWindowClassName = Left$(buf, InStr(1, buf, vbNullChar) - 1)
End Function


' =====================================================
'  WindowsAPIにより Cloaked かを判定（別デスクトップの被せ等の検出に利用）
'   hwnd : ウィンドウハンドル
'   戻値 : Cloaked=True / Cloakedでない=False
' =====================================================
Private Function IsWindowCloaked(ByVal hwnd As LongPtr) As Boolean
    Dim cloaked As Long
    Dim hResult As Long

    hResult = DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, cloaked, LenB(cloaked))

    If hResult = S_OK Then
        If cloaked <> 0 Then
            IsWindowCloaked = True
        Else
            IsWindowCloaked = False
        End If
    Else
        ' APIが失敗すると正しく判定できないため、安全側で False
        IsWindowCloaked = False
    End If
End Function


' =====================================================
'  （備考）Pleasanter画面で「サイトから移動します？」ダイアログが出ないように、
'          タブを閉じるユーティリティ
'   keep_page_like_keyword : 残しておきたいページ名に含まれる語（例：イントラネット・システムメニュー）
' =====================================================
Public Function CleanIEWindow(Optional ByVal keep_page_like_keyword As String = "イントラネット・システムメニュー") As Boolean
    Dim winInfoAry() As WindowInfo
    winInfoAry = GetWindowInfoArray

    Dim i As Long
    For i = 0 To UBound(winInfoAry)
        Call CloseTabByName(winInfoAry(i).wHandle, keep_page_like_keyword)
    Next
End Function


' =====================================================
'  指定ウィンドウのタブを名前で閉じる
'   hwnd     : ウィンドウハンドル
'   page_name: タブ名（に含まれる語）
' =====================================================
Private Function CloseTabByName(ByVal hwnd As LongPtr, ByVal page_name As String) As Boolean
    Dim uiWindow As IUIAutomationElement
    Set uiWindow = m_uia.ElementFromHandle(ByVal hwnd)

    Dim uiTabItemAry As IUIAutomationElementArray

    ' Edge のIEモードではタブは UIA のタブ項目として取得できる。クラス名でフィルタするか？
    ' ここでは ControlType=TabItem で収集している。
    Set uiTabItemAry = GetUIElementArray(UIA_ControlTypePropertyId, UIA_TabItemControlTypeId, _
                                         , , uiWindow)

    ' TODO: 必要なら ClassName="EdgeTab" などのフィルタを追加する
    ' Debug.Print "tabs:", uiTabItemAry.Length
    If uiTabItemAry Is Nothing Or uiTabItemAry.Length = 0 Then
        ' タブがない＝単一ウィンドウと判断し、ウィンドウ自体を閉じる
        Call CloseWindow(hwnd)
        Exit Function
    End If

    Dim i As Long
    For i = 0 To uiTabItemAry.Length - 1
        ' Debug.Print "tab:", uiTabItemAry.GetElement(i).CurrentName
        If InStr(uiTabItemAry.GetElement(i).CurrentName, page_name) > 0 Then
            ' アクティブタブを閉じる処理に委譲
            Call CloseActiveTabWindow(hwnd)
        End If
    Next
End Function


' =====================================================
'  指定ハンドルからプロセス名を取得
' =====================================================
Private Function GetProcessName(ByVal hwnd As LongPtr) As String
    Dim con As Object
    Dim items As Object
    Dim pid As LongPtr

    Set con = CreateObject("WbemScripting.SWbemLocator").ConnectServer
    GetWindowThreadProcessId hwnd, pid

    Set items = con.ExecQuery("Select Name From Win32_Process Where ProcessId = " & pid & "")
    If items.Count = 1 Then
        GetProcessName = items.ItemIndex(0).Name
    End If
End Function


' =====================================================
'  IEモードウィンドウを閉じる（IE8 互換）
' =====================================================
Public Sub QuitPage()
    If HasTabWindow(Me.hwnd) Then
        ' Class_IEのQuitPageと同様にタブ単位で閉じる
        Call CloseActiveTabWindow(Me.hwnd)
    Else
        ' タブが無い場合はウィンドウごと閉じる
        Call CloseWindow(Me.hwnd)
    End If
End Sub


' =====================================================
'  対象ウィンドウにタブが含まれるかをチェック
' =====================================================
Private Function HasTabWindow(ByVal hwnd As LongPtr) As Boolean
    Dim uiWindow As IUIAutomationElement
    Set uiWindow = m_uia.ElementFromHandle(ByVal hwnd)
    If uiWindow Is Nothing Then
        Err.Raise 999, , "(HasTabWindow)" & vbCrLf & "UIWindowの取得に失敗しました。"
    End If

    Dim uiTab As IUIAutomationElement
    If m_isExplore Then
        Set uiTab = GetUIElement(UIA_ControlTypePropertyId, UIA_TabControlTypeId, _
                                 UIA_NamePropertyId, "タブ行", uiWindow)
    Else
        Set uiTab = GetUIElement(UIA_ControlTypePropertyId, UIA_TabControlTypeId, _
                                 UIA_NamePropertyId, "タブ バー", uiWindow)
    End If

    If Not uiTab Is Nothing Then
        HasTabWindow = True
    End If
End Function


' =====================================================
'  アクティブなタブを閉じる（タブが無い場合は失敗する）
'   ※ ここではタブコンテナの特定のみ（実際の「閉じる」操作は環境次第で要追記）
' =====================================================
Private Function CloseActiveTabWindow(ByVal hwnd As LongPtr) As Boolean
    Dim uiWindow As IUIAutomationElement
    Set uiWindow = m_uia.ElementFromHandle(ByVal hwnd)
    If uiWindow Is Nothing Then
        Err.Raise 999, , "(CloseActiveTabWindow)" & vbCrLf & "UIWindowの取得に失敗しました。"
    End If

    Dim uiTab As IUIAutomationElement
    If m_isExplore Then
        Set uiTab = GetUIElement(UIA_ControlTypePropertyId, UIA_TabControlTypeId, _
                                 UIA_NamePropertyId, "タブ行", uiWindow)
    Else
        Set uiTab = GetUIElement(UIA_ControlTypePropertyId, UIA_TabControlTypeId, _
                                 UIA_NamePropertyId, "タブ バー", uiWindow)
    End If

    ' TODO: タブコンテナ配下から「閉じる」ボタンを見つけ、InvokePattern を実行する処理を加える
    ' 今回の画像範囲ではここまで。暫定で成功扱いを返す。
End Function


' =====================================================
'  対象ウィンドウを閉じる（UIAutomation の WindowPattern.Close）
' =====================================================
Private Function CloseWindow(ByVal hwnd As LongPtr) As Boolean
    Dim uiWindow As IUIAutomationElement
    Dim uiWindowPattern As IUIAutomationWindowPattern

    If hwnd <= 0 Then Exit Function

    Set uiWindow = m_uia.ElementFromHandle(ByVal hwnd)
    If uiWindow Is Nothing Then
        Err.Raise 999, , "(CloseWindow)" & vbCrLf & "UIWindowの取得に失敗しました。"
    End If

    Set uiWindowPattern = uiWindow.GetCurrentPattern(UIA_PatternIds.UIA_WindowPatternId)
    If uiWindowPattern Is Nothing Then
        Err.Raise 999, , "(CloseWindow)" & vbCrLf & "WindowPatternの取得に失敗しました。"
    End If

    uiWindowPattern.Close

    If WaitUntilWindowClosed(hwnd) = False Then
        Err.Raise 999, , "(CloseWindow)" & vbCrLf & "ウィンドウを閉じることができませんでした。"
    End If

    CloseWindow = True
End Function


' =====================================================
'  指定した条件の UIA 要素を1件取得
'  property_id1/value1 は必須。property_id2/value2 は省略可。
'  parent_element 省略時は RootElement を対象。
' =====================================================
Private Function GetUIElement(ByVal property_id1 As Long, ByVal property_value1 As Variant, _
                              Optional ByVal property_id2 As Long = 0, Optional ByVal property_value2 As Variant = Empty, _
                              Optional ByRef parent_element As IUIAutomationElement = Nothing) As IUIAutomationElement
    Dim cnd1 As IUIAutomationCondition
    Dim cnd2 As IUIAutomationCondition
    Dim cndAll As IUIAutomationCondition

    Set cnd1 = m_uia.CreatePropertyCondition(property_id1, property_value1)
    If (property_id2 <> 0) And (property_value2 <> Empty) Then
        Set cnd2 = m_uia.CreatePropertyCondition(property_id2, property_value2)
        Set cndAll = m_uia.CreateAndCondition(cnd1, cnd2)
    Else
        Set cndAll = cnd1
    End If

    If parent_element Is Nothing Then
        Set parent_element = m_uia.GetRootElement
    End If

    Set GetUIElement = parent_element.FindFirst(TreeScope_Subtree, cndAll)
End Function


' =====================================================
'  指定した条件の UIA 要素を配列取得
' =====================================================
Private Function GetUIElementArray(ByVal property_id1 As Long, ByVal property_value1 As Variant, _
                                   Optional ByVal property_id2 As Long = 0, Optional ByVal property_value2 As Variant = Empty, _
                                   Optional ByRef parent_element As IUIAutomationElement = Nothing) As IUIAutomationElementArray
    Dim cnd1 As IUIAutomationCondition
    Dim cnd2 As IUIAutomationCondition
    Dim cndAll As IUIAutomationCondition

    Set cnd1 = m_uia.CreatePropertyCondition(property_id1, property_value1)
    If (property_id2 <> 0) And (property_value2 <> Empty) Then
        Set cnd2 = m_uia.CreatePropertyCondition(property_id2, property_value2)
        Set cndAll = m_uia.CreateAndCondition(cnd1, cnd2)
    Else
        Set cndAll = cnd1
    End If

    If parent_element Is Nothing Then
        Set parent_element = m_uia.GetRootElement
    End If

    Set GetUIElementArray = parent_element.FindAll(TreeScope_Subtree, cndAll)
End Function


' =====================================================
'  ウィンドウを取得する関数（ラッパー）
'   title_element     : 検索キーワード（タイトル or URL）
'   wait_new_window   : 新規ウィンドウが出たハンドルを待つか（既定 False）
'   full_match        : True=完全一致 / False=部分一致
'   wait_after        : 取得後に待機（既定 False）
'   loop_mode         : IEの挙動に応じてループするか（既定 True）
'   timeout_milli_sec : タイムアウト（ミリ秒、0で無制限 ※タイムアウトなし）
'   戻値              : 成功=True / 失敗=False
' =====================================================
Public Function LoadWindow(ByVal title_element As String, _
                           Optional wait_new_window As Boolean = False, _
                           Optional full_match As Boolean = False, _
                           Optional wait_after As Boolean = False, _
                           Optional loop_mode As Boolean = True, _
                           Optional timeout_milli_sec As Long = 0) As Boolean

    ' TODO: ブラウザが1つも起動していない場合にも対応できているか？
    If wait_new_window Then
        Call WaitNewWindow(title_element, full_match, timeout_milli_sec)
    End If

    LoadWindow = LoadWindowByTitleOrURL(title_element, full_match, wait_after, loop_mode, timeout_milli_sec)
End Function


' =====================================================
'  LoadWindow のサブルーチン。IEobj を取得。
'   target_keyword  : IEobj の特定用キーワード（タイトル or URL）
'   full_match      : True=完全一致
'   wait_after      : 取得後に待つか
'   loop_mode       : 取得できなければリトライするか
'   timeout_milli_sec: タイムアウト（ミリ秒）
' =====================================================
Private Function LoadWindowByTitleOrURL(ByVal target_keyword As String, ByVal full_match As Boolean, _
                                        ByVal wait_after As Boolean, ByVal loop_mode As Boolean, ByVal timeout_milli_sec As Long) As Boolean
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim returnHwnd As LongPtr
    Dim startTime As Double, timeDiff As Double
    startTime = Timer()

Continue:
    Set htmlDoc = GetHtmlDocument(target_keyword, full_match, returnHwnd)

    If Not htmlDoc Is Nothing Then
        ' メンバー変数の初期化
        Call InitializeByWindowHandle(returnHwnd)

        If wait_after Then
            Call Me.Wait(timeout_milli_sec)
        End If

        LoadWindowByTitleOrURL = True
        Exit Function
    End If

    If loop_mode And (htmlDoc Is Nothing) Then
        If timeout_milli_sec > 0 Then
            timeDiff = (Timer() - startTime) * 1000#
            If timeDiff >= timeout_milli_sec Then
                Debug.Print "timeout:", timeDiff
                Exit Function
            End If
        End If

        DoEvents
        GoTo Continue
    End If
End Function


' =====================================================
'  ページから一意の要素（id / name / class / tag）を取得する
'   keyword      : 要素のid, name, ClassName, TagNameを指定する文字列
'   element_type : HtmlUniqueElement の種別
'                  TYPE_ID       id指定。この場合 element_index は無視
'                  TYPE_NAME     name指定。要素コレクションの index を使用
'                  TYPE_CLASS_NAME ClassName指定。要素コレクションの index を使用
'                  TYPE_TAG_NAME   TagName指定。要素コレクションの index を使用
'   parent_elem  : 検索を開始する親の HTML要素/ドキュメント（既定: m_htmlDoc）
'   element_index: コレクションの index（既定 0）
'   戻り値       : 見つかった HTMLエレメント（Object型）/ 失敗時 Nothing
' =====================================================
Public Function GetUniqueElement(ByVal keyword As String, _
                                 ByVal element_type As HtmlUniqueElement, _
                                 Optional parent_elem As Object = Nothing, _
                                 Optional element_index As Long = 0) As Object

    If (parent_elem Is Nothing) Then Set parent_elem = m_htmlDoc
    Set GetUniqueElement = Nothing

    Select Case element_type
        Case HtmlUniqueElement.TYPE_ID
            If Not (IsNull(parent_elem.getElementById(keyword))) Then
                Set GetUniqueElement = parent_elem.getElementById(keyword)
            End If

        Case HtmlUniqueElement.TYPE_NAME
            If Not (parent_elem.getElementsByName(keyword)(element_index) Is Nothing) Then
                Set GetUniqueElement = parent_elem.getElementsByName(keyword)(element_index)
            End If

        Case HtmlUniqueElement.TYPE_CLASS_NAME
            If Not (parent_elem.getElementsByClassName(keyword)(element_index) Is Nothing) Then
                Set GetUniqueElement = parent_elem.getElementsByClassName(keyword)(element_index)
            End If

        Case HtmlUniqueElement.TYPE_TAG_NAME
            If Not (parent_elem.getElementsByTagName(keyword)(element_index) Is Nothing) Then
                Set GetUniqueElement = parent_elem.getElementsByTagName(keyword)(element_index)
            End If
    End Select
End Function


' =====================================================
'  指定したタグ配下から、属性値/InnerTextで目的の要素を取得する（基本版）
'   tag_name          : 対象タグ
'   parent_elem       : 検索を開始するHTMLエレメント/ドキュメント（省略時: m_htmlDoc）
'   attribute_keyword : 属性に含まれる文字列（既定: Empty）
'   inner_text_keyword: InnerTextに含まれる文字列（既定: Empty）
'   target_count      : N番目に一致した要素を返す（既定: 1）
'   nohavechild_only  : 子タグを持たない要素のみ対象にする（既定: False）
'   戻り値            : 見つかったHTMLエレメント（Object）/ 失敗時 Nothing
' =====================================================
Public Function GetElementInTagCollection(ByVal tag_name As String, _
                                          Optional parent_elem As Object = Nothing, _
                                          Optional attribute_keyword As String = Empty, _
                                          Optional inner_text_keyword As String = Empty, _
                                          Optional target_count As Long = 1, _
                                          Optional nohavechild_only As Boolean = False) As Object
    Set GetElementInTagCollection = Nothing
    If parent_elem Is Nothing Then Set parent_elem = m_htmlDoc

    Dim currentElement As Object
    Dim attributeText As String
    Dim Match(0 To 1) As Boolean
    Dim foundCount As Long

    Erase Match
    foundCount = 0

    For Each currentElement In parent_elem.getElementsByTagName(tag_name)

        ' 子要素の有無でフィルタ
        If nohavechild_only Then
            If currentElement.getElementsByTagName(tag_name).Length > 0 Then GoTo Continue
        End If

        ' 属性部分の文字列抽出（outerHTMLの先頭タグ部）
        If (attribute_keyword = Empty) Then
            Match(0) = True
        Else
            If InStr(currentElement.outerHTML, ">") > 0 Then
                attributeText = Mid(currentElement.outerHTML, 2, InStr(currentElement.outerHTML, ">") - 2)
            Else
                attributeText = currentElement.outerHTML
            End If

            If InStr(1, attributeText, attribute_keyword, vbTextCompare) > 0 Then Match(0) = True
        End If

        ' InnerText の一致確認
        If (inner_text_keyword = Empty) Then
            Match(1) = True
        Else
            If InStr(1, currentElement.innerText, inner_text_keyword, vbTextCompare) > 0 Then Match(1) = True
        End If

        ' 両方一致でカウント
        If Match(0) And Match(1) Then
            foundCount = foundCount + 1
            If (foundCount >= target_count) Then
                Set GetElementInTagCollection = currentElement
                Exit Function
            End If
        Else
            Erase Match
        End If

Continue:
    Next
End Function


' =====================================================
'  属性/InnerText/完全一致など複数パターンでタグ要素を検索する（拡張版）
'   tag_name       : タグ名
'   parent_elem    : 検索を開始するエレメント/ドキュメント（省略時: m_htmlDoc）
'   keyword1       : 1つ目のキーワード（Empty可）
'   keyword2       : 2つ目のキーワード（Empty可）
'   match_pattern  : MatchPattern 列挙（LIKE_BOTH が既定）
'                    LIKE_BOTH      … 属性/InnerText の両方を部分一致
'                    LIKE_ATTRIBUTE … 属性だけ部分一致
'                    PERFECT_INNER  … InnerText を完全一致
'                    LIKE_INNER     … InnerText を部分一致
'   target_count   : N番目の一致要素を返す（既定: 1）
'   nohavechild_only: 子タグのない要素のみ対象（既定: False）
'   戻り値         : 見つかったHTMLエレメント（Object）/ 失敗時 Nothing
' =====================================================
Public Function GetElementInTagCollectionEx(ByVal tag_name As String, _
                                            Optional parent_elem As Object = Nothing, _
                                            Optional keyword1 As String = Empty, _
                                            Optional keyword2 As String = Empty, _
                                            Optional match_pattern As MatchPattern = MatchPattern.LIKE_BOTH, _
                                            Optional target_count As Long = 1, _
                                            Optional nohavechild_only As Boolean = False) As Object
    If parent_elem Is Nothing Then Set parent_elem = m_htmlDoc
    Set GetElementInTagCollectionEx = Nothing
    If keyword1 = Empty And keyword2 = Empty Then Exit Function

    Dim foundCount As Long: foundCount = 0
    Dim currentElement As Object
    Dim strAttribute As String, strInnerText As String
    Dim Match(0 To 1) As Boolean

    Erase Match

    For Each currentElement In parent_elem.getElementsByTagName(tag_name)
        If nohavechild_only Then
            If currentElement.getElementsByTagName(tag_name).Length > 0 Then GoTo Continue
        End If

        ' Attribute の文字列
        strAttribute = Mid(currentElement.outerHTML, 2, InStr(currentElement.outerHTML, ">") - 2)

        ' InnerText を outerHTML から切り出す（簡易）
        Dim cutStartPos As Long, cutEndPos As Long
        cutStartPos = InStr(2, currentElement.outerHTML, ">")
        If cutStartPos = 0 Then
            strInnerText = Empty
        Else
            cutEndPos = InStr(cutStartPos + 1, currentElement.outerHTML, "<")
            If cutEndPos = 0 Then
                strInnerText = Empty
            Else
                strInnerText = Replace(Mid(currentElement.outerHTML, cutStartPos + 1, cutEndPos - cutStartPos - 1), "&nbsp;", " ")
            End If
        End If

        ' 属性側の一致
        Select Case match_pattern
            Case MatchPattern.LIKE_BOTH, MatchPattern.LIKE_ATTRIBUTE
                If InStr(1, strAttribute, keyword1, vbTextCompare) > 0 Then Match(0) = True
                If InStr(1, strAttribute, keyword2, vbTextCompare) > 0 Then Match(0) = True
        End Select

        ' InnerText 側の一致
        Select Case match_pattern
            Case MatchPattern.LIKE_BOTH, MatchPattern.LIKE_INNER
                If InStr(1, strInnerText, keyword1, vbTextCompare) > 0 Then Match(1) = True
                If InStr(1, strInnerText, keyword2, vbTextCompare) > 0 Then Match(1) = True

            Case MatchPattern.PERFECT_INNER
                If StrComp(strInnerText, keyword1, vbTextCompare) = 0 Then Match(1) = True
                If StrComp(strInnerText, keyword2, vbTextCompare) = 0 Then Match(1) = True
        End Select

        If (keyword1 = Empty) Then Match(0) = True
        If (keyword2 = Empty) Then Match(1) = True

        If Match(0) And Match(1) Then
            foundCount = foundCount + 1
            If foundCount >= target_count Then
                Set GetElementInTagCollectionEx = currentElement
                Exit Function
            End If
        End If

        Erase Match
Continue:
    Next
End Function


' =====================================================
'  IEの読み込み待ち
'   timeout_mili_sec      : タイムアウト（ミリ秒。0=無制限）
'   wait_before           : Unloadイベントが完了するのを事前に待つ（既定: False）
'   before_timeout_mili_sec: 事前待ちのタイムアウト（既定: TIMEOUT_NORMAL）
' =====================================================
Public Function Wait(Optional timeout_mili_sec As Long = 0, _
                     Optional wait_before As Boolean = False, _
                     Optional before_timeout_mili_sec As Long = TIMEOUT_NORMAL) As Boolean
    Dim startTime As Double, timeDiff As Double
    startTime = Timer()

    If wait_before Then
        Do While Not (m_onUnloadEvent)
            DoEvents
            timeDiff = (Timer() - startTime) * 1000#
            If m_timeoutListener Then Debug.Print "timeoutcheck:", timeDiff
            If m_timeoutOutput Then m_timeoutItems.Add Now & "[timeoutcheck]" & timeDiff
            If (timeDiff > before_timeout_mili_sec) Then Exit Do
        Loop
    End If

    Wait = False

retry:
    If (WaitComplete(m_htmlDoc, timeout_mili_sec)) Then
        If (WaitEachIFrame(timeout_mili_sec)) Then
            Wait = True
        Else
            GoTo retry
        End If
    End If
End Function


' =====================================================
'  指定タイトルによる新規ウィンドウの待機
'   new_window_title : 対象ウィンドウタイトル
'   full_match       : 完全一致フラグ（既定: False）
'   wait_window_title: もう一方のタイトルの指定（省略可）
'   timeout_mili_sec : タイムアウト（既定: TIMEOUT_NORMAL）
' =====================================================
Private Function WaitNewWindow(ByVal new_window_title As String, _
                               Optional ByVal full_match As Boolean = False, _
                               Optional ByVal wait_window_title As String = Empty, _
                               Optional ByVal timeout_mili_sec As Long = TIMEOUT_NORMAL) As Boolean
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim startTime As Double, timeDiff As Double: startTime = Timer()
    Dim typName As String

    Do
        Sleep 10
        DoEvents
        timeDiff = (Timer() - startTime) * 1000#
        If m_timeoutListener Then Debug.Print "timeoutcheck:", timeDiff
        If m_timeoutOutput Then m_timeoutItems.Add Now & "[timeoutcheck]" & timeDiff
        If timeout_mili_sec > 0 Then
            If timeDiff >= timeout_mili_sec Then
                WaitNewWindow = False
                Exit Function
            End If
        End If

        ' 取得途中の型不整合でオートメーション例外が出ることがあるため握りつぶす
        On Error Resume Next
        Set htmlDoc = GetHtmlDocument(new_window_title, full_match)
        If Err.Number <> 0 Then
            Debug.Print "(WaitNewWindow) エラー発生:", Err.Number, Err.Description
            Set htmlDoc = Nothing
        End If
        On Error GoTo 0

        typName = TypeName(htmlDoc)
        Loop Until (typName = "HTMLDocument")  ' 取得できたらループ脱出

    WaitNewWindow = True
End Function


' =====================================================
'  共用：既定の読み込み完了待ち
'   html_Doc        : 対象の HTMLDocument（ByRef）
'   timeout_mili_sec: タイムアウト（ミリ秒。0=無制限）
' =====================================================
Private Function WaitComplete(ByRef html_Doc As HTMLDocument, _
                              Optional ByVal timeout_mili_sec As Long = 0) As Boolean
    Dim startTime As Double, timeDiff As Double
    Dim completeCount As Long

    startTime = Timer()
    Do
        DoEvents
        Sleep 100

        timeDiff = (Timer() - startTime) * 1000#
        If m_timeoutListener Then Debug.Print "timeoutcheck:", timeDiff
        If m_timeoutOutput Then m_timeoutItems.Add Now & "[timeoutcheck]" & timeDiff

        If (timeout_mili_sec <> 0) Then
            If timeDiff >= timeout_mili_sec Then
                WaitComplete = False
                Exit Function
            End If
        End If

        If html_Doc.ReadyState = "complete" Then
            completeCount = completeCount + 1
            Debug.Print "completeCount:", completeCount
            Sleep 100
        End If

        ' 3回連続 complete を確認するまでループ
    Loop Until completeCount >= 3 _
          Or (m_onBeforeDeactivateEvent And (html_Doc.ReadyState = "complete") And (m_onLoadEvent))

    WaitComplete = True
End Function



' =====================================================
'  すべての iFrame が完了するまで待つ
'   timeout_mili_sec : タイムアウト（既定 0=無制限）
' =====================================================
Private Function WaitEachIFrame(Optional timeout_mili_sec As Long = 0) As Boolean
    Dim i As Long, iframeHtmlDoc As HTMLDocument, iframe As HTMLIFrame
    Dim startTime As Double, timeDiff As Double
    Dim redoFlag As Boolean

    WaitEachIFrame = True

    Do
        redoFlag = False
        If m_htmlDoc.getElementsByTagName("iframe").Length = 0 Then Exit Function

        For i = 0 To m_htmlDoc.getElementsByTagName("iframe").Length - 1
            If Not (WaitComplete(m_htmlDoc, timeout_mili_sec)) Then
                WaitEachIFrame = False
                Exit Function
            End If

            On Error Resume Next
            Set iframe = m_htmlDoc.getElementsByTagName("iframe")(i)
            Set iframeHtmlDoc = iframe.contentWindow.document
            On Error GoTo 0

            If Not (iframeHtmlDoc Is Nothing) Then
                startTime = Timer()
                Do
                    DoEvents
                    Do Until (m_htmlDoc.ReadyState = "complete") And (iframeHtmlDoc.ReadyState = "complete")
                    Loop
                    Sleep 10

                    timeDiff = GetTimerLap(startTime) * 1000#
                    If m_timeoutListener Then Debug.Print "timeoutcheck iframe:", timeDiff
                    If m_timeoutOutput Then m_timeoutItems.Add Now & "[timeoutcheck iframe]" & timeDiff

                    If Not (timeout_mili_sec = 0) Then
                        If timeDiff >= timeout_mili_sec Then
                            WaitEachIFrame = False
                            Exit Function
                        End If
                    End If
                Loop
            Else
                redoFlag = True
            End If

            Set iframeHtmlDoc = Nothing
        Next
    Loop Until Not (redoFlag)
End Function


' =====================================================
'  Location/Title を監視してタイトルが所望の値になるまで待つ
'   target_title     : 目標タイトル（URLの場合は LocationURL に含まれるかで判定）
'   error_title      : エラー時に現れるタイトル（見つけたら 1 を返して終了）
'   timeout_mili_sec : タイムアウト（0=無制限）
'   戻り値: 1=エラー検出 / 0=タイムアウト / -1=継続中（通常は戻らない設計）
' =====================================================
Public Function WaitByTitleChange(ByVal target_title As String, _
                                  Optional error_title As String = Empty, _
                                  Optional timeout_mili_sec As Long = 0) As Long
    Dim startTime As Double, timeDiff As Double
    startTime = Timer()
    WaitByTitleChange = -1

    While True
        If IsURL(target_title) And (InStr(Me.LocationURL, target_title) > 0) Then Exit Function
        If Not (IsURL(target_title)) And (InStr(m_htmlDoc.title, target_title) > 0) Then Exit Function

        Sleep 10

        If error_title <> Empty Then
            If InStr(m_htmlDoc.title, error_title) > 0 Then
                WaitByTitleChange = 1
                Exit Function
            End If
        End If

        timeDiff = GetTimerLap(startTime) * 1000#
        If m_timeoutListener Then Debug.Print "timeoutcheck:", timeDiff
        If m_timeoutOutput Then m_timeoutItems.Add Now & "[timeoutcheck]" & timeDiff

        If timeout_mili_sec > 0 Then
            If timeDiff >= timeout_mili_sec Then
                WaitByTitleChange = 0
                Exit Function
            End If
        End If
    Wend
End Function


' =====================================================
'  新規ウィンドウ（ハンドル）で HTMLDocument 取得可能になるまで待つ
' =====================================================
Private Function WaitNewWindowByHandle(ByVal hwnd As LongPtr, _
                                       Optional ByVal timeout_mili_sec As Long = TIMEOUT_NORMAL) As Boolean
    Const procName = "(WaitNewWindowByHandle)"
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim startTime As Double, timeDiff As Double

    startTime = Timer()
    Do
        DoEvents
        Sleep 10

        timeDiff = GetTimerLap(startTime) * 1000#
        If m_timeoutListener Then Debug.Print "timeoutcheck:", timeDiff
        If m_timeoutOutput Then m_timeoutItems.Add Now & "[timeoutcheck]" & timeDiff
        If timeout_mili_sec > 0 Then
            If timeDiff >= timeout_mili_sec Then
                WaitNewWindowByHandle = False
                Exit Function
            End If
        End If

        If IsWindow(hwnd) = 0 Then
            ' 一時ハンドルのように hWnd が消失する場合は終了
            Exit Function
        End If

        Set htmlDoc = GetHtmlDocumentByHandle(hwnd)
        Loop Until (TypeName(htmlDoc) = "HTMLDocument")  ' 取得できるまで待機

    If htmlDoc Is Nothing Then
        Err.Raise 999, procName & vbCrLf & "対象のHTMLDocumentが見つかりませんでした。" & vbCrLf & _
                       "対象ハンドル: " & hwnd
    End If

    WaitNewWindowByHandle = True
End Function


' =====================================================
'  ウィンドウが閉じるまで待機
' =====================================================
Private Function WaitUntilWindowClosed(ByVal target_hWnd As LongPtr, _
                                       Optional ByVal timeout_mili_sec As Long = 10000) As Boolean
    Dim isExisted As Boolean
    Dim startTime As Double, timeDiff As Double

    startTime = Timer()
    Do
        DoEvents
        isExisted = (IsWindow(target_hWnd) <> 0)

        timeDiff = GetTimerLap(startTime) * 1000#
        If timeDiff >= timeout_mili_sec Then
            Debug.Print "WaitUntilWindowClosed タイムアウト:", timeout_mili_sec
            Exit Function
        End If
    Loop While isExisted

    WaitUntilWindowClosed = True
End Function


' =====================================================
'  UIA要素が無効化されるまで待つ（要素の CurrentName 取得で例外→無効化とみなす）
' =====================================================
Private Function WaitUntilDisabledUIElement(ByRef ui_element As IUIAutomationElement, _
                                            Optional ByVal timeout_mili_sec As Long = 5000) As Boolean
    Dim tmpName As String
    Dim startTime As Double, timeDiff As Double

    startTime = Timer()
    Do
        On Error Resume Next
        tmpName = ui_element.CurrentName
        If Err.Number <> 0 Then
            Debug.Print Err.Number, Err.Description
            ' Officeバージョン依存のUIAutomation例外（0x80040201など）で失敗＝無効化の可能性
            WaitUntilDisabledUIElement = True
            Exit Function
        End If
        On Error GoTo 0

        timeDiff = GetTimerLap(startTime) * 1000#
    Loop Until timeDiff >= timeout_mili_sec
End Function


' =====================================================
'  Timer() の日付跨ぎ補正付き経過秒
' =====================================================
Private Function GetTimerLap(ByVal start_time As Double) As Double
    GetTimerLap = Timer() - start_time
    GetTimerLap = IIf(GetTimerLap < 0, GetTimerLap + CDbl(86400), GetTimerLap)
End Function


' =====================================================
'  ノード間を辿って、指定 InnerText を起点に end_tagname まで上昇して該当ノードを返す
'   start_inner_text : 起点のInnerText
'   tag_name         : 起点のタグ名
'   refine_tag_name  : 検索範囲を限定する親タグ（任意）
'   parent_elem      : 起点とする親ノード（既定 m_htmlDoc）
' =====================================================
Public Function NodeClimber(ByVal start_inner_text As String, ByVal tag_name As String, _
                            Optional refine_tag_name As String = "", _
                            Optional parent_elem As Object = Nothing) As Object
    Set NodeClimber = Nothing
    If parent_elem Is Nothing Then Set parent_elem = m_htmlDoc
    Dim parentNode As Object

    If refine_tag_name <> "" Then
        Set parentNode = parent_elem.getElementsByTagName(refine_tag_name)
    Else
        Set parentNode = parent_elem.all
    End If

    Dim childNode As Object, currentNode As Object
    For Each childNode In parentNode
        If childNode.innerText = start_inner_text Then
            Set currentNode = childNode
            Do Until (currentNode.tagName = tag_name)
                Set currentNode = currentNode.parentElement
                If currentNode Is Nothing Then Exit Do
            Loop
            Set NodeClimber = currentNode
            Exit Function
        End If
    Next
End Function


' =====================================================
'  文字列がURLかを簡易判定
' =====================================================
Private Function IsURL(ByVal Source As String) As Boolean
    Select Case True
        Case (Left$(Source, 4) = "http")
            IsURL = True
        Case Else
            IsURL = False
    End Select
End Function

' URLエンコード／デコード
Private Function URLDecode(ByVal encoded_text As String) As String
    With CreateObject("ScriptControl")
        .Language = "JScript"
        URLDecode = .CodeObject.decodeURI(encoded_text)
    End With
End Function

Private Function URLEncode(ByVal url As String) As String
    ' Excel 2013 以降なら WorksheetFunction.EncodeURL が使える
    On Error Resume Next
    URLEncode = WorksheetFunction.EncodeURL(url)
    On Error GoTo 0
    If Len(URLEncode) = 0 Then
        With CreateObject("ScriptControl")
            .Language = "JScript"
            URLEncode = .CodeObject.encodeURI(url)
        End With
    End If
End Function


'' =====================================================
''  ログ出力（イベント／タイムアウト）の簡易エクスポート
'' =====================================================
'Private Sub OutputExporter()
'    Dim logBook As Workbook
'    With Application
'        .ScreenUpdating = False
'        .DisplayAlerts = False
'        Set logBook = Workbooks.Add
'    End With
'
'    If m_eventOutput = True Then Call OutputRecorder(logBook, m_eventItems, "IEイベントログ")
'    If m_timeoutOutput = True Then Call OutputRecorder(logBook, m_timeoutItems, "待機ログ")
'
'    logBook.SaveAs ThisWorkbook.Path & "\" & Replace(m_setTime, ":", "") & "_log.xlsx"
'    logBook.Close False
'
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'End Sub


' =====================================================
'  HTML特殊文字の空白/改行などの正規化
'   delete_linefeed : 改行を削除（既定 False）
'   delete_space    : 連続スペースやノーブレークスペースを1つに圧縮（既定 False）
' =====================================================
Private Function CleanHtmlChar(ByVal Source As String, Optional delete_linefeed As Boolean = False, _
                               Optional delete_space As Boolean = False) As String
    Dim newStr As String, nbsp(0 To 1) As Byte
    Dim zwsp(0 To 1) As Byte
    Dim strChar(1 To 5, 1 To 2) As String, i As Long

    newStr = Source

    ' &nbsp; → Chr(160) / Chr(32) など
    nbsp(0) = 160: nbsp(1) = 0
    newStr = Replace(newStr, nbsp, Empty)

    zwsp(0) = 11: zwsp(1) = 32
    newStr = Replace(newStr, zwsp, Empty)

    strChar(1, 1) = "&nbsp;":    strChar(1, 2) = " "
    strChar(2, 1) = "&emsp;":    strChar(2, 2) = " "
    strChar(3, 1) = "&lt;":      strChar(3, 2) = "<"
    strChar(4, 1) = "&gt;":      strChar(4, 2) = ">"
    strChar(5, 1) = "&quot;":    strChar(5, 2) = """"

    For i = LBound(strChar) To UBound(strChar)
        newStr = Replace(newStr, strChar(i, 1), strChar(i, 2))
    Next

    If delete_space Then
        newStr = Replace(newStr, vbTab, " ")
        newStr = Replace(newStr, "  ", " ")
    End If

    If delete_linefeed Then
        newStr = Replace(newStr, vbCrLf, Empty)
        newStr = Replace(newStr, vbCr, Empty)
        newStr = Replace(newStr, vbLf, Empty)
    End If

    CleanHtmlChar = newStr
End Function


' =====================================================
'  配列の要素数を取得（多次元対応）
'   失敗時は 0
' =====================================================
Private Function ArrayLength(ByVal arr As Variant, Optional ByVal dimension As Long = 1) As Long
    On Error GoTo Lbl_Err
    If IsArray(arr) Then
        ArrayLength = UBound(arr, dimension) + (1 - LBound(arr, dimension))
    Else
        ArrayLength = -1
    End If
    Exit Function
Lbl_Err:
    If Err.Number = 9 Then
        ArrayLength = 0
    End If
End Function


'' ============================================
'  ========= HTMLWindow2 / HTMLDocument のイベント =========
' ============================================
Private Sub m_parentWindow_onunload()
    m_onUnloadEvent = True
    m_onLoadEvent = False
    m_onBeforeDeactivateEvent = False
    If m_eventListener Then
        Debug.Print "onunload"
        If Not m_htmlDoc Is Nothing Then Debug.Print m_htmlDoc.title
    End If
    If m_eventOutput Then m_eventItems.Add "onunload | " & Now
End Sub

Private Sub m_parentWindow_onload()
    m_onLoadEvent = True
    m_onUnloadEvent = False
    If m_eventListener Then
        Debug.Print "onload"
        If Not m_htmlDoc Is Nothing Then Debug.Print m_htmlDoc.title
    End If
    If m_eventOutput Then m_eventItems.Add "onload | " & Now
End Sub

Private Sub m_htmlDoc_onreadystatechange()
    If m_htmlDoc.ReadyState = "complete" Then m_onUnloadEvent = False
    If m_eventListener Then
        Debug.Print m_htmlDoc.ReadyState
        If Not m_htmlDoc Is Nothing Then Debug.Print m_htmlDoc.title
    End If
    If m_eventOutput Then m_eventItems.Add "complete | " & Now
End Sub

Private Function m_htmlDoc_onbeforedeactivate() As Boolean
    m_onBeforeDeactivateEvent = True
    If m_eventListener Then
        Debug.Print "onbeforedeactivate"
        If Not m_htmlDoc Is Nothing Then Debug.Print m_htmlDoc.title
    End If
    If m_eventOutput Then m_eventItems.Add "onbeforedeactivate | " & Now
End Function


