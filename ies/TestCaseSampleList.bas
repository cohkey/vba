Option Explicit

' --- ヘルパー関数群 ---

' msedge.exe を Shell で起動し、ShellWindows から IE モードのウィンドウを探す
Private Function LaunchEdge(ByVal url As String, Optional ByVal timeoutSeconds As Single = 10) As SHDocVw.InternetExplorer
    Dim msedgePath As String
    msedgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" ' 環境に合わせて変更
    Dim processID As Long
    processID = Shell(msedgePath & " """ & url & """", vbNormalFocus)

    Dim sws As New SHDocVw.ShellWindows
    Dim ie As SHDocVw.InternetExplorer
    Dim t0 As Single: t0 = Timer
    Do
        For Each ie In sws
            ' URL が含まれていれば対象とする（空の場合は任意の IE モードウィンドウ）
            If url = "" Or InStr(ie.LocationURL, url) > 0 Then
                Set LaunchEdge = ie
                Exit Function
            End If
        Next ie
        DoEvents
        If Timer - t0 > timeoutSeconds Then Exit Do
    Loop
    ' 見つからなかった場合は Nothing
    Set LaunchEdge = Nothing
End Function

' IE の Document.readyState が "complete" になるのを待つ（タイムアウト付き）
Private Function WaitForDocumentComplete(ByVal ie As SHDocVw.InternetExplorer, Optional ByVal timeoutSeconds As Single = 10) As Boolean
    Dim startTime As Single: startTime = Timer
    Do While ie.Document.readyState <> "complete"
        DoEvents
        If Timer - startTime > timeoutSeconds Then
            WaitForDocumentComplete = False
            Exit Function
        End If
    Loop
    WaitForDocumentComplete = True
End Function

' URL エンコード（簡易版）
Public Function URLEncode(ByVal str As String) As String
    Dim i As Long, c As String, encoded As String
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        Select Case Asc(c)
            Case 48 To 57, 65 To 90, 97 To 122
                encoded = encoded & c
            Case Else
                encoded = encoded & "%" & Hex(Asc(c))
        End Select
    Next i
    URLEncode = encoded
End Function

' --- テストケース群 ---

' 1. 正常な URL へのナビゲーション
Public Sub Test_EdgeBrowser_NormalNavigation()
    Debug.Print "Running Test_EdgeBrowser_NormalNavigation"
    Dim url As String: url = "https://www.example.com"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "NormalNavigation: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "NormalNavigation: NG (Document load timeout)"
        ie.Quit: Exit Sub
    End If
    ' 例: ページタイトル "Example Domain" を検証
    Call AssertEqual("NormalNavigation Title", "Example Domain", ie.Document.Title)
    ie.Quit
End Sub

' 2. 同一ウィンドウで複数回のナビゲーション
Public Sub Test_EdgeBrowser_MultipleNavigation()
    Debug.Print "Running Test_EdgeBrowser_MultipleNavigation"
    Dim url1 As String, url2 As String
    url1 = "https://www.example.com"
    url2 = "https://www.iana.org/domains/reserved"

    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url1)
    If ie Is Nothing Then
        Debug.Print "MultipleNavigation: NG (IE window not found for first URL)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "MultipleNavigation: NG (First document load timeout)"
        ie.Quit: Exit Sub
    End If
    Call AssertEqual("MultipleNavigation Title 1", "Example Domain", ie.Document.Title)

    ' 同一ウィンドウで別 URL へナビゲーション
    ie.Navigate url2
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "MultipleNavigation: NG (Second document load timeout)"
        ie.Quit: Exit Sub
    End If
    Call AssertEqual("MultipleNavigation Title 2", "IANA — IANA-managed Reserved Domains", ie.Document.Title)
    ie.Quit
End Sub

' 3. 無効な URL 入力
Public Sub Test_EdgeBrowser_InvalidURL()
    Debug.Print "Running Test_EdgeBrowser_InvalidURL"
    Dim url As String: url = "htp://invalid-url"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url, 5) ' 短いタイムアウトで試行
    If ie Is Nothing Then
        Debug.Print "InvalidURL: OK (IE window not found as expected)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie, 5) Then
        Debug.Print "InvalidURL: OK (Document load timeout as expected)"
        ie.Quit: Exit Sub
    Else
        Debug.Print "InvalidURL: NG (Document unexpectedly loaded)"
    End If
    ie.Quit
End Sub

' 4. 空文字列の URL
Public Sub Test_EdgeBrowser_EmptyURL()
    Debug.Print "Running Test_EdgeBrowser_EmptyURL"
    Dim url As String: url = ""
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url, 5)
    If ie Is Nothing Then
        Debug.Print "EmptyURL: OK (No IE window launched for empty URL)"
    Else
        Debug.Print "EmptyURL: NG (IE window launched unexpectedly for empty URL)"
        ie.Quit
    End If
End Sub

' 5. ドキュメントの読み込みタイムアウトのシミュレーション
Public Sub Test_EdgeBrowser_DocumentLoadTimeout()
    Debug.Print "Running Test_EdgeBrowser_DocumentLoadTimeout"
    Dim url As String: url = "https://www.example.com"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "DocumentLoadTimeout: NG (IE window not found)"
        Exit Sub
    End If
    ' 故意に短いタイムアウト（1秒）を指定してタイムアウト発生を確認
    If Not WaitForDocumentComplete(ie, 1) Then
        Debug.Print "DocumentLoadTimeout: OK (Timeout occurred as expected)"
        ie.Quit: Exit Sub
    Else
        Debug.Print "DocumentLoadTimeout: NG (Document loaded within short timeout unexpectedly)"
    End If
    ie.Quit
End Sub

' 6. リダイレクト処理の検証
Public Sub Test_EdgeBrowser_RedirectHandling()
    Debug.Print "Running Test_EdgeBrowser_RedirectHandling"
    ' 例: http://example.com は https://www.example.com へリダイレクトする
    Dim url As String: url = "http://example.com"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "RedirectHandling: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "RedirectHandling: NG (Document load timeout)"
        ie.Quit: Exit Sub
    End If
    Call AssertEqual("RedirectHandling Title", "Example Domain", ie.Document.Title)
    ie.Quit
End Sub

' 7. 認証が必要なページへのアクセス
Public Sub Test_EdgeBrowser_AuthenticationPage()
    Debug.Print "Running Test_EdgeBrowser_AuthenticationPage"
    ' 例: httpbin.org の basic-auth ページ（※認証ダイアログが出る可能性があるので注意）
    Dim url As String: url = "https://httpbin.org/basic-auth/user/passwd"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url, 10)
    If ie Is Nothing Then
        Debug.Print "AuthenticationPage: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie, 10) Then
        Debug.Print "AuthenticationPage: SKIP (Authentication dialog likely blocked automation)"
        ie.Quit: Exit Sub
    End If
    ' 例として、フォーム（パスワード入力）の存在を確認
    Dim doc As Object: Set doc = ie.Document
    Dim inputs As Object: Set inputs = doc.getElementsByTagName("input")
    Dim foundLogin As Boolean: foundLogin = False
    Dim i As Long
    For i = 0 To inputs.Length - 1
        If LCase(inputs(i).Type) = "password" Then
            foundLogin = True: Exit For
        End If
    Next i
    If foundLogin Then
        Debug.Print "AuthenticationPage: OK (Login form detected)"
    Else
        Debug.Print "AuthenticationPage: NG (Login form not detected)"
    End If
    ie.Quit
End Sub

' 8. SSL 証明書エラーの検証
Public Sub Test_EdgeBrowser_SSLCertificateError()
    Debug.Print "Running Test_EdgeBrowser_SSLCertificateError"
    ' 例: https://self-signed.badssl.com/ は SSL 証明書エラーが発生する
    Dim url As String: url = "https://self-signed.badssl.com/"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url, 10)
    If ie Is Nothing Then
        Debug.Print "SSLCertificateError: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie, 10) Then
        Debug.Print "SSLCertificateError: OK (Document load timeout due to SSL error)"
        ie.Quit: Exit Sub
    Else
        Dim bodyText As String: bodyText = ie.Document.body.innerText
        If InStr(bodyText, "certificate") > 0 Then
            Debug.Print "SSLCertificateError: OK (SSL error detected in document)"
        Else
            Debug.Print "SSLCertificateError: NG (SSL error not detected)"
        End If
    End If
    ie.Quit
End Sub

' 9. 特定要素の存在確認とクリック動作
Public Sub Test_EdgeBrowser_ElementExistenceAndClick()
    Debug.Print "Running Test_EdgeBrowser_ElementExistenceAndClick"
    ' 仮想のテストページ（※実際の URL に変更してください）
    Dim url As String: url = "https://www.example.com/testpage_with_button"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "ElementExistenceAndClick: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "ElementExistenceAndClick: NG (Document load timeout)"
        ie.Quit: Exit Sub
    End If
    Dim doc As Object: Set doc = ie.Document
    Dim btn As Object
    On Error Resume Next
    Set btn = doc.getElementById("testButton")
    On Error GoTo 0
    If btn Is Nothing Then
        Debug.Print "ElementExistenceAndClick: NG (Button not found)"
    Else
        btn.Click
        ' ページ遷移などの変化を待機（例: 2秒）
        Application.Wait Now + TimeValue("00:00:02")
        Call AssertEqual("ElementExistenceAndClick Title", "Button Clicked", ie.Document.Title)
    End If
    ie.Quit
End Sub

' 10. JavaScript 実行結果の検証
Public Sub Test_EdgeBrowser_JavaScriptExecutionResult()
    Debug.Print "Running Test_EdgeBrowser_JavaScriptExecutionResult"
    ' 仮想のテストページ（※実際の URL に変更してください）
    Dim url As String: url = "https://www.example.com/testpage_with_js"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "JavaScriptExecutionResult: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "JavaScriptExecutionResult: NG (Document load timeout)"
        ie.Quit: Exit Sub
    End If
    ' 例: ページ内に定義された getTestValue() 関数の返り値を検証（期待値 "Hello"）
    Dim result As Variant
    result = ie.Document.parentWindow.execScript("return getTestValue();", "JScript")
    Call AssertEqual("JavaScriptExecutionResult", "Hello", result)
    ie.Quit
End Sub

' 11. ページリロード動作の検証
Public Sub Test_EdgeBrowser_PageReload()
    Debug.Print "Running Test_EdgeBrowser_PageReload"
    Dim url As String: url = "https://www.example.com"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "PageReload: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "PageReload: NG (Initial document load timeout)"
        ie.Quit: Exit Sub
    End If
    Dim initialTitle As String: initialTitle = ie.Document.Title
    ie.Refresh
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "PageReload: NG (Reload document load timeout)"
        ie.Quit: Exit Sub
    End If
    Call AssertEqual("PageReload Title", initialTitle, ie.Document.Title)
    ie.Quit
End Sub

' 12. 複数ウィンドウ間の切替検証
Public Sub Test_EdgeBrowser_MultipleWindowSwitching()
    Debug.Print "Running Test_EdgeBrowser_MultipleWindowSwitching"
    Dim url1 As String, url2 As String
    url1 = "https://www.example.com"
    url2 = "https://www.iana.org"

    Dim ie1 As SHDocVw.InternetExplorer, ie2 As SHDocVw.InternetExplorer
    Set ie1 = LaunchEdge(url1)
    Set ie2 = LaunchEdge(url2)

    Dim sws As New SHDocVw.ShellWindows, count As Long: count = 0
    Dim tempIE As SHDocVw.InternetExplorer
    For Each tempIE In sws
        If InStr(tempIE.LocationURL, "example.com") > 0 Or InStr(tempIE.LocationURL, "iana.org") > 0 Then
            count = count + 1
        End If
    Next tempIE
    Call AssertEqual("MultipleWindowSwitching Count", 2, count)

    ie1.Quit: ie2.Quit
End Sub

' 13. 既存ブラウザへの再利用（同一ウィンドウでの連続ナビゲーション）
Public Sub Test_EdgeBrowser_BrowserReuse()
    Debug.Print "Running Test_EdgeBrowser_BrowserReuse"
    Dim url1 As String, url2 As String
    url1 = "https://www.example.com"
    url2 = "https://www.iana.org"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url1)
    If ie Is Nothing Then
        Debug.Print "BrowserReuse: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "BrowserReuse: NG (Initial document load timeout)"
        ie.Quit: Exit Sub
    End If
    ie.Navigate url2
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "BrowserReuse: NG (Second document load timeout)"
        ie.Quit: Exit Sub
    End If
    Call AssertEqual("BrowserReuse Title", "IANA — IANA-managed Reserved Domains", ie.Document.Title)
    ie.Quit
End Sub

' 14. ブラウザ終了処理の検証
Public Sub Test_EdgeBrowser_BrowserQuit()
    Debug.Print "Running Test_EdgeBrowser_BrowserQuit"
    Dim url As String: url = "https://www.example.com"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "BrowserQuit: NG (IE window not found)"
        Exit Sub
    End If
    ie.Quit
    Application.Wait Now + TimeValue("00:00:02")

    Dim sws As New SHDocVw.ShellWindows, found As Boolean: found = False
    Dim tempIE As SHDocVw.InternetExplorer
    For Each tempIE In sws
        If InStr(tempIE.LocationURL, "example.com") > 0 Then
            found = True: Exit For
        End If
    Next tempIE
    If found Then
        Debug.Print "BrowserQuit: NG (IE window still exists)"
    Else
        Debug.Print "BrowserQuit: OK (IE window closed)"
    End If
End Sub

' 15. 特殊文字・クエリパラメータ付き URL の検証
Public Sub Test_EdgeBrowser_SpecialCharactersURL()
    Debug.Print "Running Test_EdgeBrowser_SpecialCharactersURL"
    Dim url As String
    url = "https://www.example.com/search?q=" & URLEncode("テスト ケース") & "&lang=ja"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url)
    If ie Is Nothing Then
        Debug.Print "SpecialCharactersURL: NG (IE window not found)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie) Then
        Debug.Print "SpecialCharactersURL: NG (Document load timeout)"
        ie.Quit: Exit Sub
    End If
    ' 簡易検証：タイトルに "Example" が含まれるかチェック
    If InStr(ie.Document.Title, "Example") > 0 Then
        Debug.Print "SpecialCharactersURL: OK"
    Else
        Debug.Print "SpecialCharactersURL: NG (Unexpected title)"
    End If
    ie.Quit
End Sub

' 16. ネットワーク障害時の動作のシミュレーション
Public Sub Test_EdgeBrowser_NetworkFailure()
    Debug.Print "Running Test_EdgeBrowser_NetworkFailure"
    ' 到達不可能な IP を指定
    Dim url As String: url = "http://10.255.255.1"
    Dim ie As SHDocVw.InternetExplorer
    Set ie = LaunchEdge(url, 5)
    If ie Is Nothing Then
        Debug.Print "NetworkFailure: OK (IE window not found as expected)"
        Exit Sub
    End If
    If Not WaitForDocumentComplete(ie, 5) Then
        Debug.Print "NetworkFailure: OK (Document load timeout as expected)"
        ie.Quit: Exit Sub
    Else
        Debug.Print "NetworkFailure: NG (Document unexpectedly loaded)"
    End If
    ie.Quit
End Sub

' --- すべてのテストを連続実行 ---
Public Sub RunAllTests()
    Debug.Print "=== Running All Edge Browser Tests ==="
    Test_EdgeBrowser_NormalNavigation
    Test_EdgeBrowser_MultipleNavigation
    Test_EdgeBrowser_InvalidURL
    Test_EdgeBrowser_EmptyURL
    Test_EdgeBrowser_DocumentLoadTimeout
    Test_EdgeBrowser_RedirectHandling
    Test_EdgeBrowser_AuthenticationPage
    Test_EdgeBrowser_SSLCertificateError
    Test_EdgeBrowser_ElementExistenceAndClick
    Test_EdgeBrowser_JavaScriptExecutionResult
    Test_EdgeBrowser_PageReload
    Test_EdgeBrowser_MultipleWindowSwitching
    Test_EdgeBrowser_BrowserReuse
    Test_EdgeBrowser_BrowserQuit
    Test_EdgeBrowser_SpecialCharactersURL
    Test_EdgeBrowser_NetworkFailure
    Debug.Print "=== All Tests Completed ==="
End Sub
