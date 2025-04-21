Sub ConvertHtmlToPDF()
    Dim fso As Object, wsh As Object
    Dim htmlPath As String, pdfPath As String, command As String

    htmlPath = "C:\path\to\file\sample.html"   '★変換したいHTMLファイル
    pdfPath =  "C:\path\to\file\sample.pdf"    '★出力PDFファイル

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(htmlPath) Then
        MsgBox "HTMLファイルが見つかりません: " & htmlPath, vbExclamation
        Exit Sub
    End If

    'Edgeヘッドレス印刷コマンド組み立て
    command = "msedge.exe --headless --disable-gpu --print-to-pdf=""" & pdfPath & """ --pdf-no-header-footer """ & htmlPath & """"
    '※--pdf-no-header-footer：ヘッダー/フッター（ページ番号や日付）の非表示&#8203;:contentReference[oaicite:7]{index=7}
    '  HTMLパスは file:/// プレフィックスを付けても良い（推奨）&#8203;:contentReference[oaicite:8]{index=8}

    Set wsh = CreateObject("WScript.Shell")
    wsh.Run command, 0, True  'Edgeを非表示(0)・待機モード(True)で実行
    If fso.FileExists(pdfPath) Then
        MsgBox "PDF出力成功: " & pdfPath, vbInformation
    Else
        MsgBox "PDF出力失敗: " & pdfPath, vbCritical
    End If
End Sub
