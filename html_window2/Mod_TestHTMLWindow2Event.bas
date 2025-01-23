Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)


Private Sub TestHtmlWindow2Evnets()

  Dim htmlDoc As HTMLDocument
  Set htmlDoc = GetWindow("Example Domain")    'https://example.com/

  Dim clsHtmlWin As Cls_Window2
  Set clsHtmlWin = New Cls_Window2
  Call clsHtmlWin.Initialize(htmlDoc.parentWindow)

  Dim url As String
  url = "https://example.com/"

  ' 指定のURLにナビゲート
  clsHtmlWin.myHtmlWindow.navigate url

  ' ページの読み込み完了を待つ
  Do While htmlDoc.readyState <> "complete"
      DoEvents
      Sleep 20
      Debug.Print htmlDoc.readyState
  Loop
  Debug.Print "complete"
End Sub