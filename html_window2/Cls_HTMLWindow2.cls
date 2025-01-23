Option Explicit

Private WithEvents htmlWin As HTMLWindow2


Public Sub Initialize(ByRef html_window2_ As HTMLWindow2)
  Set htmlWin = html_window2_
End Sub

Property Get myHtmlWindow() As HTMLWindow2
  Set myHtmlWindow = htmlWin
End Property


Private Sub htmlWin_onafterprint()
  Debug.Print "onafterprint"
End Sub

Private Sub htmlWin_onbeforeprint()
  Debug.Print "onbeforeprint"
End Sub

Private Sub htmlWin_onbeforeunload()
  Debug.Print "onbeforeunload"
End Sub

Private Sub htmlWin_onblur()
  Debug.Print "onblur"
End Sub

Private Sub htmlWin_onerror(ByVal description As String, ByVal url As String, ByVal line As Long)
  Debug.Print "onerror: " & description, url, line
End Sub

Private Sub htmlWin_onfocus()
  Debug.Print "onfocus"
End Sub

Private Sub htmlWin_onload()
  Debug.Print "onload"
'  Stop
End Sub

Private Sub htmlWin_onresize()
  Debug.Print "onresize"
End Sub

Private Sub htmlWin_onunload()
  Debug.Print "onunload"
End Sub
