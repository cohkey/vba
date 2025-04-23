Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'─── 日跨ぎ補正付きラップ計測 ───────────────────────────────────────
Private Function getTimerLap(ByVal startTime As Double) As Double
    Dim lap As Double
    lap = Timer() - startTime
    If lap < 0 Then lap = lap + 86400#
    getTimerLap = lap
End Function

'─── ① 純粋 Timer() ベースのタイムアウトテスト ───────────────────────────────
Private Sub PureTimerTimeout()
    Dim t0 As Double
    t0 = Timer()
    Debug.Print "【PureTimer】 Start at "; Format(Now, "yyyy-mm-dd HH:nn:ss")

    Do While True
        DoEvents
        Sleep 100
        If Timer() - t0 >= 120 Then
            Debug.Print "【PureTimer】 Timeout at "; Format(Now, "yyyy-mm-dd HH:nn:ss")
            Exit Do
        End If
    Loop

    Debug.Print "（このサブは深夜を跨ぐと無限ループになります）"
End Sub

'─── ② getTimerLap ベースのタイムアウトテスト ─────────────────────────
Private Sub LapTimerTimeout()
    Dim t0 As Double
    t0 = Timer()
    Debug.Print "【getTimerLap】 Start at "; Format(Now, "yyyy-mm-dd HH:nn:ss")

    Do While True
        DoEvents
        Sleep 100
        If getTimerLap(t0) >= 120 Then
            Debug.Print "【getTimerLap】 Timeout at "; Format(Now, "yyyy-mm-dd HH:nn:ss")
            Exit Do
        End If
    Loop

    Debug.Print "（深夜を跨いでも正しく2分後に抜けます）"
End Sub

'─── ③ 23:59 監視ループ付きラッパー ─────────────────────────────────
'    マクロ実行直後に現在時刻を監視し、
'    23:59 ちょうどまたは直後になったら指定のテストを呼び出します。

' ■ 純粋 Timer() テスト用ラッパー
Public Sub RunPureAt2359()
    Debug.Print "? RunPureAt2359: 23:59 になるまで待機中..."
    Do
        DoEvents
        Sleep 500    ' 0.5秒ごとにチェック
    Loop While Format(Now, "HH:mm") <> "23:59"


    Debug.Print "→ 23:59 到達: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    PureTimerTimeout
End Sub

' ■ getTimerLap テスト用ラッパー
Public Sub RunLapAt2359()
    Debug.Print "? RunLapAt2359: 23:59 になるまで待機中..."
    Do
        DoEvents
        Sleep 500
    Loop While Format(Now, "HH:mm") <> "23:59"

    Debug.Print "→ 23:59 到達: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    LapTimerTimeout
End Sub

