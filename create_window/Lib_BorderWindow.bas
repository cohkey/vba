Option Explicit

'================================================================
' Windows API 宣言
'================================================================

' ウィンドウを作成する (拡張スタイル指定可能)
Private Declare PtrSafe Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As LongPtr, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As LongPtr, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As LongPtr, _
    ByVal hMenu As LongPtr, _
    ByVal hInstance As LongPtr, _
    ByVal lpParam As LongPtr) As LongPtr

' ウィンドウを破棄(削除)する
Private Declare PtrSafe Function DestroyWindow Lib "user32.dll" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" ( _
    ByVal hWnd As LongPtr, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long

'SetWindowPos関数
Public Declare PtrSafe Function SetWindowPos Lib "user32" _
    (ByVal hWnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As LongPtr) As Long

' ウィンドウを表示状態にする
Private Declare PtrSafe Function ShowWindow Lib "user32.dll" ( _
    ByVal hWnd As LongPtr, _
    ByVal nCmdShow As Long) As Long

' ウィンドウ全体のDCを取得(非クライアント領域含む)
Private Declare PtrSafe Function GetWindowDC Lib "user32.dll" ( _
    ByVal hWnd As LongPtr) As LongPtr

' DCを解放する
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" ( _
    ByVal hWnd As LongPtr, _
    ByVal hdc As LongPtr) As Long

' 単純なスリープを行う(ミリ秒)
Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" ( _
    ByVal dwMilliseconds As Long)

' 直線や矩形の枠線を描画(指定したペン/ブラシを用いて)
Private Declare PtrSafe Function RectangleAPI Lib "gdi32.dll" Alias "Rectangle" ( _
    ByVal hdc As LongPtr, _
    ByVal nLeftRect As Long, _
    ByVal nTopRect As Long, _
    ByVal nRightRect As Long, _
    ByVal nBottomRect As Long) As Long

' ブラシやペンなどのGDIオブジェクトを削除する
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" ( _
    ByVal hObject As LongPtr) As Long

' ペンを作成する (線のスタイル、太さ、色を指定)
Private Declare PtrSafe Function CreatePen Lib "gdi32.dll" ( _
    ByVal fnPenStyle As Long, _
    ByVal nWidth As Long, _
    ByVal crColor As Long) As LongPtr

' DCに選択するオブジェクト(ブラシやペンなど)を切り替える
Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" ( _
    ByVal hdc As LongPtr, _
    ByVal hObj As LongPtr) As LongPtr

' ストックオブジェクト(既定オブジェクト)を取得する
Private Declare PtrSafe Function GetStockObject Lib "gdi32.dll" ( _
    ByVal nObject As Long) As LongPtr


' ペンを作成(間接指定)
Private Declare PtrSafe Function CreatePenIndirect Lib "gdi32.dll" ( _
    ByRef lpLogPen As LOGPEN) As LongPtr

'ウィンドウの座標を取得する関数
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

'POINTAPI構造体
Public Type POINTAPI
    x As Long
    y As Long
End Type

'LOGPEN構造体
Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

'================================================================
' 定数定義
'================================================================
Private Const WS_EX_LAYERED As LongPtr = &H80000
Private Const WS_POPUP As LongPtr = &H80000000
Private Const WS_VISIBLE As LongPtr = &H10000000
Private Const SW_SHOW As Long = 5


Public Const HWND_TOP = 0           '手前にセット
Public Const HWND_BOTTOM = 1        '後ろにセット
Public Const HWND_TOPMOST = (-1&)      '常に手前にセット
Public Const HWND_NOTOPMOST = (-2&)    '常に手間を解除


'wFlagsに使う定数
Public Const SWP_NOSIZE = &H1       'ウィンドウのサイズを変更しない
Public Const SWP_NOMOVE = &H2       'ウィンドウの位置を変更しない
Public Const SWP_NOZORDER = &H4 'Zオーダーを変更しない
Public Const SWP_FRAMECHANGED = &H20 'ウィンドウを再描画する
Public Const SWP_SHOWWINDOW = &H40  'ウィンドウを表示する
Public Const SWP_NOACTIVATE = &H10 'ウィンドウをアクティブにしない

'dwFlagsに使う定数
Public Const LWA_COLORKEY = &H1& '透過する色を、第二引数(crey)で指定する
Public Const LWA_ALPHA = &H2& 'ウィンドウ全体の透過率を、第三引数(bAlpha)で指定する

' ペンスタイル
Private Const PS_SOLID As Long = 0

' ストックオブジェクトの定義
Private Const NULL_BRUSH As Long = 5  ' 中身を塗りつぶさないブラシ
Private Const NULL_PEN As Long = 8      ' 透明ペン




Public g_hBorderWindow As LongPtr





'================================================================
' ウィンドウを作成して、枠線のみを赤色に描画し10秒後破棄するサンプル
'================================================================
Public Function CreateBorderWindow(ByVal h_parent As LongPtr, _
                                                    ByVal x_pos_ As Long, _
                                                    ByVal y_pos_ As Long, _
                                                    ByVal width_ As Long, _
                                                    ByVal height_ As Long, _
                                                    Optional ByVal rgb_color_ As Long = 255) As LongPtr


    Dim hWnd As LongPtr
    hWnd = CreateLayeredWindow(h_parent, x_pos_, y_pos_, width_, height_, rgb_color_)
    If hWnd = 0 Then
        Stop
        Exit Function
    End If

    g_hBorderWindow = hWnd

'    DoEvents

    ' 背景色を透明にする
    If ClearWindowBackground(hWnd) = False Then
        Stop
        Exit Function
    End If

    ' 枠線を描画
    If DrawWindowBorder(hWnd) = False Then
        Stop
        Exit Function
    End If

    CreateBorderWindow = hWnd

'    Call Sleep(10000)
'
'    ' ウィンドウ破棄
'    Call DestroyWindow(hWnd)

End Function


Private Function CreateLayeredWindow(ByVal h_parent As LongPtr, _
                                                                ByVal x_pos_ As Long, _
                                                                ByVal y_pos_ As Long, _
                                                                ByVal width_ As Long, _
                                                                ByVal height_ As Long, _
                                                Optional ByVal rgb_color_ As Long = 255) As LongPtr

    '-----------------------------
    ' ローカル変数宣言 (キャメルケース)
    '-----------------------------
    Dim hWnd As LongPtr

    Dim xPos As Long
    Dim yPos As Long
    Dim widthWindow As Long
    Dim heightWindow As Long

    '-----------------------------
    ' ウィンドウ作成 (Layered + Popup + Visible)
    ' Staticクラスは簡易サンプル用
    '-----------------------------
    hWnd = CreateWindowEx( _
                    WS_EX_LAYERED, _
                    "Static", _
                    "Custom Border Window", _
                    WS_POPUP Or WS_VISIBLE, _
                    x_pos_, y_pos_, _
                    width_, height_, _
                    h_parent, 0, 0, 0)

    'ウィンドウを透明化
    SetLayeredWindowAttributes hWnd, vbWhite, 0, LWA_COLORKEY
'    SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA

    'ウィンドウを最前面化
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    DoEvents

    CreateLayeredWindow = hWnd
End Function


Private Function ClearWindowBackground(Optional ByVal hwnd_ As LongPtr) As Boolean
    If hwnd_ = 0 Then
        Exit Function
    End If

    '対象ウィンドウの座標を取得
    Dim r As RECT
    GetWindowRect hwnd_, r

    '対象ウィンドウのデバイスコンテキストを取得
    Dim hdc As LongPtr
    hdc = GetWindowDC(hwnd_)

    '-----------------------
    '透明ペンの準備

    '透明なブラシ（システムで定義済み）のポインタを取得
    Dim hNullPen As LongPtr
    hNullPen = GetStockObject(NULL_PEN)

    '対象のデバイスコンテキストで使用するペンとして透明ペンを指定
    Dim hOldPen As LongPtr
    hOldPen = SelectObject(hdc, hNullPen)
    '※hOldPenには、これまでhdcで使っていたペンのポインタが入る

    '==================================================================

    '透明ペンと透明ブラシを使って、中が透明な四角形を描画
    RectangleAPI hdc, 0, 0, r.right - r.left + 1, r.bottom - r.top + 1

    '対象のデバイスコンテキストで使用するペンとブラシを元に戻す
    SelectObject hdc, hOldPen

    'デバイスコンテキストを解放する
    ReleaseDC hwnd_, hdc

    ClearWindowBackground = True
End Function



Private Function DrawWindowBorder(ByVal hwnd_ As LongPtr, Optional ByVal color_number As Long = 255) As Boolean
    '対象ウィンドウの座標を取得
    Dim r As RECT

    If hwnd_ = 0 Then
        Exit Function
    End If

    GetWindowRect hwnd_, r

    '対象ウィンドウのデバイスコンテキストを取得
    Dim hdc As LongPtr
    hdc = GetWindowDC(hwnd_)

    '==================================================================
    '赤ペンの準備

    '太さ３のペンを作成（色は引数で指定）
    Dim redPen As LOGPEN
    redPen.lopnColor = color_number
    redPen.lopnWidth.x = 3


    '作成したペンのポインタを取得
    Dim hRedPen As LongPtr
    hRedPen = CreatePenIndirect(redPen)

    '対象のデバイスコンテキストで使用するペンとして赤ペンを指定
    Dim hOldPen As LongPtr
    hOldPen = SelectObject(hdc, hRedPen)
    '※hOldPenには、これまでhdcで使っていたペンのポインタが入る
    '==================================================================

    '==================================================================
    '透明ブラシの準備

    '透明なブラシ（システムで定義済み）のポインタを取得
    Dim hNullBrush As LongPtr
    hNullBrush = GetStockObject(NULL_BRUSH)

    '対象のデバイスコンテキストで使用するブラシとして透明なブラシを準備
    Dim hOldBrush As LongPtr
    hOldBrush = SelectObject(hdc, hNullBrush)
    '※hOldBrushには、これまでhdcで使っていたブラシのポインタが入る
    '==================================================================

    DoEvents

    '赤ペンと透明ブラシを使って、中が透明な四角形を描画
    RectangleAPI hdc, 1, 1, r.right - r.left - 1, r.bottom - r.top - 1


    '対象のデバイスコンテキストで使用するペンとブラシを元に戻す
    SelectObject hdc, hOldPen
    SelectObject hdc, hOldBrush

    '準備した赤ペンのメモリ領域を解放する（透明ブラシはシステムのデフォルトなので解放不要）
    DeleteObject hRedPen

    'デバイスコンテキストを解放する
    ReleaseDC hwnd_, hdc

    DrawWindowBorder = True
End Function


'================================================================
' RGB値をLong型に変換 (WindowsのRGBマクロ相当)
'   / 引数: r, g, b(0～255)
'   / 戻り値: (b * &H10000) + (g * &H100) + r
'================================================================
Public Function RgbColor(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    RgbColor = (b * &H10000) + (g * &H100) + r
End Function
