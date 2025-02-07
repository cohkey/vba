Option Explicit

#If Win64 Then
    ' ウィンドウ操作用 API (64bit)
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

    ' MSAA Accessible API (64bit)
    Private Declare PtrSafe Function AccessibleChildren Lib "oleacc.dll" ( _
        ByVal paccContainer As IAccessible, _
        ByVal iChildStart As Long, _
        ByVal cChildren As Long, _
        ByRef rgvarChildren As Variant, _
        ByRef pcObtained As Long) As Long

    Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc.dll" ( _
        ByVal hWnd As LongPtr, _
        ByVal dwObjectID As Long, _
        ByRef riid As UUID, _
        ByRef ppvObject As IAccessible) As Long
#Else
    ' 32bit宣言（必要な場合）
#End If

' 定数
Private Const OBJID_CLIENT As Long = &HFFFFFFFC
Private Const CHILDID_SELF As Long = 0&

' IAccessibleのIIDを保持するUUID構造体
Private Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private IID_IAccessible As UUID

' 初期化ルーチン: IID_IAccessible の値をセットする
Sub InitializeIAccessibleIID()
    With IID_IAccessible
        .Data1 = &H618736E0
        .Data2 = &H3C3D
        .Data3 = &H11CF
        .Data4(0) = &H81
        .Data4(1) = &HC
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H38
        .Data4(6) = &H9B
        .Data4(7) = &H71
    End With
End Sub

' ----------------------------------------------------------
' 再帰的に IAccessible オブジェクトの子要素を探索し、
' 指定した accName を持つオブジェクトを返す関数
' ----------------------------------------------------------
Function FindAccessibleChildByName(pAcc As IAccessible, sName As String) As IAccessible
    Dim childCount As Long, retCount As Long
    Dim varChildren As Variant
    Dim i As Long
    On Error Resume Next
    childCount = pAcc.accChildCount
    If childCount > 0 Then
        ReDim varChildren(childCount - 1)
        AccessibleChildren pAcc, 0, childCount, varChildren(0), retCount
        For i = LBound(varChildren) To UBound(varChildren)
            Dim pChild As IAccessible
            ' AccessibleChildrenで得られた子要素は IAccessible 型の場合と、数値 (子ID) の場合があるのでチェック
            If Not IsEmpty(varChildren(i)) Then
                If TypeName(varChildren(i)) = "IAccessible" Then
                    Set pChild = varChildren(i)
                    If pChild.accName(CHILDID_SELF) = sName Then
                        Set FindAccessibleChildByName = pChild
                        Exit Function
                    End If
                    ' 再帰的に子要素を検索
                    Set pChild = FindAccessibleChildByName(pChild, sName)
                    If Not pChild Is Nothing Then
                        Set FindAccessibleChildByName = pChild
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------
' Edgeウィンドウ（ClassName="Chrome_WidgetWin_1" の最初のもの）から
' "Internet Explorer モードのリロード タブ" という accName を持つボタンを
' 見つけ出し、accDoDefaultAction でクリック動作を実行する
' ----------------------------------------------------------
Sub ClickIEModeButton()
    Dim hWndEdge As LongPtr
    Dim accObj As IAccessible
    Dim hr As Long
    Dim ieModeButton As IAccessible

    ' IAccessibleのUUID初期化
    InitializeIAccessibleIID

    ' Edgeウィンドウの取得 (最初に見つかった "Chrome_WidgetWin_1")
    hWndEdge = FindWindowEx(0, 0, "Chrome_WidgetWin_1", vbNullString)
    If hWndEdge = 0 Then
        MsgBox "Edgeウィンドウが見つかりませんでした。"
        Exit Sub
    End If

    ' 指定したウィンドウから IAccessible オブジェクトを取得 (OBJID_CLIENT)
    hr = AccessibleObjectFromWindow(hWndEdge, OBJID_CLIENT, IID_IAccessible, accObj)
    If hr <> 0 Then
        MsgBox "AccessibleObjectFromWindow の呼び出しに失敗しました。hr=" & hr
        Exit Sub
    End If

    ' 再帰的に IEモードボタン (accName = "Internet Explorer モードのリロード タブ") を探索
    Set ieModeButton = FindAccessibleChildByName(accObj, "Internet Explorer モードのリロード タブ")
    If ieModeButton Is Nothing Then
        MsgBox "IEモードボタンが見つかりませんでした。"
        Exit Sub
    End If

    ' 操作対象ウィンドウを前面に出す
    SetForegroundWindow hWndEdge

    ' IEモードボタンの既定動作 (accDoDefaultAction) を実行
    On Error Resume Next
    ieModeButton.accDoDefaultAction CHILDID_SELF
    If Err.Number <> 0 Then
        MsgBox "accDoDefaultAction の実行に失敗しました: " & Err.Description
    Else
        MsgBox "IEモードボタンが実行されました。"
    End If
    On Error GoTo 0
End Sub
