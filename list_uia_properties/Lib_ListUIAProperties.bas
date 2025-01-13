Option Explicit

'***********************************************************************
' グローバル変数
'***********************************************************************
Private g_uia As CUIAutomation8    ' UIAutomation全体で使う
Private m_rowIndex As Long         ' Excel出力行を管理

'***********************************************************************
' 概要  ：Microsoft Edge(Chromium) ウィンドウ(ClassName="Chrome_WidgetWin_1") を最初に見つけ、
'         その要素をルートとして配下の全要素を再帰的に走査し、
'         各要素のプロパティをExcelシートへ出力する
'
' 引数  ：なし
' 戻り値：なし
'***********************************************************************
Public Sub CollectEdgeUIProperties()

    Debug.Print "==== CollectEdgeUIProperties 開始 ===="

    On Error GoTo ERR_HANDLER

    '-------------------------------------------------------------------
    ' 1. UIAutomation初期化
    '-------------------------------------------------------------------
    Set g_uia = New CUIAutomation8
    If g_uia Is Nothing Then
        MsgBox "UI Automationの初期化に失敗しました。", vbExclamation
        Debug.Print "UI Automation初期化失敗: g_uia Is Nothing"
        Exit Sub
    Else
        Debug.Print "UI Automation初期化成功"
    End If

    '-------------------------------------------------------------------
    ' 2. ルート要素取得
    '-------------------------------------------------------------------
    Dim rootElement As IUIAutomationElement
    Set rootElement = g_uia.GetRootElement
    If rootElement Is Nothing Then
        MsgBox "ルート要素が取得できませんでした。", vbExclamation
        Debug.Print "GetRootElement 失敗: rootElement Is Nothing"
        Exit Sub
    End If
    Debug.Print "ルート要素取得成功"

    '-------------------------------------------------------------------
    ' 3. 検索条件作成 (ClassName="Chrome_WidgetWin_1")
    '-------------------------------------------------------------------
    Const UIA_ClassNamePropertyId As Long = 30012

    Dim conditionClass As IUIAutomationCondition
    Set conditionClass = g_uia.CreatePropertyCondition( _
        UIA_ClassNamePropertyId, "Chrome_WidgetWin_1")

    Debug.Print "検索条件: Chrome_WidgetWin_1"

    '-------------------------------------------------------------------
    ' 4. 最初のEdgeウィンドウ要素を取得 (FindFirst - Subtree)
    '-------------------------------------------------------------------
    Dim edgeElement As IUIAutomationElement
    Set edgeElement = rootElement.FindFirst(TreeScope_Subtree, conditionClass)

    If edgeElement Is Nothing Then
        MsgBox "Chrome_WidgetWin_1 の最初の要素が見つかりません。", vbExclamation
        Debug.Print "Edge要素見つからず"
        Exit Sub
    End If

    Debug.Print "Edgeウィンドウ要素取得成功"

    '-------------------------------------------------------------------
    ' 5. Excel出力の準備（ヘッダ行）
    '-------------------------------------------------------------------
    Dim ws_active As Worksheet
    Set ws_active = ActiveSheet

    ws_active.Cells(1, 1).Value = "Level"
    ws_active.Cells(1, 2).Value = "ParentRuntimeId"
    ws_active.Cells(1, 3).Value = "RuntimeId"
    ws_active.Cells(1, 4).Value = "Name"
    ws_active.Cells(1, 5).Value = "AutomationId"
    ws_active.Cells(1, 6).Value = "ControlTypeID"
    ws_active.Cells(1, 7).Value = "ControlTypeLabel"
    ws_active.Cells(1, 8).Value = "ClassName"
    ws_active.Cells(1, 9).Value = "BoundingRectangle"
    ws_active.Cells(1, 10).Value = "FrameworkId"
    ws_active.Cells(1, 11).Value = "IsEnabled"
    ws_active.Cells(1, 12).Value = "IsOffscreen"
    ws_active.Cells(1, 13).Value = "ProviderDescription"
    ws_active.Cells(1, 14).Value = "HasKeyboardFocus"
    ws_active.Cells(1, 15).Value = "ProcessId"
    ws_active.Cells(1, 16).Value = "AcceleratorKey"
    ws_active.Cells(1, 17).Value = "AccessKey"
    ws_active.Cells(1, 18).Value = "AriaRole"

    m_rowIndex = 2

    '-------------------------------------------------------------------
    ' 6. 再帰的に全要素を走査して出力
    '-------------------------------------------------------------------
    Debug.Print "===== 再帰スキャン開始 ====="
    Call ScanElements(edgeElement, 1, "N/A")  ' レベル1、親RuntimeId="N/A"で開始
    Debug.Print "===== 再帰スキャン終了 ====="

    MsgBox "Edgeウィンドウ配下の全要素のプロパティを取得しました。", vbInformation
    Debug.Print "==== CollectEdgeUIProperties 完了 ===="
    Exit Sub

'-----------------------------------------------------------------------
' エラーハンドリング
'-----------------------------------------------------------------------
ERR_HANDLER:
    MsgBox "エラーが発生しました: " & Err.Number & " - " & Err.Description, vbCritical
    Debug.Print "エラー発生: " & Err.Number & " - " & Err.Description
End Sub

'***********************************************************************
' 概要  ：要素 targetElem を起点に配下の全子要素を再帰的に走査し、
'         必要なプロパティをExcelに出力する
'
' 引数  ：ByVal targetElem       As IUIAutomationElement - 走査の起点となる要素
'         ByVal level            As Long                 - 階層レベル
'         ByVal parentRuntimeId  As String              - 親要素のRuntimeId
'
' 戻り値：なし
'***********************************************************************
Private Sub ScanElements(ByVal targetElem As IUIAutomationElement, _
                         ByVal level As Long, _
                         ByVal parentRuntimeId As String)

    On Error GoTo ERR_Scan

    '-------------------------------------------------------------------
    ' プロパティID定義
    '-------------------------------------------------------------------
    Const UIA_RuntimeIdPropertyId As Long = 30000
    Const UIA_BoundingRectanglePropertyId As Long = 30001
    Const UIA_ProcessIdPropertyId As Long = 30002
    Const UIA_ControlTypePropertyId As Long = 30003
    Const UIA_NamePropertyId As Long = 30005
    Const UIA_AcceleratorKeyPropertyId As Long = 30006
    Const UIA_AccessKeyPropertyId As Long = 30007
    Const UIA_AutomationIdPropertyId As Long = 30011
    Const UIA_ClassNamePropertyId As Long = 30012
    Const UIA_IsEnabledPropertyId As Long = 30010
    Const UIA_IsOffscreenPropertyId As Long = 30022
    Const UIA_FrameworkIdPropertyId As Long = 30024
    Const UIA_HasKeyboardFocusPropertyId As Long = 30036
    Const UIA_ProviderDescriptionPropertyId As Long = 30107
    Const UIA_AriaRolePropertyId As Long = 30101

    '-------------------------------------------------------------------
    ' 1. 各種プロパティ取得
    '-------------------------------------------------------------------
    Dim runtimeIdVal As Variant: runtimeIdVal = targetElem.GetCurrentPropertyValue(UIA_RuntimeIdPropertyId)
    Dim boundingRectVal As Variant: boundingRectVal = targetElem.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
    Dim processIdVal As Variant: processIdVal = targetElem.GetCurrentPropertyValue(UIA_ProcessIdPropertyId)
    Dim controlTypeVal As Variant: controlTypeVal = targetElem.GetCurrentPropertyValue(UIA_ControlTypePropertyId)
    Dim nameVal As Variant: nameVal = targetElem.GetCurrentPropertyValue(UIA_NamePropertyId)
    Dim acceleratorKeyVal As Variant: acceleratorKeyVal = targetElem.GetCurrentPropertyValue(UIA_AcceleratorKeyPropertyId)
    Dim accessKeyVal As Variant: accessKeyVal = targetElem.GetCurrentPropertyValue(UIA_AccessKeyPropertyId)
    Dim automationIdVal As Variant: automationIdVal = targetElem.GetCurrentPropertyValue(UIA_AutomationIdPropertyId)
    Dim classNameVal As Variant: classNameVal = targetElem.GetCurrentPropertyValue(UIA_ClassNamePropertyId)
    Dim isEnabledVal As Variant: isEnabledVal = targetElem.GetCurrentPropertyValue(UIA_IsEnabledPropertyId)
    Dim isOffscreenVal As Variant: isOffscreenVal = targetElem.GetCurrentPropertyValue(UIA_IsOffscreenPropertyId)
    Dim frameworkIdVal As Variant: frameworkIdVal = targetElem.GetCurrentPropertyValue(UIA_FrameworkIdPropertyId)
    Dim hasKeyboardFocusVal As Variant: hasKeyboardFocusVal = targetElem.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId)
    Dim providerDescVal As Variant: providerDescVal = targetElem.GetCurrentPropertyValue(UIA_ProviderDescriptionPropertyId)
    Dim ariaRoleVal As Variant: ariaRoleVal = targetElem.GetCurrentPropertyValue(UIA_AriaRolePropertyId)

    '-------------------------------------------------------------------
    ' 2. Excelへ出力
    '-------------------------------------------------------------------
    Dim ws_active As Worksheet
    Set ws_active = ActiveSheet

    Dim runIdStr As String
    If IsArray(runtimeIdVal) Then
        runIdStr = ArrayToString(runtimeIdVal, "-")
    Else
        runIdStr = CStr(runtimeIdVal)
    End If

    Dim boundingRectStr As String
    If IsArray(boundingRectVal) Then
        boundingRectStr = ArrayToString(boundingRectVal, ",")
    Else
        boundingRectStr = CStr(boundingRectVal)
    End If

    ' ControlTypeID → 文字列ラベル
    Dim ctlTypeId As Long
    If Not IsError(controlTypeVal) And Not IsNull(controlTypeVal) And controlTypeVal <> "" Then
        ctlTypeId = CLng(controlTypeVal)
    Else
        ctlTypeId = -1
    End If

    Dim ctlTypeLabel As String
    ctlTypeLabel = ControlTypeIdToLabel(ctlTypeId)

    ' 出力
    ws_active.Cells(m_rowIndex, 1).Value = level
    ws_active.Cells(m_rowIndex, 2).Value = parentRuntimeId
    ws_active.Cells(m_rowIndex, 3).Value = runIdStr
    ws_active.Cells(m_rowIndex, 4).Value = CStr(nameVal)
    ws_active.Cells(m_rowIndex, 5).Value = CStr(automationIdVal)
    ws_active.Cells(m_rowIndex, 6).Value = ctlTypeId
    ws_active.Cells(m_rowIndex, 7).Value = ctlTypeLabel
    ws_active.Cells(m_rowIndex, 8).Value = CStr(classNameVal)
    ws_active.Cells(m_rowIndex, 9).Value = boundingRectStr
    ws_active.Cells(m_rowIndex, 10).Value = CStr(frameworkIdVal)
    ws_active.Cells(m_rowIndex, 11).Value = CStr(isEnabledVal)
    ws_active.Cells(m_rowIndex, 12).Value = CStr(isOffscreenVal)
    ws_active.Cells(m_rowIndex, 13).Value = CStr(providerDescVal)
    ws_active.Cells(m_rowIndex, 14).Value = CStr(hasKeyboardFocusVal)
    ws_active.Cells(m_rowIndex, 15).Value = CStr(processIdVal)
    ws_active.Cells(m_rowIndex, 16).Value = CStr(acceleratorKeyVal)
    ws_active.Cells(m_rowIndex, 17).Value = CStr(accessKeyVal)
    ws_active.Cells(m_rowIndex, 18).Value = CStr(ariaRoleVal)

    Debug.Print "Level=" & level & " RuntimeId=" & runIdStr & _
                " Name=" & CStr(nameVal) & " ClassName=" & CStr(classNameVal)

    m_rowIndex = m_rowIndex + 1

    '-------------------------------------------------------------------
    ' 3. 子要素を取得 → 再帰的に処理
    '-------------------------------------------------------------------
    Dim children As IUIAutomationElementArray
    Set children = targetElem.FindAll(TreeScope_Children, g_uia.CreateTrueCondition)

    If Not children Is Nothing Then
        Dim childCount As Long
        childCount = children.Length

        Debug.Print "└ 子要素数=" & childCount & " (Level=" & level & ", RuntimeId=" & runIdStr & ")"

        Dim i As Long
        Dim childElem As IUIAutomationElement

        For i = 0 To childCount - 1
            Set childElem = children.GetElement(i)
            If Not childElem Is Nothing Then
                Call ScanElements(childElem, level + 1, runIdStr)
            End If
        Next i
    Else
        Debug.Print "└ 子要素取得がNothing (Level=" & level & ", RuntimeId=" & runIdStr & ")"
    End If

    Exit Sub

'-----------------------------------------------------------------------
' エラーハンドリング (ScanElements)
'-----------------------------------------------------------------------
ERR_Scan:
    Debug.Print "ScanElementsエラー: " & Err.Number & " - " & Err.Description
End Sub

'***********************************************************************
' 概要  ：ControlTypeID → 文字列ラベル変換
'***********************************************************************
Private Function ControlTypeIdToLabel(ByVal ctlTypeId As Long) As String
    Select Case ctlTypeId
        Case 50000: ControlTypeIdToLabel = "Button"
        Case 50001: ControlTypeIdToLabel = "Calendar"
        Case 50002: ControlTypeIdToLabel = "CheckBox"
        Case 50003: ControlTypeIdToLabel = "ComboBox"
        Case 50004: ControlTypeIdToLabel = "Edit"
        Case 50005: ControlTypeIdToLabel = "Hyperlink"
        Case 50006: ControlTypeIdToLabel = "Image"
        Case 50007: ControlTypeIdToLabel = "ListItem"
        Case 50008: ControlTypeIdToLabel = "List"
        Case 50009: ControlTypeIdToLabel = "Menu"
        Case 50010: ControlTypeIdToLabel = "MenuBar"
        Case 50011: ControlTypeIdToLabel = "MenuItem"
        Case 50012: ControlTypeIdToLabel = "ProgressBar"
        Case 50013: ControlTypeIdToLabel = "RadioButton"
        Case 50014: ControlTypeIdToLabel = "ScrollBar"
        Case 50015: ControlTypeIdToLabel = "Slider"
        Case 50016: ControlTypeIdToLabel = "Spinner"
        Case 50017: ControlTypeIdToLabel = "StatusBar"
        Case 50018: ControlTypeIdToLabel = "Tab"
        Case 50019: ControlTypeIdToLabel = "TabItem"
        Case 50020: ControlTypeIdToLabel = "Text"
        Case 50021: ControlTypeIdToLabel = "ToolBar"
        Case 50022: ControlTypeIdToLabel = "ToolTip"
        Case 50023: ControlTypeIdToLabel = "Tree"
        Case 50024: ControlTypeIdToLabel = "TreeItem"
        Case 50025: ControlTypeIdToLabel = "Custom"
        Case 50026: ControlTypeIdToLabel = "Group"
        Case 50027: ControlTypeIdToLabel = "Thumb"
        Case 50028: ControlTypeIdToLabel = "DataGrid"
        Case 50029: ControlTypeIdToLabel = "DataItem"
        Case 50030: ControlTypeIdToLabel = "Document"
        Case 50031: ControlTypeIdToLabel = "SplitButton"
        Case 50032: ControlTypeIdToLabel = "Window"
        Case 50033: ControlTypeIdToLabel = "Pane"
        Case 50034: ControlTypeIdToLabel = "Header"
        Case 50035: ControlTypeIdToLabel = "HeaderItem"
        Case 50036: ControlTypeIdToLabel = "Table"
        Case 50037: ControlTypeIdToLabel = "TitleBar"
        Case 50038: ControlTypeIdToLabel = "Separator"
        Case Else
            ControlTypeIdToLabel = "Unknown(" & ctlTypeId & ")"
    End Select
End Function

'***********************************************************************
' 概要  ：Variant配列を文字列に結合して返す (非配列ならそのまま文字列化)
'***********************************************************************
Private Function ArrayToString(ByVal arr_values As Variant, _
                               ByVal delimiter As String) As String
    If Not IsArray(arr_values) Then
        ArrayToString = CStr(arr_values)
        Exit Function
    End If

    Dim resultStr As String
    Dim i As Long
    resultStr = ""

    For i = LBound(arr_values) To UBound(arr_values)
        resultStr = resultStr & CStr(arr_values(i))
        If i < UBound(arr_values) Then
            resultStr = resultStr & delimiter
        End If
    Next i

    ArrayToString = resultStr
End Function


