Option Explicit


#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub EnumerateFileTypeListItems()
    Dim uiAuto As CUIAutomation
    Set uiAuto = New CUIAutomation

    Dim root As IUIAutomationElement
    Set root = uiAuto.GetRootElement()

    ' 1. 「名前を付けて保存」ダイアログを取得
    Dim condSaveAs As IUIAutomationCondition
    Set condSaveAs = uiAuto.CreatePropertyCondition(UIA_NamePropertyId, "名前を付けて保存")

    Dim saveAsDialog As IUIAutomationElement
    Set saveAsDialog = root.FindFirst(TreeScope_Subtree, condSaveAs)
    If saveAsDialog Is Nothing Then
        MsgBox "名前を付けて保存ダイアログが見つかりません。"
        Exit Sub
    End If

    ' 2. ファイルの種類コンボボックスの取得
    ' まずは AutomationId ("FileTypeControlHost") を利用
    Dim condFileTypeCombo As IUIAutomationCondition
    Set condFileTypeCombo = uiAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "FileTypeControlHost")

    Dim comboBox As IUIAutomationElement
    Set comboBox = saveAsDialog.FindFirst(TreeScope_Subtree, condFileTypeCombo)

    ' AutomationId で取得できなければ、NameとControlTypeのAND条件で取得
    If comboBox Is Nothing Then
        Dim condName As IUIAutomationCondition
        Set condName = uiAuto.CreatePropertyCondition(UIA_NamePropertyId, "ファイルの種類:")
        Dim condType As IUIAutomationCondition
        Set condType = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ComboBoxControlTypeId)

        Dim condAnd As IUIAutomationCondition
        Set condAnd = uiAuto.CreateAndCondition(condName, condType)

        Set comboBox = saveAsDialog.FindFirst(TreeScope_Subtree, condAnd)
    End If

    If comboBox Is Nothing Then
        MsgBox "ファイルの種類コンボボックスが見つかりません。"
        Exit Sub
    End If

    ' 3. コンボボックスを展開
    Dim expandPattern As IUIAutomationExpandCollapsePattern
    On Error Resume Next
    Set expandPattern = comboBox.GetCurrentPattern(UIA_ExpandCollapsePatternId)
    On Error GoTo 0

    If expandPattern Is Nothing Then
        MsgBox "ExpandCollapsePattern が取得できません。"
        Exit Sub
    End If

    expandPattern.Expand
    Sleep 500  ' UI更新待ち

    ' 4. 展開後のリスト要素を取得
    Dim condList As IUIAutomationCondition
    Set condList = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ListControlTypeId)

    Dim fileTypeList As IUIAutomationElement
    Set fileTypeList = comboBox.FindFirst(TreeScope_Children, condList)
    If fileTypeList Is Nothing Then
        ' コンボボックス直下に見つからない場合、ダイアログ全体から探す
        Set fileTypeList = saveAsDialog.FindFirst(TreeScope_Subtree, condList)
    End If

    If fileTypeList Is Nothing Then
        MsgBox "ファイルの種類リストが見つかりません。"
        Exit Sub
    End If

    ' 5. リスト内のすべての ListItem (選択項目) を取得
    Dim condListItem As IUIAutomationCondition
    Set condListItem = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ListItemControlTypeId)

    Dim listItems As IUIAutomationElementArray
    Set listItems = fileTypeList.FindAll(TreeScope_Children, condListItem)

    Dim i As Long
    Dim msg As String
    msg = "ファイルの種類一覧:" & vbCrLf
    Dim listItem As IUIAutomationElement
    For i = 0 To listItems.length - 1
        Set listItem = listItems.GetElement(i)
        msg = msg & listItem.CurrentName & vbCrLf
    Next i

    ' 6. コンボボックスを閉じる (Collapse)
    expandPattern.Collapse
    Sleep 500  ' UI更新待ち

    MsgBox msg
End Sub

