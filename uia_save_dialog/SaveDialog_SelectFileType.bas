Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub ChangeFileTypeInSaveAsDialog()
    Dim uiAuto As CUIAutomation
    Set uiAuto = New CUIAutomation

    Dim root As IUIAutomationElement
    Set root = uiAuto.GetRootElement()

    ' 1. 「名前を付けて保存」ダイアログを探す
    Dim condSaveAs As IUIAutomationCondition
    Set condSaveAs = uiAuto.CreatePropertyCondition(UIA_NamePropertyId, "名前を付けて保存")

    Dim saveAsDialog As IUIAutomationElement
    Set saveAsDialog = root.FindFirst(TreeScope_Subtree, condSaveAs)
    If saveAsDialog Is Nothing Then
        MsgBox "名前を付けて保存ダイアログが見つかりませんでした。"
        Exit Sub
    End If

    ' 2. ファイルの種類コンボボックスを取得
    ' まずは AutomationId を利用 (Inspect.exe で確認: "FileTypeControlHost")
    Dim condFileTypeCombo As IUIAutomationCondition
    Set condFileTypeCombo = uiAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "FileTypeControlHost")

    Dim comboBox As IUIAutomationElement
    Set comboBox = saveAsDialog.FindFirst(TreeScope_Subtree, condFileTypeCombo)

    ' AutomationId で取得できなければ、NameとControlTypeをAND条件で指定
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

    ' 3. ExpandCollapsePattern を取得してコンボボックスを展開
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
        Set fileTypeList = saveAsDialog.FindFirst(TreeScope_Subtree, condList)
    End If

    If fileTypeList Is Nothing Then
        MsgBox "ファイルの種類リストが見つかりませんでした。"
        Exit Sub
    End If

    ' 5. リスト内から目的の項目を取得 (例: "Web ページ、単一ファイル (*.mhtml)")
    Dim condItem As IUIAutomationCondition
    Set condItem = uiAuto.CreatePropertyCondition(UIA_NamePropertyId, "Web ページ、単一ファイル (*.mhtml)")

    Dim targetItem As IUIAutomationElement
    Set targetItem = fileTypeList.FindFirst(TreeScope_Children, condItem)

    If targetItem Is Nothing Then
        MsgBox "指定のファイル種類が見つかりません。"
        Exit Sub
    End If

    ' 6. SelectionItemPattern を利用して項目を選択
    Dim selPattern As IUIAutomationSelectionItemPattern
    Set selPattern = targetItem.GetCurrentPattern(UIA_SelectionItemPatternId)

    If selPattern Is Nothing Then
        MsgBox "SelectionItemPattern が取得できませんでした。"
        Exit Sub
    End If

    selPattern.Select
    Sleep 500  ' 選択が反映されるのを待つ

    ' 7. コンボボックスを閉じる (Collapse)
    expandPattern.Collapse
    Sleep 500  ' UIの更新待ち

    ' 8. ValuePattern で現在の値を取得して、選択が反映されているかチェック
    Dim valPattern As IUIAutomationValuePattern
    Set valPattern = comboBox.GetCurrentPattern(UIA_ValuePatternId)

    If Not valPattern Is Nothing Then
        If valPattern.CurrentValue = "Web ページ、単一ファイル (*.mhtml)" Then
            MsgBox "ファイルの種類を正しく切り替えました。"
        Else
            MsgBox "選択が正しく反映されていません。現在の値: " & valPattern.CurrentValue
        End If
    Else
        MsgBox "ValuePattern が取得できませんでした。"
    End If
End Sub

