Option Explicit

' セルにチェックボックスを配置し、状態をセル値に反映
Public Sub AddCheckboxesWithLinkedCell()
    Dim wsSheet As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    Dim chkBox As Shape
    Dim chkBoxName As String

    ' シートと範囲を設定
    Set wsSheet = ThisWorkbook.Sheets("Sheet1")
    Set targetRange = wsSheet.Range("B2:B10") ' チェックボックスを配置するセル範囲

    ' 範囲内の各セルにチェックボックスを追加
    For Each cell In targetRange
        ' チェックボックスの名前を設定
        chkBoxName = "CheckBox_" & cell.Row

        ' 既存のチェックボックスがある場合は削除
        On Error Resume Next
        wsSheet.Shapes(chkBoxName).Delete
        On Error GoTo 0

        ' チェックボックスを追加
        Set chkBox = wsSheet.Shapes.AddFormControl(Type:=xlCheckBox, _
            Left:=cell.Left, Top:=cell.Top, Width:=cell.Width, Height:=cell.Height)

        ' チェックボックスの設定
        chkBox.Name = chkBoxName
        chkBox.ControlFormat.Caption = "" ' ラベルを非表示にする
        chkBox.ControlFormat.LinkedCell = cell.Address ' セルと連動
    Next cell

    MsgBox "チェックボックスを追加しました。", vbInformation
End Sub
