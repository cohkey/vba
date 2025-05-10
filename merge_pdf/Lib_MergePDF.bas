
Sub MergePDFs_AcrobatCOM()
    ' Acrobat アプリケーション／PDDoc オブジェクトの変数定義
    Dim AcroApp    As Acrobat.CAcroApp
    Dim BaseDoc    As Acrobat.CAcroPDDoc
    Dim AddDoc     As Acrobat.CAcroPDDoc
    Dim outputPath As String
    Dim i          As Long

    ' マージしたい PDF ファイルのパスを配列にセット
    Dim files As Variant
    files = Array("C:\temp\File1.pdf", "C:\temp\File2.pdf", "C:\temp\File3.pdf")

    ' Acrobat を起動し、最初の PDF を開く
    Set AcroApp = CreateObject("AcroExch.App")
    Set BaseDoc = CreateObject("AcroExch.PDDoc")
    If Not BaseDoc.Open(files(0)) Then
        MsgBox "最初の PDF を開けませんでした。", vbCritical
        Exit Sub
    End If

    ' 2 件目以降の PDF を末尾に挿入
    For i = 1 To UBound(files)
        Set AddDoc = CreateObject("AcroExch.PDDoc")
        If AddDoc.Open(files(i)) Then
            BaseDoc.InsertPages BaseDoc.GetNumPages() - 1, AddDoc, 0, AddDoc.GetNumPages(), True
            AddDoc.Close
        End If
    Next i

    ' 保存先を指定してマージ PDF を出力
    outputPath = "C:\temp\Merged_AcrobatCOM.pdf"
    If BaseDoc.Save(PDSaveFull, outputPath) Then
        MsgBox "マージ完了: " & outputPath, vbInformation
    Else
        MsgBox "保存に失敗しました。", vbCritical
    End If

    ' 後始末
    BaseDoc.Close
    AcroApp.Exit
End Sub



Sub MergePDFs_JustPDF()
    Dim pdfList    As String
    Dim outputPath As String
    Dim cmd        As String
    Dim exePath    As String

    ' JustPDF のコマンドラインツール実行ファイルのパス
    exePath = "C:\Program Files\JustSystems\JustPDF\PdfCmdCreator.exe"

    ' 結合したい PDF ファイルをスペース区切りで列挙（各パスはダブルクォートで囲む）
    pdfList = """" & "C:\temp\File1.pdf" & """ """ & "C:\temp\File2.pdf" & """ """ & "C:\temp\File3.pdf" & """"

    ' 出力ファイルのパス
    outputPath = "C:\temp\Merged_JustPDF.pdf"

    ' コマンド文字列を組み立て
    cmd = """" & exePath & """ /combine " & pdfList & " /out """ & outputPath & """"

    ' 非表示で実行
    Shell cmd, vbHide
    MsgBox "JustPDF でマージ処理を起動しました: " & outputPath, vbInformation
End Sub
