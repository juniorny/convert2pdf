Option Explicit

Dim fso, excelApp, currentPath, folder, files, file
Dim docPath, pdfPath, wordApp, doc, xlsPath, workbook

MsgBox "点击<确定>开始转换...", vbInformation, "提示"

Set fso = CreateObject("Scripting.FileSystemObject")
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False

currentPath = fso.GetParentFolderName(WScript.ScriptFullName)
Set folder = fso.GetFolder(currentPath)
Set files = folder.Files

For Each file In files
    Dim ext
    ext = LCase(fso.GetExtensionName(file))
    
    If ext = "doc" Or ext = "docx" Then
        ' 对于每个 Word 文档，单独启动 Word 实例
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        wordApp.DisplayAlerts = 0
        
        docPath = file.Path
        pdfPath = fso.BuildPath(currentPath, fso.GetBaseName(file) & ".pdf")
        
        On Error Resume Next
        Set doc = wordApp.Documents.Open(docPath, , True)
        If Err.Number <> 0 Then
            WScript.Echo "无法打开 Word 文件: " & docPath & " 错误: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            WScript.Sleep 1000
            doc.ExportAsFixedFormat pdfPath, 17
            If Err.Number <> 0 Then
                WScript.Echo "导出PDF失败: " & docPath & " 错误: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            doc.Close False
        End If
        On Error GoTo 0
        
        wordApp.Quit
        Set wordApp = Nothing
        
    ElseIf ext = "xls" Or ext = "xlsx" Then
        xlsPath = file.Path
        pdfPath = fso.BuildPath(fso.GetParentFolderName(xlsPath), fso.GetBaseName(file) & ".pdf")
        
        On Error Resume Next
        Set workbook = excelApp.Workbooks.Open(xlsPath)
        If Err.Number <> 0 Then
            WScript.Echo "无法打开 Excel 文件: " & xlsPath & " 错误: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            workbook.ExportAsFixedFormat 0, pdfPath
            If Err.Number <> 0 Then
                WScript.Echo "导出Excel PDF失败: " & xlsPath & " 错误: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            workbook.Close False
        End If
        On Error GoTo 0
    End If
Next

excelApp.Quit
Set excelApp = Nothing
Set fso = Nothing

MsgBox "转换完成!" & vbCrLf & "作者：时钟震动宇宙 小红书539345873", vbInformation, "提示"
