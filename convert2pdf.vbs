' VBScript to convert Word documents to PDF
' Make sure Word is installed on your computer

MsgBox "点击<确定>开始转换...", vbInformation, "提示"

Set fso = CreateObject("Scripting.FileSystemObject")
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False

Set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetParentFolderName(WScript.ScriptFullName)

Set folder = fso.GetFolder(currentPath)
Set files = folder.Files

For Each file In files
    If LCase(fso.GetExtensionName(file)) = "doc" Or LCase(fso.GetExtensionName(file)) = "docx" Then
        docPath = file.Path
        pdfPath = fso.BuildPath(currentPath, fso.GetBaseName(file) & ".pdf")

        Set doc = wordApp.Documents.Open(docPath)
        doc.SaveAs pdfPath, 17 ' 17 is the WdSaveFormat for PDF
        doc.Close
	ElseIf LCase(fso.GetExtensionName(file)) = "xls" Or LCase(fso.GetExtensionName(file)) = "xlsx" Then
            ' 处理Excel文件
            xlsPath = file.Path
            pdfPath = fso.BuildPath(fso.GetParentFolderName(file.Path), fso.GetBaseName(file) & ".pdf")
            Set workbook = excelApp.Workbooks.Open(xlsPath)
            workbook.ExportAsFixedFormat 0, pdfPath ' 0 represents PDF format
            workbook.Close False
    End If
Next

wordApp.Quit
Set wordApp = Nothing
excelApp.Quit
Set excelApp = Nothing
Set fso = Nothing

MsgBox "转换完成!【v1.1.0】" & vbCrLf & "Copyright @ 2024 广东电网审计中心数智审计部 All Rights Reserved.", vbInformation, "提示"
