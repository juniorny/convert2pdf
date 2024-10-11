' VBScript to convert Word documents to PDF
' Make sure Word is installed on your computer

Set fso = CreateObject("Scripting.FileSystemObject")
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

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
    End If
Next

MsgBox "转换完成!" & vbCrLf & "Copyright @ 2024 广东电网审计中心数智审计部 All Rights Reserved."

wordApp.Quit
Set wordApp = Nothing
Set fso = Nothing
