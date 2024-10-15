' VBScript to convert Word documents to PDF
' Make sure Word is installed on your computer

Set fso = CreateObject("Scripting.FileSystemObject")
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

Set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetParentFolderName(WScript.ScriptFullName)

Set folder = fso.GetFolder(currentPath)

Sub TraverseFolder(folder)
    ' 遍历当前目录中的所有文件
	For Each file In folder.Files
		If LCase(fso.GetExtensionName(file)) = "doc" Or LCase(fso.GetExtensionName(file)) = "docx" Then
			docPath = file.Path
			pdfPath = fso.BuildPath(fso.GetParentFolderName(file.Path), fso.GetBaseName(file) & ".pdf")

			Set doc = wordApp.Documents.Open(docPath)
			doc.SaveAs pdfPath, 17 ' 17 is the WdSaveFormat for PDF
			doc.Close
		End If
	Next
	' 递归遍历子目录
	For Each subFolder In folder.SubFolders
		TraverseFolder subFolder 
	Next
End Sub

' 开始遍历
TraverseFolder folder
MsgBox "转换完成!" & vbCrLf & "Copyright @ 2024 广东电网审计中心数智审计部 All Rights Reserved."

wordApp.Quit
Set wordApp = Nothing
Set fso = Nothing
