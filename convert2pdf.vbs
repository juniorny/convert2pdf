' VBScript to convert Word documents to PDF
' Make sure Word is installed on your computer

Set fso = CreateObject("Scripting.FileSystemObject")
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

Set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetParentFolderName(WScript.ScriptFullName)

Set folder = fso.GetFolder(currentPath)

Sub TraverseFolder(folder)
    ' ������ǰĿ¼�е������ļ�
	For Each file In folder.Files
		If LCase(fso.GetExtensionName(file)) = "doc" Or LCase(fso.GetExtensionName(file)) = "docx" Then
			docPath = file.Path
			pdfPath = fso.BuildPath(fso.GetParentFolderName(file.Path), fso.GetBaseName(file) & ".pdf")

			Set doc = wordApp.Documents.Open(docPath)
			doc.SaveAs pdfPath, 17 ' 17 is the WdSaveFormat for PDF
			doc.Close
		End If
	Next
	' �ݹ������Ŀ¼
	For Each subFolder In folder.SubFolders
		TraverseFolder subFolder 
	Next
End Sub

' ��ʼ����
TraverseFolder folder
MsgBox "ת�����!" & vbCrLf & "Copyright @ 2024 �㶫�����������������Ʋ� All Rights Reserved."

wordApp.Quit
Set wordApp = Nothing
Set fso = Nothing
