' InstallTemplatesOnly.vbs

' Get the objects used by this script.
Dim fso, wsh, checkPath, templatesPath

Set fso = CreateObject("Scripting.FileSystemObject")

Set wsh = WScript.CreateObject("WScript.Shell")

checkPath = wsh.SpecialFolders("Mydocuments") & templatesPath

templatesPath = "\Tecsys\iTopia\Template\"

workingDir = fso.GetFolder(".") & templatesPath

'check for prev install
If Not fso.FolderExists(checkPath) Then

	'MsgBox("Path checked - " & checkPath)

	MsgBox("Path for templates does not exist. Check to see if Itopia add-in is installed.")
	
	Wscript.Quit

End If

'Loop through files in Folder
For Each templateFile In fso.GetFolder(workingDir).Files
	
	'MsgBox("Template File to Copy - " & templateFile)
	
	copyTemplate fso, wsh, templateFile, templatesPath
	
Next

'Inform user of task completion
MsgBox "Templates have been updated."

Wscript.Quit

'sub for copying template from toCopy to copyPath
Sub copyTemplate(fso, wsh, toCopy, copyPath)
	
	Dim source, destination
	
	source = toCopy
	
	'MsgBox("Within copyTemplate - Source: " & source)
	
	destination = wsh.SpecialFolders("MyDocuments") & copyPath
	
	'MsgBox("Within copyTemplate - Dest: " & destination)
	
	fso.copyFile source, destination
	
End Sub