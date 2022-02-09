' Installaddin.vbs

' Get the objects used by this script.
Dim excelObj, oAddin, fso, wsh, srcPath, destPath, addin

addin = "\iTopia\Add-In\Tecsys_iTopia.xlam"

Set excelObj = CreateObject("Excel.Application")

Set fso = CreateObject("Scripting.FileSystemObject")

Set wsh = WScript.CreateObject("WScript.Shell")

' Make Excel visible in case something goes wrong.
excelObj.Visible = True

' Create a temporary workbook (required to access add-ins)
excelObj.Workbooks.Add

' Get the current folder.
srcPath = fso.GetFolder(".") & "\Tecsys"

destPath = wsh.SpecialFolders("Mydocuments") & "\Tecsys"

'check for prev install
If Not fso.FolderExists(destPath) Then

fso.CreateFolder(destPath)

Else

MsgBox "Folder already exists. Check for previous install."
excelObj.Quit
WScript.Quit

End If

' Copy the file to the template folder.
fso.CopyFolder srcPath, destPath

' Add the add-in to Excel.
Set oAddin = excelObj.AddIns.Add(destpath & addin, true)

' Mark the add-in as installed so Excel loads it.
oAddin.Installed = True

'Inform user of task completion
MsgBox "Add-In Installed. Excel will now close." 

' Close Excel.
excelObj.Quit
Set excelObj = Nothing

