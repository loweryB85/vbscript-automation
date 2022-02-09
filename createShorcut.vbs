	'create desktop shortcut to update template in future
	Set Shell = CreateObject("WScript.Shell")
	DesktopPath = Shell.SpecialFolders("Desktop")
	Set link = Shell.CreateShortcut(DesktopPath & "\<name of link>.lnk")
	link.Arguments = "1 2 3"
	link.Description = "<Insert description of link>"
	link.IconLocation = "<Insert path of icon>"
	link.TargetPath = "<Path of Target>"
	link.WorkingDirectory = "<Path of working directory (of target)>"
	link.Save