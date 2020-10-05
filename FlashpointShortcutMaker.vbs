' Flashpoint Shortcut Maker by nosamu

scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
fpLauncher = scriptDir & "\Launcher\Flashpoint.exe"
clifp = scriptdir & "\CLIFp.exe"
If Not FileExists(clifp) Then 
	msgbox "Please move this script into your Flashpoint folder. Also make sure that CLIFp.exe is in your Flashpoint folder."
	WScript.Quit
End If
Set wshShell = CreateObject("WScript.Shell")
desktopDir = wshShell.ExpandEnvironmentStrings("%HOMEDRIVE%%HOMEPATH%") & "\Desktop"

shortcutName = InputBox("Enter a name for your shortcut.", "Create Flashpoint Shortcut")
If shortcutName = "" Then WScript.Quit
gameID = InputBox("Enter the ID of the game or animation.", "Create Flashpoint Shortcut")
If gameID = "" Then WScript.Quit

shortcutFile = desktopDir & "\" & shortcutName & ".lnk"
clifpArgs = "--auto " & gameID

Set shortcutObj = wshShell.CreateShortcut(shortcutFile)
	shortcutObj.TargetPath = clifp
	shortcutObj.Arguments = clifpArgs
	shortcutObj.IconLocation = fpLauncher
shortcutObj.Save
msgBox "Shortcut saved!"

Function FileExists(FilePath)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
  Else
    FileExists=CBool(0)
  End If
End Function
