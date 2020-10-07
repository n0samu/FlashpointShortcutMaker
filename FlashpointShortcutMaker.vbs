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

shortcutName = ""
Do While Not IsKosherFilename(shortcutName) 
	shortcutName = InputBox("Enter a name for your shortcut.", "Create Flashpoint Shortcut")
	If shortcutName = "" Then
		WScript.Quit
	ElseIf Not IsKosherFilename(shortcutName) Then
		msgbox "A shortcut cannot contain any of the following characters: \/:*?""<>| Try again."
    End If
Loop

gameID = ""
Do While Not IsUUID(gameID) 
	gameID = InputBox("Enter the ID of the game or animation.", "Create Flashpoint Shortcut")
	If gameID = "" Then 
		WScript.Quit
	ElseIf Not IsUUID(gameID) Then
		msgbox "Not a valid game/animation ID! Try again."
	End If
Loop

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

Function IsUUID(strUUID)
	If isnull(strUUID) Then
		IsUUID = False
		Exit Function
	End If
	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "^({|\()?[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}(}|\))?$"
	IsUUID = regEx.Test(strUUID)
	Set RegEx = Nothing
End Function

Function IsKosherFilename(fileName)
    If fileName = "" Then
        IsKosherFilename = false
		Exit Function
    End If

	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "[\\\/:*?""<>|]"
	IsKosherFilename = NOT regEx.Test(fileName)
	Set RegEx = Nothing
End Function
