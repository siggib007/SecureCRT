#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.
strInFile = "C:\Temp\testips.txt"

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 5

	passwd = crt.Dialog.Prompt("Enter new password:", "Passwd", "", True)
	if strPWD = "" then exit sub

    Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strhost = objFileIn.readline
		cmd = "/SSH2 /ACCEPTHOSTKEYS /L cloud-user " & strhost
		crt.Session.Connect cmd
		crt.Session.Log False
		crt.Screen.Send "sudo -i" & chr(13)
		crt.Screen.WaitForString "]# "
		crt.Screen.Send "echo -e " & chr(34) & passwd & "\n" & passwd & chr(34) & "| (passwd --stdin root)" & chr(13)
		crt.Screen.WaitForString "]# "
		crt.Screen.Send "exit" & chr(13)
		crt.Screen.WaitForString  "]$ "
		crt.Screen.Send "exit" & chr(13)
		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if
	wend
	objFileIn.close
	Set objFileIn  = Nothing
	Set fso = Nothing

	msgbox "All Done, Cleanup complete"
End Sub
