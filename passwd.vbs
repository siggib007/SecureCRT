#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "sudo passwd root" & chr(13)
	crt.Screen.WaitForString "New password: "
	crt.Screen.Send "" & chr(13)
	crt.Screen.WaitForString "Retype new password: "
	crt.Screen.Send "" & chr(13)
	crt.Screen.WaitForString "Retype new password: "
	crt.Screen.Send "" & chr(13)
End Sub
