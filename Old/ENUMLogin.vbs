#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "sudo su -" & chr(13)
	crt.Screen.WaitForString "Password: "
	crt.Screen.Send "Love2Learn." & chr(13)
	crt.Screen.WaitForString "# "
	crt.Screen.Send "bash" & chr(13)
End Sub
