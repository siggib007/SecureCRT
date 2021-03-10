#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "juniper" & chr(13)
	crt.Screen.WaitForString "Password:"
	crt.Screen.Send "Clouds" & chr(13)
	crt.Screen.WaitForString "> "
	crt.Screen.Send "set cli screen-length 0" & chr(13)
	crt.Screen.Send "set cli screen-width 200" & chr(13)
End Sub
