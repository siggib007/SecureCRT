#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.WaitForString "MMEBOT01#"
	crt.Screen.Send "show int g3/9" & chr(13)
	crt.Screen.WaitForString "MMEBOT01#"
	crt.Screen.Send chr(27) & "[A status" & chr(13)
	crt.Screen.WaitForString "MMEBOT01#"
	crt.Screen.Send "exit" & chr(13)
End Sub
