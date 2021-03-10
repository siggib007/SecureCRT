#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "[CSMBOT11]> "
	crt.Screen.Send "c d d 31" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "ENTER USERNAME < "
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "ENTER PASSWORD < "
	crt.Screen.Send chr(27) & "Alogout" & chr(13)
End Sub
