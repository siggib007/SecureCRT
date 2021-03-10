#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "term len 0" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "show vlan br" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "sh ip int br" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "sh cdp nei" & chr(13)
End Sub
