#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "sh route vrf GI stat " & chr(124) & " i 0.0.0.0/0" & chr(13)
End Sub
