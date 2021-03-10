#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "cd /home/tenable-io01/" & chr(13)
	crt.Screen.WaitForString "tenable-io01]$ "
	crt.Screen.Send "source magentarc " & chr(13)
End Sub
