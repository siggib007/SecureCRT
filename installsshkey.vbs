#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "mkdir .ssh; wget http://patools.gsm1900.org/sbjarna -O .ssh/authorized_keys" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "exit" & chr(13)
End Sub
