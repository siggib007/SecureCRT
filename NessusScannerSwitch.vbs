#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.

Sub Main
	crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed status" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed unlink --force" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "sudo systemctl stop nessusd" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli fix --reset" & chr(13)
	crt.Screen.WaitForString "Do you want to proceed? (y/n) [n]: "
	crt.Screen.Send "y" & chr(13)
	crt.Screen.WaitForString  "]$ "
	crt.Screen.Send "sudo systemctl start nessusd" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "sudo systemctl status nessusd" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed link --cloud --key=7d78a7fe383df5ca68c38a0371841ff03a07f3bbb6a2017d9f3af527b927bb2c" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed status" & chr(13)
	crt.Screen.WaitForString "]$ "
	crt.Screen.Send "exit" & chr(13)
End Sub
