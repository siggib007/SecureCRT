#$language = "VBScript"
#$interface = "1.0"


strUserName="adm_sbjarna"

'Place a test file in the below location with the linking key on the first line
'Then the IP address of the scanners to change after that, one IP per line.
strInFile = "C:\temp\ScannerList.txt"
Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 5

	If crt.Session.Connected Then
		crt.Session.Disconnect
	end if

 	' crt.Screen.Synchronous = True

	iLoc = instrrev(strInFile,".")
	strOutFile = left(strInFile,iLoc) & "log"
	i = 1

	msgbox("Log will be written to " & strOutFile)

  	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	Set objFileOut  = fso.OpenTextFile(strOutFile, ForWriting, True)
	strKey = objFileIn.readline
	' msgbox("Key=" & strKey)
	While not objFileIn.atendofstream
		strhost = objFileIn.readline
		objFileOut.writeline "Connecting to #" & i & " " & strhost
		cmd = "/SSH2 /ACCEPTHOSTKEYS /L " & strUserName & " " & strhost
		on error resume next
		crt.Session.Connect cmd
		on error goto 0
		If crt.Session.Connected Then
			crt.Screen.Send chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.WaitForString "~]$ "
			crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed status" & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.WaitForString "~]$ "
			crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed unlink --force" & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.WaitForString "~]$ "
			crt.Screen.Send "sudo systemctl stop nessusd" & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli fix --reset" & chr(13)
			crt.Screen.WaitForString "Do you want to proceed? (y/n) [n]: "
			crt.Screen.Send "y" & chr(13)
			crt.Screen.WaitForString  " ~]$ "
			crt.Screen.Send "sudo systemctl start nessusd" & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.Send "sudo systemctl status nessusd" & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed link --cloud --key=" & strKey & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.Send "sudo /opt/nessus/sbin/nessuscli managed status" & chr(13)
			crt.Screen.WaitForString "[" & strUserName & "@"
			crt.Screen.Send "exit" & chr(13)
		  objFileOut.writeline "Completed #" & i & " " & strhost & vbcrlf
      If crt.Session.Connected Then
        crt.Session.Disconnect
      end if
		else
      objFileOut.writeline "Failed to connect to #" & i & " " & strhost & vbcrlf
		i = i + 1
	wend
	i = i - 1
	objFileIn.close
	objFileOut.close
	Set objFileIn  = Nothing
	set objFileOut = Nothing
	Set fso = Nothing

	msgbox "Completed " & i & " hosts, Cleanup complete"

End Sub
