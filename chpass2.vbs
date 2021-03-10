#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.
strInFile = "C:\automate\bolt\mso090919.txt"

function CleanString(strInput)
	strTemp = replace(strInput,vblf,"")
	strTemp = replace(strTemp,vbcr,"")
	strTemp = trim(strTemp)
	CleanString = strTemp
end function

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 5

	If crt.Session.Connected Then
		crt.Session.Disconnect
	end if

	iLoc = instrrev(strInFile,".")
	strOutFile = left(strInFile,iLoc) & "log"
	i = 1

	' msgbox(strOutFile)

	' exit sub

	strPWD = crt.Dialog.Prompt("Enter new password:", "Passwd", "", True)
	if strPWD = "" then exit sub

    Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	Set objFileOut  = fso.OpenTextFile(strOutFile, ForWriting, True)
	strhost = objFileIn.readline
	While not objFileIn.atendofstream
		strhost = objFileIn.readline
		objFileOut.writeline "Connecting to #" & i & " " & strhost
		cmd = "/SSH2 /ACCEPTHOSTKEYS /L cloud-user " & strhost
		crt.Session.Connect cmd
		' crt.Session.Log False
		crt.Screen.Send "sudo passwd root" & vbcr
		crt.Screen.WaitForString "New password: "
		crt.Screen.Send strPWD & vbcr
		strResults = crt.Screen.ReadString("Retype new password: ")
		strResults = CleanString(strResults)
		if strResults <> "" then
			objFileOut.writeline strResults
		end if
		crt.Screen.Send strPWD & vbcr
		strResults = crt.Screen.ReadString("Retype new password: ")
		strResults = CleanString(strResults)
		if strResults <> "" then
			objFileOut.writeline strResults
		end if
		' crt.Screen.WaitForString "Retype new password: "
		crt.Screen.Send strPWD & vbcr
		crt.Screen.WaitForString vbcr
		strResults = crt.Screen.ReadString("~]$")
		strResults = replace(strResults,"[cloud-user@","")
		strResults = CleanString(strResults)
		objFileOut.writeline strResults
		crt.Screen.Send "exit" & vbcr
		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if
		objFileOut.writeline "Completed #" & i & " " & strhost & vbcrlf
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
