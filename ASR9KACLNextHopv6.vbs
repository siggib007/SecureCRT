#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile

' User Spefified values, specify values here per your needs
strInFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\OMWIPv6ACL\SL3-ARGlist1.csv"
strOutFile   = "C:\Users\sbjarna\Documents\IP Projects\Automation\OMWIPv6ACL\SL3-ARGIPv6NextHop1.csv"

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 2
	const VerifyCmd = "show run ipv6 access-list "
	const FirstCol = 4

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult
	dim strResultParts, strNextHop, objNetHop, iPrompt, strACLName, bcont

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	set objNetHop = CreateObject("Scripting.Dictionary")
	bcont = True

	objFileOut.writeline "Device,ACL,Next Hop"

	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		strNextHop = ""
		strACLName = "n/a"

		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if

		ConCmd = "/SSH2 "  & host
		on error resume next
		crt.Session.Connect ConCmd
		on error goto 0

		If crt.Session.Connected Then
			crt.Screen.Synchronous = True
			crt.Screen.WaitForString "#",Timeout
			nError = Err.Number
			strErr = Err.Description
			If nError <> 0 Then
				result = "Error " & nError & ": " & strErr
			end if
			crt.Screen.Send("term len 0" & vbcr)
			crt.Screen.WaitForString "#",Timeout
			crt.Screen.Send(VerifyCmd & vbcr)
			do while true
				iPrompt=crt.Screen.WaitForStrings ("ipv6 access-list ", "nexthop1 ", "#", Timeout)
				select case iPrompt
					case 0
						' msgbox "Timeout"
						' bcont=false
						exit do
					case 3
						' msgbox "Found prompt"
						exit do
					case 1
						strACLName=trim(crt.Screen.Readstring (vbcrlf,Timeout))
						objNetHop.RemoveAll
					case 2
						strNextHop=trim(crt.Screen.Readstring (vbcrlf,Timeout))
						if not objNetHop.exists(strNextHop) then
							objNetHop.add strNextHop,""
							objFileOut.writeline host & "," & strACLName  & "," & strNextHop
						end if
					case else
						msgbox "Unexpected choice #" & iPrompt
						exit do
				end select
			loop
			crt.Session.Disconnect
			if strNextHop = "" then objFileOut.writeline host & "," & strACLName  & ",-"
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend

	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

	Set fso = Nothing

	msgbox "All Done, Cleanup complete"

end sub
