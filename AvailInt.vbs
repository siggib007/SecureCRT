#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 01/27/17                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\ASAList.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\ASA-Avail-Ints.csv" ' The name of the output file, CSV file listing results


'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 2
	const VerifyCmd = "show interface status | include disable"

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError
	dim iPrompt, strTemp, strErr, bErr, bFound, strTempParts

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "Device,interface"

	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

		bErr = false
		bFound = false

		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if

		ConCmd = "/SSH2 /ACCEPTHOSTKEYS "  & host
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
				iPrompt=crt.Screen.WaitForStrings (vbcrlf,"%","#", Timeout)
				' objFileOut.writeline host & ",Promt:" & iPrompt
				select case iPrompt
					case 0
						strTemp = "Timeout"
						bErr = True
						' exit do
					case 1
						strTemp=trim(crt.Screen.Readstring (vbcrlf,"#",Timeout))
						if crt.Screen.MatchIndex = 2 then exit do
						if instr(strTemp,"base-T") > 0 or instr(strTemp,"10/100") > 0 then
							bFound = True
							strTempParts = split(strTemp," ")
							strTemp = strTempParts(0)
						else
							strTemp = ""
						end if
					case 2
						strTemp=trim(crt.Screen.Readstring (vbcrlf,Timeout))
						bErr = True
						' exit do
					case 3
						if bFound then
							exit do
						else
							strTemp = "found nothing"
							bErr = True
						end if
					case else
						strTemp = "Unexpected choice #" & iPrompt
						bErr = True
						' exit do
				end select
				if bErr = True then
					objFileOut.writeline host & ",Error:" & strTemp
					exit do
				else
					if strTemp <> "" then objFileOut.writeline host & "," & strTemp
				end if
			loop
			crt.Session.Disconnect
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.write host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend

	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

	Set fso = Nothing

	msgbox "All Done, Cleanup complete"

end sub
