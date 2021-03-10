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
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGAudit\ARGList012617.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGAudit\ARGIntCount.csv" ' The name of the output file, CSV file listing results


'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 2
	const VerifyCmd = "show interfaces description | include admin-down"

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, bErr, iGigCount, iTenGigCount, iHundredGig
	dim iPrompt, strTemp, iTotalGig, iTotalTenGig, iTotalHundregG

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True
	iTotalGig = 0
	iTotalTenGig = 0
	iTotalHundregG = 0

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "ARG,Gig Ports,TenGig Ports,Hundred Gig Ports"

	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

		bErr = false
		iGigCount = 0
		iTenGigCount = 0
		iHundredGig = 0

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
				iPrompt=crt.Screen.WaitForStrings ("Gi","Te","Hu","%","#", Timeout)
				select case iPrompt
					case 0
						strTemp = "Timeout"
						bErr = True
						exit do
					case 1
						iGigCount = iGigCount + 1
					case 2
						iTenGigCount = iTenGigCount + 1
					case 3
						iHundredGig = iHundredGig + 1
					case 4
						strTemp=trim(crt.Screen.Readstring (vbcrlf,Timeout))
						bErr = True
						exit do
					case 5
						exit do
					case else
						strTemp = "Unexpected choice #" & iPrompt
						bErr = True
						exit do
				end select
			loop
			if bErr = True then
				objFileOut.writeline host & ",Error:" & strTemp
			else
				objFileOut.writeline host & "," & iGigCount & "," & iTenGigCount & "," & iHundredGig
			end if
			crt.Session.Disconnect
		iTotalGig = iTotalGig + iGigCount
		iTotalTenGig = iTotalTenGig + iTenGigCount
		iTotalHundregG = iTotalHundregG + iHundredGig
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.write host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend

	objFileOut.writeline "Grand Total," & iTotalGig & "," & iTotalTenGig & "," & iTotalHundregG
	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

	Set fso = Nothing

	msgbox "All Done, Cleanup complete"

end sub
