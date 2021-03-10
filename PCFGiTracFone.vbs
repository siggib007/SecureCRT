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
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\PCFList7.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\PCFAudit7-17.csv" ' The name of the output file, CSV file listing results
const Timeout    = 35    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strARGPair
	dim strOut, strOutPath, iRespone, strTracFone, strGiV4, strGiV6, bContinue, strTwilio, strCintex

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	strOut = ""
	if not fso.FileExists(strInFile) Then
		msgbox "Input file " & strInFile & " not found, exiting"
		exit sub
	end if

	if not fso.FolderExists(strOutPath) then
		CreatePath (strOutPath)
		strOut = strOut & vbcrlf & """" & strOutPath & """ did not exists so I created it" & vbcrlf
	end if
	if strOut <> "" then
		msgbox strOut
	end if

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	'Opening both intput and output files
	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	objFileOut.writeline "PCF,ARG,Gi v4,Gi v6,TracFone,Twilio,Cintext"

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	Do until objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		strARGPair = strParts(1)
		strTracFone = "--"
		strGiV4 = "--"
		strGiV6 = "--"
		bContinue=True

		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if

		ConCmd = "/SSH2 /ACCEPTHOSTKEYS "  & host
		on error resume next
		crt.Session.Connect ConCmd
		on error goto 0

		If crt.Session.Connected Then
			crt.Screen.Synchronous = True
			iRespone = crt.Screen.WaitForStrings ("WARNING","#",Timeout)
			select case iRespone
				case 0
					objFileOut.writeline host & ",Not Connected,timeout waiting for prompt "
					crt.Session.Disconnect
				case 1
					objFileOut.writeline host & ",Not Connected,Critical Warning on login "
					crt.Session.Disconnect
				case 2
					nError = Err.Number
					strErr = Err.Description
					If nError <> 0 Then
						result = "Error " & nError & ": " & strErr
					end if
					crt.Screen.Send("context gi" & vbcr)
					crt.Screen.WaitForString "#",Timeout
					crt.Screen.Send("show ip interface name tracfone_loopback" & vbcr)
					iRespone = crt.Screen.WaitForStrings ("IP Address: ","does not exist!",Timeout)
					if iRespone = 1 then
						strTracFone=trim(crt.Screen.Readstring("Subnet Mask:",Timeout))
					else
						strTracFone = "None"
					end if
					crt.Screen.WaitForString "#",Timeout
					crt.Screen.Send("show ip interface name twilio_loopback" & vbcr)
					iRespone = crt.Screen.WaitForStrings ("IP Address: ","does not exist!",Timeout)
					if iRespone = 1 then
						strTwilio=trim(crt.Screen.Readstring("Subnet Mask:",Timeout))
					else
						strTwilio = "None"
					end if
					crt.Screen.WaitForString "#",Timeout
					crt.Screen.Send("show ip interface name radius_nas_cintex" & vbcr)
					iRespone = crt.Screen.WaitForStrings ("IP Address: ","does not exist!",Timeout)
					if iRespone = 1 then
						strCintex=trim(crt.Screen.Readstring("Subnet Mask:",Timeout))
					else
						strCintex = "None"
					end if
					crt.Screen.WaitForString "#",Timeout
					crt.Screen.Send("show ip interface name gi_loopback" & vbcr)
					iRespone = crt.Screen.WaitForStrings ("IP Address: ",Timeout)
					if iRespone = 1 then
						strGiV4=trim(crt.Screen.Readstring("Subnet Mask:",Timeout))
					else
						strGiV4 = "None"
					end if
					crt.Screen.WaitForString "#",Timeout
					crt.Screen.Send("show ipv6 interface name gi_ipv6_loopback" & vbcr)
					iRespone = crt.Screen.WaitForStrings ("Unicast Address: ",Timeout)
					if iRespone = 1 then
						strGiV6=trim(crt.Screen.Readstring(vbcrlf,Timeout))
					else
						strGiV6 = "does not exist!"
					end if
					objFileOut.writeline host & "," & strARGPair & "," & strGiV4 & "," & strGiV6 & "," & strTracFone & "," & strTwilio & "," & strCintex
				case else
					objFileOut.write host & ",Not Connected,Unexpected choice #" & iRespone
					crt.Session.Disconnect
			end select
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.write host & ",Not Connected,Error " & nError & ": " & strErr
		end if
		crt.Session.Disconnect
	Loop

	objFileOut.close
	objFileIn.close
	Set objFileIn  = Nothing
	Set objFileOut = Nothing

	Set fso = Nothing

	msgbox "All Done, Cleanup complete"

end sub

Function CreatePath (strFullPath)
'-------------------------------------------------------------------------------------------------'
' Function CreatePath (strFullPath)                                                               '
'                                                                                                 '
' This function takes a complete path as input and builds that path out as nessisary.             '
'-------------------------------------------------------------------------------------------------'
dim pathparts, buildpath, part, fso

Set fso = CreateObject("Scripting.FileSystemObject")

	pathparts = split(strFullPath,"\")
	buildpath = ""
	for each part in pathparts
		if buildpath<>"" then
			if buildpath = "\" then
				buildpath = buildpath & part
			else
				buildpath = buildpath & "\" & part
			end if
			if not fso.FolderExists(buildpath) then
				fso.CreateFolder(buildpath)
			end if
		else
			if part="" then
				buildpath = "\"
			else
				buildpath = part
			end if
		end if
	next
end function
