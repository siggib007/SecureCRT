#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, AuditCmd

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\ASYStorm\AkamaiInts.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\ASYStorm\AkamaiStormControl.csv" ' The name of the output file, CSV file listing results
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, x,y, strTemp, bCont, strInterface
	dim strResultParts, strOut, strOutPath, objDevName, strBaseLine, strTest, strPrefix1, IPAddr, iLineCount, strLastHost
	dim iBLevel, iMLevel, iFound, strDescr, strVlanName, iVlanID, strVersion

	If crt.Session.Connected Then
		crt.Session.Disconnect
	end if

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	strOut = ""
	strLastHost = ""
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

	'Write a header for output file
	objFileOut.writeline "primaryIPAddress,hostName,DeviceType,Port,Description,VLAN_Name,VlanID,BroadCastValue,MulticastValue"

	'Skip over the first header line 
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		strInterface = strParts(1)
		IPAddr = ""
		strDescr = strParts(2)
		strVlanName = strParts(3)
		iVlanID = strParts(4)
		AuditCmd = "show running-config interface " & strInterface
		iBLevel = "none"
		iMLevel = "none"

		if strLastHost <> host then 
			If crt.Session.Connected Then
				crt.Session.Disconnect
			end if

			ConCmd = "/SSH2 /ACCEPTHOSTKEYS "  & host
			on error resume next
			crt.Session.Connect ConCmd
			on error goto 0
			If crt.Session.Connected Then
				crt.Screen.WaitForString "#",Timeout
				nError = Err.Number
				strErr = Err.Description
				If nError <> 0 Then
					result = "Error " & nError & ": " & strErr
				end if
				crt.Screen.Send("term len 0" & vbcr)
				crt.Screen.WaitForString "#",Timeout
				crt.Screen.Send("show ver" & vbcr)
				crt.Screen.WaitForString vbcrlf,Timeout
				strVersion=replace(left(trim(crt.screen.Readstring(vbcr,Timeout)),36),",","")
				crt.Screen.WaitForString "#",Timeout				
			end if
		end if

		If crt.Session.Connected Then
			crt.Screen.Send(AuditCmd & vbcr)
			iFound = crt.Screen.WaitForStrings ("broadcast level ","multicast level ","#",Timeout)
			if iFound = 1 or iFound = 2 then
				strResult=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				if iFound = 1 then iBLevel = strResult
				if iFound = 2 then iMLevel = strResult
				iFound = crt.Screen.WaitForStrings ("broadcast level ","multicast level ","#",Timeout)
				if iFound = 1 or iFound = 2 then
					strResult=trim(crt.Screen.Readstring (vbcrlf,Timeout))
					if iFound = 1 then iBLevel = strResult
					if iFound = 2 then iMLevel = strResult
					crt.Screen.WaitForString "#",Timeout
				end if
			end if

			objFileOut.writeline IPAddr & "," & host & "," & strVersion & "," & strInterface & "," & strDescr & "," & strVlanName & "," & iVlanID & "," & iBLevel & "," & iMLevel
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline IPAddr & "," & host & ",Not Connected,Error " & nError & ": " & strErr
		end if
		strLastHost = host
	wend

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
