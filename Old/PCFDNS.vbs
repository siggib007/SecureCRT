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
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\PCFList.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\AllPCFDNSSetting.csv" ' The name of the output file, CSV file listing results
const Timeout    = 35    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strARGPair, ObjDNSServer
	dim strOut, strOutPath, iRespone, strDNS_IP, ObjIgnore

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))

	set ObjDNSServer = CreateObject("Scripting.Dictionary")
	set ObjIgnore = CreateObject("Scripting.Dictionary")

	ObjIgnore.add "0.0.0.0",""
	ObjIgnore.add "<none>",""
	ObjIgnore.add "::",""

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
	objFileOut.writeline "PCF,ARG,Primary IPv4,Secondary IPv4,Primary IPv6,Secondary IPv6,Others"

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	Do until objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		strARGPair = strParts(1)
		ObjDNSServer.removeall

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
					ObjDNSServer.add ",Not Connected,timeout waiting for prompt ",""
					crt.Session.Disconnect
				case 1
					ObjDNSServer.add ",Not Connected,Critical Warning on login ",""
					crt.Session.Disconnect
				case 2
					nError = Err.Number
					strErr = Err.Description
					If nError <> 0 Then
						result = "Error " & nError & ": " & strErr
					end if
					crt.Screen.Send("show apn all" & vbcr)
					do while true
						iRespone = crt.Screen.WaitForStrings ("primary dns: ","#",Timeout)
						if iRespone = 2 then
							exit do
						end if
						strDNS_IP=trim(crt.Screen.Readstring("secondary dns:",Timeout))
						if not ObjDNSServer.exists(strDNS_IP) then
							ObjDNSServer.add strDNS_IP,""
						end if
						strDNS_IP=trim(crt.Screen.Readstring(vbcrlf,Timeout))
						if not ObjDNSServer.exists(strDNS_IP) then
							ObjDNSServer.add strDNS_IP,""
						end if
						crt.Screen.WaitForString "ipv6 dns primary server :",Timeout
						strDNS_IP=trim(crt.Screen.Readstring(vbcrlf,Timeout))
						if not ObjDNSServer.exists(strDNS_IP) then
							ObjDNSServer.add strDNS_IP,""
						end if
						crt.Screen.WaitForString "ipv6 dns secondary server :",Timeout
						strDNS_IP=trim(crt.Screen.Readstring(vbcrlf,Timeout))
						if not ObjDNSServer.exists(strDNS_IP) then
							ObjDNSServer.add strDNS_IP,""
						end if
						crt.Screen.WaitForString "IPv6 Primary DNS server address:",Timeout
						strDNS_IP=trim(crt.Screen.Readstring(vbcrlf,Timeout))
						if not ObjDNSServer.exists(strDNS_IP) then
							ObjDNSServer.add strDNS_IP,""
						end if
						crt.Screen.WaitForString "IPv6 Secondary DNS server address:",Timeout
						strDNS_IP=trim(crt.Screen.Readstring(vbcrlf,Timeout))
						if not ObjDNSServer.exists(strDNS_IP) then
							ObjDNSServer.add strDNS_IP,""
						end if
					loop
				case else
					ObjDNSServer.add ",Not Connected,Unexpected choice #" & iRespone, ""
					crt.Session.Disconnect
			end select
			strOut = host & "," & strARGPair
			for each strDNS_IP in ObjDNSServer
				if not ObjIgnore.exists(strDNS_IP) then
					strOut = strOut & "," & strDNS_IP
				end if
			next
			objFileOut.writeline strOut
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
