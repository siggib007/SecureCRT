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
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\BGPTimers\BadTimers2.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\BGPTimers\BadTimersAuditBFDversion2.csv" ' The name of the output file, CSV file listing results
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, strOSType, iLineCount, strIPAddr, strLast
	dim strResultParts, strOut, strOutPath, x, strGlobalTimer, strLineParts, strNeighbor, strTimer, iLASNum, iRASNum, strTemplate, bBFD
	dim strVersion, strHardware

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

	'Write a header for output file
	objFileOut.writeline "primaryIPAddress,hostName,LocalAS,Hardware,OSVersion,GlobalConfig,NeighborIP,RemoteAS,CurTimer,BFD"

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		iLASNum = ""
		iRASNum = "n/a"
		strLast = ""
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		if ubound(strParts)>0 then
			strIPAddr = strParts(1)
		else
			strIPAddr = ""
		end if

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
			crt.Screen.Send("term len 0" & vbcr)
			crt.Screen.WaitForString "#",Timeout
			crt.Screen.Send("show ver" & vbcr)
			crt.Screen.WaitForString vbcrlf,Timeout
			strOSType=replace(left(trim(crt.screen.Readstring("(",Timeout)),36),",","")
			strTimer = "default"
			strNeighbor = ""
			strGlobalTimer = "None"
			bBFD = false
			if instr(strOSType,"Cisco IOS ") > 0 or instr(strOSType,"Cisco Internetwork Operating System ") > 0 Then
				crt.Screen.WaitForString " Version ",Timeout
				strVersion = replace(crt.screen.Readstring(" ",Timeout),",","")
				if instr(strOSType,"IOS-XE") > 0 Then
					crt.Screen.WaitForString "Next reboot license Level: ",Timeout
					crt.Screen.WaitForString vbcrlf&vbcrlf,Timeout
					strVersion = "IOS-XE Version " & strVersion
				else	
					crt.Screen.WaitForString "export@cisco.com."&vbcrlf&vbcrlf,Timeout
					strVersion = "IOS Version " & strVersion
				end if
				strHardware = trim(crt.screen.Readstring("(",Timeout))
				crt.Screen.WaitForString "#",Timeout
				crt.Screen.Send("show ip protocols summary | i bgp" & vbcr)
				strResult=crt.Screen.WaitForStrings ("bgp ", "#", Timeout)
				If strResult = 1 Then
					iLASNum=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				End If
				crt.Screen.Send("sh run | i timers|fall-over bfd" & vbcr)
				crt.Screen.WaitForString vbcrlf,Timeout
				strResult=trim(crt.Screen.Readstring ("#",Timeout))
				crt.Session.Disconnect
				strResultParts = split (strResult,vbcrlf)
				iLineCount = ubound(strResultParts)
				for x = 0 to iLineCount
					strLineParts = split(trim(strResultParts(x))," ")
					if ubound(strLineParts)>0 then 
						if strLineParts(0)="timers" and strLineParts(1)="bgp" then
							strGlobalTimer = strLineParts(2) & " " & strLineParts(3)
							strTimer = strLineParts(2) & " " & strLineParts(3)
						end if
					end if
					if strLineParts(0)="neighbor" then
						if strLineParts(2) = "timers" then strTimer = strLineParts(3) & " " & strLineParts(4)
						if strNeighbor = "" then
							strNeighbor = strLineParts(1)
						elseif strNeighbor = strLineParts(1) and strLineParts(2) = "fall-over" then
							bBFD = True
						else
							objFileOut.writeline strIPAddr & "," & host & "," & iLASNum & "," & strHardware & "," & strVersion & "," & strGlobalTimer & "," & strNeighbor & "," & iRASNum & "," & strTimer & "," & bBFD
							strNeighbor = strLineParts(1)
						end if
					end if
				next
				objFileOut.writeline strIPAddr & "," & host & "," & iLASNum & "," & strHardware & "," & strVersion & "," & strGlobalTimer & "," & strNeighbor & "," & iRASNum & "," & strTimer & "," & bBFD
			end if
			if instr(strOSType,"Cisco Nexus") > 0 Then
				crt.Screen.WaitForString " system:    version ",Timeout
				strVersion = crt.screen.Readstring(vbcrlf,Timeout)
				crt.Screen.WaitForString "Hardware"&vbcrlf,Timeout
				strHardware = trim(crt.screen.Readstring("(",Timeout))
				crt.Screen.WaitForString "#",Timeout
				crt.Screen.Send("show run | section ""router bgp""" & vbcr)
				crt.Screen.WaitForString vbcrlf,Timeout
				strResult=trim(crt.Screen.Readstring ("#",Timeout))
				crt.Session.Disconnect
				strResultParts = split (strResult,vbcrlf)
				iLineCount = ubound(strResultParts)
				for x = 0 to iLineCount
					strLineParts = split(trim(strResultParts(x))," ")
					if strLineParts(0)="router" then
						iLASNum = strLineParts(2)
					end if
					if strLineParts(0)="neighbor" then
						if strLast = "Template" then
							strNeighbor = "Template " & strTemplate
						end if
						If strNeighbor <> "" Then
							objFileOut.writeline strIPAddr & "," & host & "," & iLASNum & "," & strHardware & "," & strVersion & "," & strGlobalTimer & "," & strNeighbor & "," & iRASNum & "," & strTimer & "," & bBFD
							strTimer = "default"
							strNeighbor = ""
							strTemplate = ""
							strGlobalTimer = "None"
							bBFD = false
						end if
						strNeighbor = strLineParts(1)
						strLast = "neighbor"
						if ubound(strLineParts) > 2 then iRASNum = strLineParts(3)
					end if
					if strLineParts(0)="template" then
						If strTemplate <> "" Then
							objFileOut.writeline strIPAddr & "," & host & "," & iLASNum & "," & strHardware & "," & strVersion & "," & strGlobalTimer & ",Template " & strTemplate & "," & iRASNum & "," & strTimer & "," & bBFD
							strTimer = "default"
							strNeighbor = ""
							strTemplate = ""
							strGlobalTimer = "None"
							bBFD = false
						end if
						strTemplate = strLineParts(2)
						strLast = "Template"
					end if
					if strLineParts(0)="bfd" then bBFD = True
					if strLineParts(0)="inherit" and strTimer = "default" then
						strTimer = "Template: " & strLineParts(2)
					end if
					if strLineParts(0)="timers" then
						if strLast = "Template" then
							strNeighbor = "Template " & strTemplate
						end if
						If strNeighbor = "" Then
							strGlobalTimer = strLineParts(2) & " " & strLineParts(3)
						else
							strTimer = strLineParts(1) & " " & strLineParts(2)
						End If
					end if
				next
				If strNeighbor <> "" Then
					objFileOut.writeline strIPAddr & "," & host & "," & iLASNum & "," & strHardware & "," & strVersion & "," & strGlobalTimer & "," & strNeighbor & "," & iRASNum & "," & strTimer & "," & bBFD
					strTimer = "default"
					strNeighbor = ""
					strTemplate = ""
					strGlobalTimer = "None"
					bBFD = false
				end if
			end if
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
