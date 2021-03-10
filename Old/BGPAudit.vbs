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
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\BGPTimers\BGPTimerDevices.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\BGPTimers\AuditOut.csv" ' The name of the output file, CSV file listing results
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, strVersion, iLineCount
	dim strResultParts, strOut, strOutPath, x, bGlobalTimer, strLineParts, strNeighbor, strTemp, strTimer, iLASNum

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
	objFileOut.writeline "hostName,LocalAS,OSType,GlobalConfig?,NeighborIP1,NeighborIP2,CurTimer"

	'Skip over the first header line 
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		iLASNum = ""
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

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
			strVersion=replace(left(trim(crt.screen.Readstring(vbcr,Timeout)),36),",","")
			crt.Screen.WaitForString "#",Timeout
			if instr(strVersion,"Cisco IOS")>0 Then
				crt.Screen.Send("show ip protocols summary | i bgp" & vbcr)
				strResult=crt.Screen.WaitForStrings ("bgp ", "#", Timeout)
				If strResult = 1 Then
					iLASNum=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				End If
				crt.Screen.Send("sh run | i timers" & vbcr)
				crt.Screen.WaitForString vbcrlf,Timeout
				strResult=trim(crt.Screen.Readstring ("#",Timeout))
				crt.Session.Disconnect
				strResultParts = split (strResult,vbcrlf)
				iLineCount = ubound(strResultParts)
				strTimer = ""
				strNeighbor = ""
				bGlobalTimer = False
				for x = 0 to iLineCount
					if trim(strResultParts(x)) = "timers bgp 1 3" then
						bGlobalTimer = True
					end if
					strLineParts = split(trim(strResultParts(x))," ")
					if strLineParts(0)="neighbor" then
						strNeighbor = strNeighbor & "," & strLineParts(1)
						if strTimer="" then 
							strTimer = strLineParts(3) & " " & strLineParts(4)
						elseif strTimer <> strLineParts(3) & " " & strLineParts(4) Then
							strTimer = strTimer & "/" & strLineParts(3) & " " & strLineParts(4)
						end if 
					end if 
				next
				objFileOut.writeline host & "," & iLASNum & "," & strVersion & "," & bGlobalTimer & strNeighbor & "," & strTimer
			end if 
			if instr(strVersion,"Cisco Nexus")>0 Then
				crt.Screen.Send("show run | section ""router bgp""" & vbcr)
				crt.Screen.WaitForString vbcrlf,Timeout
				strResult=trim(crt.Screen.Readstring ("#",Timeout))
				crt.Session.Disconnect
				strResultParts = split (strResult,vbcrlf)
				iLineCount = ubound(strResultParts)
				strTemp = ""
				strTimer = ""
				strNeighbor = ""
				bGlobalTimer = False
				for x = 0 to iLineCount
					if trim(strResultParts(x)) = "timers bgp 1 3" then
						bGlobalTimer = True
					end if
					strLineParts = split(trim(strResultParts(x))," ")
					if strLineParts(0)="router" then
						iLASNum = strLineParts(2)
					end if 					
					if strLineParts(0)="neighbor" then
						strTemp = "," & strLineParts(1) & ";" & strLineParts(3)
					end if 
					if strLineParts(0)="timers" then
						if strTimer="" then 
							strTimer = strLineParts(1) & " " & strLineParts(2)
						elseif strTimer <> strLineParts(1) & " " & strLineParts(2) Then
							strTimer = strTimer & "/" & strLineParts(1) & " " & strLineParts(2)
						end if 
						If strTemp = "" Then
							bGlobalTimer = True
							strTimer = strLineParts(2) & " " & strLineParts(3)
						End If
					end if 
					if trim(strResultParts(x)) = "timers 1 3" then
						strNeighbor = strNeighbor & strTemp 
					end if 
				next
				objFileOut.writeline host & "," & iLASNum & "," & strVersion & "," & bGlobalTimer & strNeighbor & "," & strTimer
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
