#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 12/19/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, dictSubnets

' User Spefified values, specify values here per your needs
strInFile     = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\ARGList012617.csv"
strOutFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\ARGACLs\ARGACLAudit031717.csv"


'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 5

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, iPrompt, strResultParts, strLineParts
	dim strOut, strOutPath, IPAddr, objACLDict, strACL, strInterface, iLineCount, strIntDescr, strIPVer, strInt, x, strIntIP

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	set objACLDict = CreateObject("Scripting.Dictionary")
	set dictSubnets = CreateObject("Scripting.Dictionary")

	InitializeDicts

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

	' Creating a File System Object to interact with the File System

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "primaryIPAddress,hostName,ACL Name,Line Count,Interface, Interface description,IP Address"
	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		objACLDict.RemoveAll
		strInterface=""
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		IPAddr = strParts(1)

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
			crt.Screen.Send("show running-config | include access-list" & vbcr)
			crt.Screen.WaitForString vbcr,Timeout
			do while true
				strResult=trim(crt.Screen.Readstring (vbcrlf,"#",Timeout))
				if crt.Screen.MatchIndex = 2 or crt.Screen.MatchIndex = 0 then
					exit do
				end if
				strResultParts = split (strResult," ")
				if ubound(strResultParts) = 2 then
					if strResultParts(1) = "access-list" then
						strIPVer = strResultParts(0)
						strACL = strResultParts(2)
						if not objACLDict.exists(strACL) then
							objACLDict.add strACL, strIPVer
						end if
					end if
				end if
			loop
			for each strACL in objACLDict
				strInterface = ""
				crt.Screen.Send("show access-lists " & objACLDict(strACL) & " " & strACL & " | utility wc lines" & vbcr)
				crt.Screen.WaitForString "GMT" & vbcrlf,Timeout
				iLineCount=trim(crt.Screen.Readstring (vbcrlf,Timeout))
				crt.Screen.WaitForString "#",Timeout
				crt.Screen.Send("show access-lists " & objACLDict(strACL) & " " & strACL & "  usage pfilter location all " & vbcr)
				do while true
					iPrompt=crt.Screen.WaitForStrings ("Interface :", "#", Timeout)
					select case iPrompt
						case 0
							objFileOut.writeline IPAddr & "," & host & "," & strACL & ",Timeout"
						case 2
							exit do
						case 1
							strInterface=strInterface & "|" & trim(crt.Screen.Readstring (vbcrlf,Timeout))
						case else
							msgbox "Unexpected choice #" & iPrompt
							exit do
					end select
				loop
				if strInterface = "" then
			    	if strACL <> "3" and strACL <> "4" then
						crt.Screen.Send("show running-config | include " & strACL & vbcr)
						crt.Screen.WaitForStrings "configuration...",Timeout
						strResult=trim(crt.Screen.Readstring ("RP/0/R",Timeout))
						strResultParts = split(strResult,vbcrlf)
						for x=0 to ubound(strResultParts)
							if InStr(strResultParts(x),"access-list")=0 and strResultParts(x)<>"" then
								strLineParts=split(trim(strResultParts(x))," ")
								strInterface=strLineParts(0)
								if strInterface = "access-class" then
									strInterface = "line default"
								end if
							end if
						next
					else
						strInterface = "Assumed NTP"
					end if
					if strInterface = "" then
						strInterface = "unbound"
					end if
					objFileOut.writeline IPAddr & "," & host & "," & strACL & "," & iLineCount & "," & strInterface
				else
					strResultParts = split (strInterface,"|")
					for x=0 to ubound(strResultParts)
						strInt = strResultParts(x)
						if strInt <> "" then
							crt.Screen.Send("show running-config interface " & strInt & vbcr)
							crt.Screen.WaitForStrings "description",Timeout
							strIntDescr=trim(crt.Screen.Readstring (vbcrlf,Timeout))
							strIntDescr=replace(strIntDescr,",",";")
							crt.Screen.WaitForStrings "address",Timeout
							strIntIP=trim(crt.Screen.Readstring (vbcrlf,Timeout))
							strLineParts=split(strIntIP," ")
							if ubound(strLineParts)=1 then
								if dictSubnets.exists(strLineParts(1)) then
									strIntIP = strLineParts(0) & dictSubnets.Item(strLineParts(1))
								else
									strIntIP = strLineParts(0) & "***" & strLineParts(1) & "***"
								end if
							end if
							objFileOut.writeline IPAddr & "," & host & "," & strACL & "," & iLineCount & "," & strInt & "," & strIntDescr & "," & strIntIP
						end if
					next
				end if
			next
			crt.Session.Disconnect
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline IPAddr & "," & host & ",Not Connected,Error " & nError & ": " & strErr
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

sub InitializeDicts
	dictSubnets.add "255.255.255.255", "/32"
	dictSubnets.add "255.255.255.254", "/31"
	dictSubnets.add "255.255.255.252", "/30"
	dictSubnets.add "255.255.255.248", "/29"
	dictSubnets.add "255.255.255.240", "/28"
	dictSubnets.add "255.255.255.224", "/27"
	dictSubnets.add "255.255.255.192", "/26"
	dictSubnets.add "255.255.255.128", "/25"
	dictSubnets.add "255.255.255.0", "/24"
	dictSubnets.add "255.255.254.0", "/23"
	dictSubnets.add "255.255.252.0", "/22"
	dictSubnets.add "255.255.248.0", "/21"
	dictSubnets.add "255.255.240.0", "/20"
	dictSubnets.add "255.255.224.0", "/19"
	dictSubnets.add "255.255.192.0", "/18"
	dictSubnets.add "255.255.128.0", "/17"
	dictSubnets.add "255.255.0.0", "/16"
	dictSubnets.add "255.254.0.0", "/15"
	dictSubnets.add "255.252.0.0", "/14"
	dictSubnets.add "255.248.0.0", "/13"
	dictSubnets.add "255.240.0.0", "/12"
	dictSubnets.add "255.224.0.0", "/11"
	dictSubnets.add "255.192.0.0", "/10"
	dictSubnets.add "255.128.0.0", "/9"
	dictSubnets.add "255.0.0.0", "/8"
	dictSubnets.add "254.0.0.0", "/7"
	dictSubnets.add "252.0.0.0", "/6"
	dictSubnets.add "248.0.0.0", "/5"
	dictSubnets.add "240.0.0.0", "/4"
	dictSubnets.add "224.0.0.0", "/3"
	dictSubnets.add "192.0.0.0", "/2"
	dictSubnets.add "128.0.0.0", "/1"
end sub