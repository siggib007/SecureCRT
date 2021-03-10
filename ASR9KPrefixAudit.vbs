#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, strFolder, strPrefixName, iStartCompare

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiPrefix\ARGList032917.csv" ' Input file, comma seperated. format:IP, DeviceName
strOutFile       = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiPrefix\ARG-GI-Out-Audit032917.csv" ' The name of the output file, CSV file listing results
strFolder        = "C:\Users\sbjarna\Documents\IP Projects\Automation\GiPrefix\Romani" ' Folder to save individual prefix sets to
strPrefixName    = "Gi-Out" ' Name of prefix set to look at and compare
iStartCompare    = 1    ' 0 based. 1,2 or 3 recomended. What line in the prefix set should the comparison start. Line 0 is the time stamp at the top of all IOS-XR show run commands.
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.
const CompareAll = False ' Compare prefix sets even if they are different lengths. False is recomended.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, x,y, strTemp, bCont, bOneMissing
	dim strResultParts, strOut, strOutPath, objDevName, strBaseLine, strTest, strPrefix1, IPAddr, VerifyCmd, iLineCount, iCompare

	VerifyCmd = "show run prefix-set " & strPrefixName
	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))
	Set fso = CreateObject("Scripting.FileSystemObject")

	strOut = ""
	if not fso.FileExists(strInFile) Then
		msgbox "Input file " & strInFile & " not found, exiting"
		exit sub
	end if
	if not fso.FolderExists(strFolder) then
		CreatePath (strFolder)
		strOut = strOut & """" & strFolder & """ did not exists so I created it" & vbcrlf
	end if

	if not fso.FolderExists(strOutPath) then
		CreatePath (strOutPath)
		strOut = strOut & vbcrlf & """" & strOutPath & """ did not exists so I created it" & vbcrlf
	end if
	if strOut <> "" then
		msgbox strOut
	end if

	if right(strFolder,1)<>"\" then
		strFolder = strFolder & "\"
	end if

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	' Creating a File System Object to interact with the File System

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	objFileOut.writeline "primaryIPAddress,hostName,1stPrefix,CompareTest"
	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		bOneMissing = False
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		IPAddr = strParts(1)

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
			'crt.Screen.WaitForString "#",Timeout
			crt.Screen.Send(VerifyCmd & vbcr)
			crt.Screen.WaitForString vbcrlf,Timeout
			strResult=trim(crt.Screen.Readstring (vbcrlf&"RP/",Timeout))
			crt.Session.Disconnect
			strResultParts = split (strResult,vbcrlf)
			strTest = ""
			if ubound(strResultParts) > 2 then
				strPrefix1 = trim(strResultParts(2))
			end if
			if right(strPrefix1,1)="," then
				strPrefix1 = left(strPrefix1,len(strPrefix1)-1)
			end if
			if not isarray(strBaseLine) then
				strBaseLine = strResultParts
			end if
			if ubound(strBaseLine) = ubound(strResultParts) then
				bCont = True
				iLineCount = ubound(strBaseLine)
			else
				if CompareAll = True then
					bCont = True
				else
					bCont = False
				end if
				if ubound(strBaseLine) = ubound(strResultParts)+1 and strBaseLine(3) = strResultParts(2) then
					strTest = "First line Missing otherwise "
					bCont = True
					bOneMissing = True
				else
					strTest = "Prefix set length does not match. Baseline is " & ubound(strBaseLine) & " lines, this is " & ubound(strResultParts) & " lines. "
					bOneMissing = False
				end if
				if ubound(strBaseLine) > ubound(strResultParts) then
					iLineCount = ubound(strResultParts)
				else
					iLineCount = ubound(strBaseLine)
				end if
			end if
			if bCont = True then
				strTemp = ""
				if bOneMissing then
					iCompare=3
				else
					iCompare = iStartCompare
				end if
				for x=iCompare to iLineCount
					if bOneMissing Then
						y=x-1
					else
						y=x
					end if
					if strBaseLine(x) <> strResultParts(y) then
						strTemp = strTemp & x-1 & " "
					end if
				next
				' objFileOut.writeline "bCont=True; strTemp=" & strTemp & "; strTest=" & strTest
				if strTemp = "" then
					strTest = strTest & "Pass"
				else
					strTest = strTest & "Prefix " & strTemp & "does not match. "
				end if
			end if
			set objDevName = fso.OpenTextFile(strFolder & host & ".txt", ForWriting, True)
			objDevName.writeline strResult
			objDevName.close
			objFileOut.writeline IPAddr & "," & host & "," & strPrefix1 & "," & strTest
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
