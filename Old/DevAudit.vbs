#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, strFolder, AuditCmd, user

' User Spefified values, specify values here per your needs
strInFile        = "C:\Users\sbjarna\Documents\VMAS\FWGList.csv" ' Input file, comma seperated. First value device name, first line header
strOutFile       = "C:\Users\sbjarna\Documents\VMAS\FWGPools.csv" ' The name of the output file, CSV file listing results
strFolder        = "C:\Users\sbjarna\Documents\VMAS\FWGPoolOut" ' Folder to save individual command output to
AuditCmd = "show run all-partitions | include ip nat pool "
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.
const CompareAll = True ' Compare prefix sets even if they are different lengths. False is recomended.
user = "sbjarna"

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, strResultParts, strOut, strOutPath, objDevName, iLineCount, passwd

	strOutPath = left (strOutFile, InStrRev (strOutFile,"\"))

	' Creating a File System Object to interact with the File System
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

	passwd = crt.Dialog.Prompt("Enter " & user & "'s password for " & host, "Login", "", True)
	'Opening both intput and output files
	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)

	'Skip over the first header line
	strLine = objFileIn.readline
	'Start Looping through the input file
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)

		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if

		ConCmd = "/SSH2 /ACCEPTHOSTKEYS /L " & user & " /PASSWORD " & passwd & " "  & host
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
			' crt.Screen.Send("term len 0" & vbcr)
			' crt.Screen.WaitForString "#",Timeout
			crt.Screen.Send(AuditCmd & vbcr)
			crt.Screen.WaitForString vbcrlf,Timeout
			strResult=trim(crt.Screen.Readstring ("#",Timeout))
			crt.Session.Disconnect

			strResultParts = split(strResult,vbcrlf)
			iLineCount = ubound(strResultParts)

			set objDevName = fso.OpenTextFile(strFolder & host & ".txt", ForWriting, True)
			objDevName.writeline strResult
			objDevName.close
			objFileOut.writeline host & "," & iLineCount
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
