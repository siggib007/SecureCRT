#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, strFolder, strACLName, iStartCompare

' User Spefified values, specify values here per your needs

' Input file, comma seperated.
' First line needs to be "ACL Name," followed by the name of the ACL you want audited.
' First cell (i.e. A1 or R1C1) isn't important, the script looks for the ACL name in R1C2 (or B1).
' Also IP version should be indicated in R1C4 and you can force full comparison despite ACL length mismatch by putting yes or true in R1C6
' Second line should be header and is ignored by the script
' remaining lines should be a list of ASR9K's to audit. Format:DeviceName, IP Address.
' User is prompted for the input file via file browser dialog.

iStartCompare    = 1    ' 0 based. 1,2 or 3 recomended. What line in the ACL should the comparison start. Line 0 is the time stamp at the top of all IOS-XR show run commands.
const Timeout    = 5    ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	' File handling constants
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8

	' button parameter options
	Const ICON_STOP = 16                 ' display the ERROR/STOP icon.
	Const ICON_QUESTION = 32             ' display the '?' icon
	Const ICON_WARN = 48                 ' display a '!' icon.
	Const ICON_INFO= 64                  ' displays "info" icon.
	Const BUTTON_OK = 0                  ' OK button only
	Const BUTTON_CANCEL = 1              ' OK and Cancel buttons
	Const BUTTON_ABORTRETRYIGNORE = 2    ' Abort, Retry, and Ignore buttons
	Const BUTTON_YESNOCANCEL = 3         ' Yes, No, and Cancel buttons
	Const BUTTON_YESNO = 4               ' Yes and No buttons
	Const BUTTON_RETRYCANCEL = 5         ' Retry and Cancel buttons

	Const DEFBUTTON1 = 0        ' First button is default
	Const DEFBUTTON2 = 256      ' Second button is default
	Const DEFBUTTON3 = 512      ' Third button is default

	' Possible MessageBox() return values
	Const IDOK = 1              ' OK button clicked
	Const IDCANCEL = 2          ' Cancel button clicked
	Const IDABORT = 3           ' Abort button clicked
	Const IDRETRY = 4           ' Retry button clicked
	Const IDIGNORE = 5          ' Ignore button clicked
	Const IDYES = 6             ' Yes button clicked
	Const IDNO = 7              ' No button clicked

	dim strParts, strLine, objFileIn, objFileOut, host, ConCmd, fso, nError, strErr, strResult, x,y, strTemp, bCont, bBase, strIPVer, bCompareAll, strCompare
	dim strResultParts, strOut, strOutPath, objDevName, strBaseLine, strTest, IPAddr, VerifyCmd, iLineCount, iCompare, iResult, iLastLine, bRange

	bCompareAll = False ' Compare ACL even if they are different lengths. False is recomended.

	strInFile = crt.Dialog.FileOpenDialog("Please select CSV input file", "Open", "", "CSV Files (*.csv)|*.csv||")

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	strLine = objFileIn.readline
	strParts = split(strLine,",")
	if ubound(strParts) >= 3 then
		strACLName = strParts(1)
		strIPVer = strParts(3)
	else
		if ubound(strParts)>=1 Then
			strACLName = strParts(1)
			iResult = crt.Dialog.MessageBox("Is this an IPv4 ACL? Yes=IPv4 No=IPv6", "Is ACL IPv4 or IPv6", ICON_QUESTION Or BUTTON_YESNO Or DEFBUTTON1 )
			If iResult = IDYES Then
				strIPVer = "ipv4"
			else
				strIPVer = "ipv6"
			end if
		else
			msgbox "Please include the ACL you want to audit along with IP version (ipv4/ipv6) in the first line of the CSV file. " _
		 	& "Please make sure the ACL name is in second field and the IP version is in the fourth field of the first line "
			exit sub
		end if
	end if
	if strACLName = "" Then
		msgbox "Please include the ACL you want to audit in the first line of the CSV file. Please make sure it is in second field of the first line."
		exit sub
	end if
	if ubound(strParts) >= 5 then
		strCompare = lcase(strParts(5))
		if strCompare = "yes" or strCompare = "true" then
			bCompareAll = True
			msgbox "Per R1C6 in the CSV file, forcing full comparison regardless of ACL length"
		end if
	end if

	select case right(strIPVer,1)
		case "4"
			strIPVer = "ipv4"
		case "6"
			strIPVer = "ipv6"
		case else
			iResult = crt.Dialog.MessageBox("IP version '" & strIPVer & "' is not recognized. Is this an IPv4 ACL? Yes=IPv4 No=IPv6", "Is ACL IPv4 or IPv6", ICON_QUESTION Or BUTTON_YESNO Or DEFBUTTON1 )
			If iResult = IDYES Then
				strIPVer = "ipv4"
			else
				strIPVer = "ipv6"
			end if
	end select

	if strACLName = "" Then
		msgbox "Please include the ACL you want to audit in the first line of the CSV file. Please make sure it is in second field of the first line."
		exit sub
	else
		iResult = crt.Dialog.MessageBox("Confirm to Audit ACL " & strACLName & " and it is " & strIPVer & "?", "Confirmation", ICON_QUESTION Or BUTTON_YESNO Or DEFBUTTON1 )
		If iResult = IDNO Then
		    msgbox "Understand ACL name is wrong, exiting"
		    Exit Sub
		End If
	end if
	strOut = ""
	if not fso.FileExists(strInFile) Then
		msgbox "Input file " & strInFile & " not found, exiting"
		exit sub
	end if

	strOutPath = left (strInFile, InStrRev (strInFile,"\"))
	if not fso.FolderExists(strOutPath) then
		CreatePath (strOutPath)
		strOut = strOut & vbcrlf & """" & strOutPath & """ did not exists so I created it" & vbcrlf
	end if
	if strOut <> "" then
		msgbox strOut
	end if

	if right(strOutPath,1)<>"\" then
		strOutPath = strOutPath & "\"
	end if


	VerifyCmd = "show run " & strIPVer & " access-list " & strACLName

	strOutFile = strOutPath & strACLName & "-List.csv"

	strFolder = strOutPath & strACLName & "\"

	if not fso.FolderExists(strFolder) then
		CreatePath (strFolder)
		strOut = strOut & """" & strFolder & """ did not exists so I created it" & vbcrlf
	end if
	if strOut <> "" then
		msgbox strOut
	end if

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	objFileOut.writeline "primaryIPAddress,hostName,CompareTest"
	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		IPAddr = strParts(1)
		bRange=False
		iLastLine=0

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
			'crt.Screen.WaitForString "#",Timeout
			crt.Screen.Send(VerifyCmd & vbcr)
			crt.Screen.WaitForString vbcrlf,Timeout
			strResult=trim(crt.Screen.Readstring (vbcrlf&"RP/",Timeout))
			crt.Session.Disconnect
			strResultParts = split (strResult,vbcrlf)
			strTest = ""

			if not isarray(strBaseLine) then
				strBaseLine = strResultParts
				bBase = True
			else
				bBase = False
			end if
			if ubound(strBaseLine) = ubound(strResultParts) then
				bCont = True
				iLineCount = ubound(strBaseLine)
			else
				if strResultParts(1) = "% No such configuration item(s)" or ubound(strResultParts) < 2 Then
					strTest = strIPVer & " ACL " & strACLName & " doesn't exists"
					bCont = False
				else
					if bCompareAll = True then
						bCont = True
					else
						bCont = False
					end if
					strTest = "ACL length does not match, " & ubound(strResultParts) & " lines. "
					if ubound(strBaseLine) > ubound(strResultParts) then
						iLineCount = ubound(strResultParts)
					else
						iLineCount = ubound(strBaseLine)
					end if
				end if
			end if
			if bBase Then
				if strBaseLine(1) = "% No such configuration item(s)" or ubound(strBaseLine) < 2 Then
					strTest = strIPVer & " ACL " & strACLName & " doesn't exists"
					strBaseLine = empty
				else
					strTest = "Baseline ACL " & ubound(strBaseLine) & " lines."
				end if
				bCont=False
			end if
			if bCont = True then
				strTemp = ""
				iCompare = iStartCompare
				for x=iCompare to iLineCount
					y=x
					if strBaseLine(x) <> strResultParts(y) then
						if iLastLine>0 and iLastLine+1 = x-1 then
							bRange = True
						else
							if bRange = True Then
								strTemp = strTemp & "-" & iLastLine & " " & x-1
							else
								strTemp = strTemp & " " & x-1
							end if
							bRange=False
						end if
						iLastLine = x-1
					end if
				next
				if bRange = True Then
					strTemp = strTemp & "-" & iLastLine
				end if
				if strTemp = "" then
					strTest = strTest & "Pass"
				else
					strTest = strTest & "ACL line " & trim(strTemp) & " does not match. "
				end if
			end if
			set objDevName = fso.OpenTextFile(strFolder & host & ".txt", ForWriting, True)
			objDevName.writeline strResult
			objDevName.close

			objFileOut.writeline IPAddr & "," & host & "," & strTest
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.write IPAddr & "," & host & ",Not Connected,Error " & nError & ": " & strErr
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
