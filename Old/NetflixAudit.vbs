#$language = "VBScript"
#$interface = "1.0"

'|----------------------------------------------------------------------------------------------------------|
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 06/23/16                                                                                      |
'|  Copyright: Siggi Bjarnason 2016                                                                         |
'|----------------------------------------------------------------------------------------------------------|

Option Explicit
dim strInFile, strOutFile, strDelOut

' User Spefified values, specify values here per your needs
strInFile    = "C:\Users\sbjarna\Documents\IP Projects\Automation\Netflix\AGC_cleanup_nexthop.csv"
strOutFile   = "C:\Users\sbjarna\Documents\IP Projects\Automation\Netflix\Audit.csv"
strDelOut    = "C:\Users\sbjarna\Documents\IP Projects\Automation\Netflix\DelOut.csv"

const DelNum       = 38
const DelCount     = 4
const ValidateDel  = false

'Nothing below here is user configurable proceed at your own risk.

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 2
	const VerifyCmd = "show run ipv4 access-list OMW-ABF-IN | include "
	const FirstCol = 4

	dim strParts, strLine, objFileIn, objFileOut, objDelOut, host, ConCmd, cmd, fso, nError, strErr, strResult, x,y
	dim strResultParts, strNextHop, objDelDICT, strDelList(), strLineNum, strOut, strResLines

	LoadDelList strDelList

	crt.screen.synchronous = true
	crt.screen.IgnoreEscape = True

	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")

	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn  = fso.OpenTextFile(strInFile, ForReading, false)
	set objDelDICT = CreateObject("Scripting.Dictionary")

	if ValidateDel = True then
		set objDelOut  = fso.OpenTextFile(strDelOut, ForWriting, True)
	end if 

	objFileOut.writeline "S=Status. (A):Adding a line there; (D):Deleting this line; (X):Adding a line after Deleting it" & vbcrlf
	objFileOut.writeline "Device,Line,S,Result,Next Hop"

	strLine = objFileIn.readline
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(1)
		objDelDICT.RemoveAll

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
			for x = FirstCol to ubound(strParts)-1
				cmd = VerifyCmd & strParts(x) & vbcrlf
				crt.Screen.Send(cmd)
				crt.Screen.WaitForString "[K",Timeout
				strResult=trim(crt.Screen.Readstring ("RP/0/",vbcrlf,Timeout))
				strResultParts = split(strResult," ")
				if ubound(strResultParts) > 5 then
					strNextHop = strResultParts(ubound(strResultParts))
				else
					strNextHop = ""
				end if
				if x<FirstCol+DelCount then
					strResult = "(D)," & strResult
					if not objDelDICT.Exists (strParts(x)) then
						objDelDICT.add strParts(x),x
					end if
				else
					if objDelDICT.Exists (strParts(x)) then
						strResult = "(X)," & strResult
					else
						strResult = "(A)," & strResult
					end if
				end if
				objFileOut.writeline host & "," & strParts(x) & "," & strResult  & "," & strNextHop
				crt.Screen.WaitForString "#", Timeout
			next
			if ValidateDel = True then
				strOut = host
				for x = 0 to DelNum-1
					cmd = VerifyCmd & strDelList(x) & vbcrlf
					crt.Screen.Send(cmd)
					crt.Screen.WaitForString "[K",Timeout
					strResult=trim(crt.Screen.Readstring ("RP/0/",Timeout))
					if instr(strResult,vbcrlf)>0 then
						strResLines=split(strResult,vbcrlf)
						for y = 0 to ubound(strResLines)
							strResultParts = split(trim(strResLines(y))," ")
							if ubound(strResultParts) > 5 then
								strOut = strOut & ", " & strResultParts(0)
							end if
						next
					end if 
					crt.Screen.WaitForString "#", Timeout
				next
				objDelOut.writeline strOut
			end if 
			crt.Session.Disconnect
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
	if ValidateDel = True then
		objDelOut.close
		set objDelOut  = Nothing
	end if 	
	Set fso = Nothing

	msgbox "All Done, Cleanup complete"

end sub

sub LoadDelList (ByRef strDelList())
redim strDelList(DelNum)
	strDelList(0)  = "172.56.128.64 0.0.0.63"
	strDelList(1)  = "172.56.129.64 0.0.0.63"
	strDelList(2)  = "172.56.130.64 0.0.0.63"
	strDelList(3)  = "172.56.131.64 0.0.0.63"
	strDelList(4)  = "172.56.132.64 0.0.0.63"
	strDelList(5)  = "172.56.133.64 0.0.0.63"
	strDelList(6)  = "172.56.134.64 0.0.0.63"
	strDelList(7)  = "172.56.136.64 0.0.0.63"
	strDelList(8)  = "172.56.138.64 0.0.0.63"
	strDelList(9)  = "172.56.139.64 0.0.0.63"
	strDelList(10) = "172.56.141.64 0.0.0.63"
	strDelList(11) = "172.56.142.64 0.0.0.63"
	strDelList(12) = "172.56.143.0 0.0.0.31"
	strDelList(13) = "172.56.143.64 0.0.0.63"
	strDelList(14) = "172.56.144.64 0.0.0.63"
	strDelList(15) = "172.56.145.64 0.0.0.63"
	strDelList(16) = "172.56.146.64 0.0.0.63"
	strDelList(17) = "208.54.32.0 0.0.0.31"
	strDelList(18) = "208.54.34.0 0.0.0.31"
	strDelList(19) = "208.54.35.0 0.0.0.63"
	strDelList(20) = "208.54.36.0 0.0.0.63"
	strDelList(21) = "208.54.37.0 0.0.0.31"
	strDelList(22) = "208.54.38.0 0.0.0.31"
	strDelList(23) = "208.54.39.0 0.0.0.31"
	strDelList(24) = "208.54.40.0 0.0.0.31"
	strDelList(25) = "208.54.44.0 0.0.0.63"
	strDelList(26) = "208.54.45.0 0.0.0.31"
	strDelList(27) = "208.54.64.0 0.0.0.31"
	strDelList(28) = "208.54.66.0 0.0.0.63"
	strDelList(29) = "208.54.67.0 0.0.0.31"
	strDelList(30) = "208.54.70.0 0.0.0.31"
	strDelList(31) = "208.54.74.128 0.0.0.63"
	strDelList(32) = "208.54.80.0 0.0.0.31"
	strDelList(33) = "208.54.83.0 0.0.0.31"
	strDelList(34) = "208.54.85.0 0.0.0.31"
	strDelList(35) = "208.54.87.0 0.0.0.31"
	strDelList(36) = "208.54.5.0 0.0.0.31"
	strDelList(37) = "172.56.140.64 0.0.0.63"
end sub