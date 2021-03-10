#$language = "VBScript"
#$interface = "1.0"
Option Explicit
public TestMode
TestMode = False

Sub Main
	const ForReading    = 1
	const ForWriting    = 2
	const ForAppending  = 8
	Const Timeout = 15
	
	dim strInFile, strOutFile, objFileIn, fso, nError, strErr, strVlanName, strMac, result, strMasq, strMOPFolder, strMOPPath
	
	'|----------------------------------------------------------------------------------------------------------|
	'|  Author: Siggi Bjarnason                                                                                 |
	'|  Authored: 02/29/16                                                                                      |
	'|  Copyright: Siggi Bjarnason 2016                                                                         |
	'|----------------------------------------------------------------------------------------------------------|
	
	' User Spefified values, specify values here per your needs
	
	strInFile    = "C:\Users\sbjarna\Documents\IP Projects\F5 MAC Masquerade\F5MACMasqueradePhase3.txt"
	strOutFile   = "C:\Users\sbjarna\Documents\IP Projects\F5 MAC Masquerade\F5Info3.txt"
	strMOPFolder = "C:\Users\sbjarna\Documents\IP Projects\F5 MAC Masquerade\Phase3MOPs\"
	
	'Nothing below here is user configurable proceed at your own risk.
	
	dim strLogFileName, strScriptFullName, strParts, strScriptName, dictHostNames, strLine, strAddr, objFileOut, continue
	dim strVersion, screenrow, host, ConCmd, cmd, strVerify,objMOPOut, objBackout, strBackoutPath, strVerifyBlock, strVerifyBackoutBlock
	dim host2, Router1, Router2, strImpl, strBackout, StrBackVerify
	
	crt.screen.synchronous = true
	
	' Creating a File System Object to interact with the File System
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)
	Set objFileIn = fso.OpenTextFile(strInFile, ForReading, false)
	
	While not objFileIn.atendofstream
		strLine = objFileIn.readline
		strParts = split(strLine,",")
		host = strParts(0)
		host2 = strParts(1)
		strVersion = "unknown"
		strMOPPath = strMOPFolder & host & "-" & host2 & "_implementation.txt"
		strBackoutPath = strMOPFolder & host & "-" & host2 & "_backout.txt"
		strVerify = ""
		strImpl = ""
		strBackout = ""
		StrBackVerify = ""
		
		If crt.Session.Connected Then
			crt.Session.Disconnect
		end if
		
		cmd = "/SSH2 "  & host
		on error resume next
		crt.Session.Connect cmd
		on error goto 0
		
		If crt.Session.Connected Then
			crt.Screen.Synchronous = True
			continue = true
			crt.Screen.WaitForString "#",Timeout
			nError = Err.Number
			strErr = Err.Description
			If nError <> 0 Then
				strVersion = "Error " & nError & ": " & strErr
			end if
			crt.Screen.Send("show /sys version" & vbCR )
			crt.Screen.WaitForString "Build",Timeout
			screenrow = crt.screen.CurrentRow - 1
			strVersion = trim(crt.Screen.Get(screenrow, 12, screenrow, 20 ))
			result=crt.Screen.WaitForStrings ("#","---(less","(END)")
			if result = 0 then
				continue = false
				objFileOut.writeline host & ", timeout"
			else
				if result = 2 or result =3 then
					crt.Screen.Send("q")
					crt.Screen.WaitForString "#",Timeout
				end if
				set objMOPOut = fso.OpenTextFile(strMOPPath, ForWriting, True)
				set objBackout = fso.OpenTextFile(strBackoutPath, ForWriting, True)
				objMOPOut.writeline "#" & vbcrlf & "##########" & host & " IMPLEMENTATION##########" & vbcrlf & "#" & vbcrlf & "#" & vbcrlf
				objBackout.writeline "#" & vbcrlf & "##########" & host & " BACKOUT##########" & vbcrlf & "#" & vbcrlf & "#" & vbcrlf
				crt.Screen.Send("show net vlan" & vbCR )
			end if
		
			do While continue
				result=crt.Screen.WaitForStrings ("Net::Vlan:","---(less","(END)", Timeout)
				do while result = 2
					crt.Screen.Send(" ")
					result=crt.Screen.WaitForStrings("Net::Vlan:","---(less","(END)", Timeout)
				loop
				if result = 3 then
					exit do
				end if
				if result = 0 then
					objFileOut.writeline host & ", timeout waiting for next vlan"
					exit do
				end if
				strVlanName = trim(crt.screen.Readstring(vbCR,Timeout))
		
				result=crt.Screen.WaitForStrings ("Mac Address ","---(less","(END)", Timeout)
				if result = 2 then
					crt.Screen.Send(" ")
					crt.Screen.WaitForString "Mac Address ",Timeout
				end if
				if result = 3 then
					exit do
				end if
				if result = 0 then
					objFileOut.writeline host & ", timeout waiting for MAC"
					exit do
				end if
				strMac = trim(crt.screen.Readstring(vbCR, Timeout))
				strParts = split (strMac, ":")
				if strParts(0) = 0 then
					strParts(0) = 2
					strMasq = join (strParts,":")
				else
					strMasq = "Undetermined"
				end if
				objFileOut.writeline host & "," & strVersion & "," & strVlanName & "," & strMac & "," & strMasq
				strImpl = strImpl & "b vlan " & strVlanName & " mac masq " & strMasq & vbcrlf
				strBackout = strBackout & "b vlan " & strVlanName & " mac masq none" & vbcrlf
			loop
			objMOPOut.writeline strImpl
			objBackout.writeline strBackout
			strVerifyBlock = vbcrlf & vbcrlf & _
							"#####Verification for " & host & " #####" & vbcrlf & vbcrlf & _
							"b vlan show | grep tag" & vbcrlf & _
							"! verify MAC address matches implementation output should looks something like this" & vbcrlf & vbcrlf & _
							"VLAN <VLAN_NAME>   tag xx  <02:Masquerade MAC> MTU 1500" & vbcrlf & _
							"VLAN <VLAN_NAME>   tag xx  <02:Masquerade MAC> MTU 1500" & vbcrlf & _
							"VLAN <VLAN_NAME>   tag 155  <02:Masquerade MAC> MTU 1500" & vbcrlf & _
							"VLAN <VLAN_NAME>   tag 158  <02:Masquerade MAC> MTU 1500" & vbcrlf & _
							"VLAN <VLAN_NAME>   tag 159  <02:Masquerade MAC> MTU 1500" & vbcrlf & vbcrlf
			objMOPOut.writeline strVerifyBlock
			objBackout.writeline strVerifyBlock
			
			' Writing out MOP for partner device
			objMOPOut.writeline "#" & vbcrlf & "##########" & host2 & " IMPLEMENTATION##########" & vbcrlf & "#" & vbcrlf & "#" & vbcrlf
			objBackout.writeline "#" & vbcrlf & "##########" & host2 & " BACKOUT##########" & vbcrlf & "#" & vbcrlf & "#" & vbcrlf
			objMOPOut.writeline strImpl
			objBackout.writeline strBackout
			strVerifyBlock = replace (strVerifyBlock,host,host2)
			objMOPOut.writeline strVerifyBlock
			strVerifyBackoutBlock = replace (strVerifyBlock,host,host2)
			objBackout.writeline strVerifyBackoutBlock
			
			objMOPOut.close
			objBackout.close
			set objMOPOut = nothing
			set objBackout = nothing
			crt.Session.Disconnect
		else
			nError = crt.GetLastError
			strErr = crt.GetLastErrorMessage
			objFileOut.writeline host & ",Not Connected,Error " & nError & ": " & strErr
		end if
	wend
	
	Set objFileIn  = Nothing
	Set objFileOut = Nothing
	Set fso = Nothing
	set dictHostNames = Nothing
	
	msgbox "All Done, Cleanup complete"
end sub

