option Explicit
const iTimeout = 5
const host = "atrou051"
' const host = "serou041"
' const host = "oclba101"

const strCommand = "list ltm virtual destination ip-protocol pool snat snatpool profiles"
const OutHeader = "Virtual,Destination,Port,IP Protocol,Pool,SNAT,Profile"

const strOutPath = "C:\Users\sbjarna\Documents\IP Projects\ESME\F5 Forklift\"
const strSuffix  = "VIPDest"
const PagePrompt = "---(less"
const EndPrompt  = "(END)"
const SysPrompt  = "#"

Sub Main
  const ForReading    = 1
  const ForWriting    = 2
  const ForAppending  = 8
  crt.screen.synchronous = true
  crt.screen.IgnoreEscape = True

  Dim cmd, result, strVirtual, strDest, strProt, strSNAT, strOutFile, fso, objFileOut, strOut, strPool, iLoc, strProfile

  Set fso = CreateObject("Scripting.FileSystemObject")
  if not fso.FolderExists(strOutPath) then
    CreatePath (strOutPath)
    strOut = strOut & vbcrlf & """" & strOutPath & """ did not exists so I created it" & vbcrlf
  end if
  if strOut <> "" then
    msgbox strOut
  end if

  if right(strOutPath,1) <> "\" then
    strOutPath = strOutPath & "\"
  end if

  strOutFile = strOutPath & host & strSuffix &".csv"
  set objFileOut = fso.OpenTextFile(strOutFile, ForWriting, True)

  'Write a header for output file
  objFileOut.writeline OutHeader

  If crt.Session.Connected Then
    crt.Session.Disconnect
  end if

  cmd = "/SSH2 "  & host
  crt.Session.Connect cmd

  crt.Screen.Synchronous = True
  crt.Screen.WaitForString "#",iTimeout
  crt.Screen.Send(strCommand & vbCR )
  crt.Screen.WaitForString vbcrlf,iTimeout
  result = crt.Screen.WaitForStrings ("(y/n)","ltm virtual ",iTimeout)
  select case result
    case 0
      msgbox "Timeout waiting for y/n "
      exit sub
    case 1
      crt.screen.Send("y")
    case 2
      strVirtual = crt.screen.Readstring(" ",iTimeout)
      strVirtual = replace(strVirtual,",","")
      strVirtual = replace(strVirtual,":",",")
      strVirtual = trim(strVirtual)
      result = WaitWithPrompt(" destination ",vbCR)
      if result = "!@#EXIT$%^" then exit sub
      if result = "!@#Timeout$%^" then 
        msgbox "Timeout while waiting for Availability"
        exit sub
      end if
      strDest = result
      result = WaitWithPrompt(" ip-protocol ",vbCR)
      if result = "!@#EXIT$%^" then exit sub
      if result = "!@#Timeout$%^" then 
        msgbox "Timeout while waiting for State"
        exit sub 
      end if
      strProt = result
      result = WaitWithPrompt(" pool ",vbCR)
      if result = "!@#EXIT$%^" then exit sub
      if result = "!@#Timeout$%^" then 
        msgbox "Timeout while waiting for Total Connections"
        exit sub 
      end if
      strPool = result
      result = WaitWithPrompt(" profiles {","{")
      if result = "!@#EXIT$%^" then exit sub
      if result = "!@#Timeout$%^" then 
        msgbox "Timeout while waiting for profiles"
        exit sub 
      end if
      strProfile = result
      result = WaitWithPrompt(" snat",vbCR)
      if result = "!@#EXIT$%^" then exit sub
      if result = "!@#Timeout$%^" then 
        msgbox "Timeout while waiting for snat"
        exit sub 
      end if
      strSNAT = result      
      objFileOut.writeline strVirtual & "," & strDest & "," & strProt & "," & strPool & "," & strSNAT & "," & strProfile
    case else
      msgbox "unexpected case"
  end select      
  do While true
    result = WaitWithPrompt("ltm virtual "," ")
    if result = "!@#EXIT$%^" then exit do
    if result = "!@#Timeout$%^" then 
      msgbox "Timeout while waiting for Virtual"
      exit do 
    end if
    strVirtual = result

    result = WaitWithPrompt(" destination ",vbCR)
    if result = "!@#EXIT$%^" then exit do
    if result = "!@#Timeout$%^" then 
      msgbox "Timeout while waiting for Availability"
      exit do 
    end if
    strDest = result

    result = WaitWithPrompt(" ip-protocol ",vbCR)
    if result = "!@#EXIT$%^" then exit do
    if result = "!@#Timeout$%^" then 
      msgbox "Timeout while waiting for State"
      exit do 
    end if
    strProt = result

    result = WaitWithPrompt(" pool ",vbCR)
    if result = "!@#EXIT$%^" then exit do
    if result = "!@#Timeout$%^" then 
      msgbox "Timeout while waiting for Total Connections"
      exit do 
    end if
    strPool = result
      
    result = WaitWithPrompt(" profiles {","{")
    if result = "!@#EXIT$%^" then exit sub
    if result = "!@#Timeout$%^" then 
      msgbox "Timeout while waiting for profiles"
      exit sub 
    end if
    strProfile = result

    result = WaitWithPrompt("snat",vbCR)
    if result = "!@#EXIT$%^" then exit do
    if result = "!@#Timeout$%^" then 
      msgbox "Timeout while waiting for snat"
      exit do 
    end if
    strSNAT = result

    objFileOut.writeline strVirtual & "," & strDest & "," & strProt & "," & strPool & "," & strSNAT & "," & strProfile
  loop
  crt.Screen.Synchronous = False
  crt.session.disconnect
  objFileOut.close
  Set objFileOut = Nothing

  Set fso = Nothing

  msgbox "All Done, Cleanup complete"

End Sub


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

Function WaitWithPrompt (strWaitFor,strReadUntil)
'-------------------------------------------------------------------------------------------------'
' Function WaitWithPrompt (strWaitFor,strReadUntil)                                               '
'                                                                                                 '
' This function takes a phrase to look for (strWaitFor)                                           '
' and returns everthing that follows until strReadUntil                                           '
'-------------------------------------------------------------------------------------------------'
  dim strResult, iPrompt

    iPrompt=crt.Screen.WaitForStrings (strWaitFor,PagePrompt,EndPrompt,SysPrompt,iTimeout)
    do while iPrompt = 2
      crt.Screen.Send(" ")
      iPrompt=crt.Screen.WaitForStrings(strWaitFor,PagePrompt,EndPrompt,SysPrompt,iTimeout)
    loop
    select case iPrompt
      case 3,4
        WaitWithPrompt = "!@#EXIT$%^"
      case 0
        WaitWithPrompt = "!@#Timeout$%^"
      case 1
        strResult = crt.screen.Readstring(strReadUntil,iTimeout)
        strResult = replace(strResult,",","")
        strResult = replace(strResult,vbcrlf,"")
        strResult = replace(strResult,":",",")
        strResult = trim(strResult)
        WaitWithPrompt = strResult
      case else
        msgbox "unexpected case"
    end select
End Function