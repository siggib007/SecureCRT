option Explicit
Dim cmd, strOutFile, fso, objFileOut, strOut, strResultParts, strResult, x, strLineWords, objPartition, strPartition

const iTimeout = 5
const host = "LBCTTN13"
' const host = "serou041"
' const host = "oclba101"
' const host = "Lab"

const strOutPath = "C:\Users\sbjarna\Documents\IP Projects\Automation\A10Audit\"
const PagePrompt = "--MORE--"
const SysPrompt  = "#"
const strSuffix = "_Audit"
Const Timeout = 2

Sub Main
  const ForReading    = 1
  const ForWriting    = 2
  const ForAppending  = 8
  crt.screen.synchronous = true
  crt.screen.IgnoreEscape = True


  Set fso = CreateObject("Scripting.FileSystemObject")
  set objPartition = CreateObject("Scripting.Dictionary")
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
  objFileOut.writeline "Partition,Object Type,Object Name,IP Address 1,IP Address 2, Mask"

  If crt.Session.Connected Then
    crt.Session.Disconnect
  end if

  if host = "Lab" then 
    cmd = "/s ""SoftAx"""
  else
    cmd = "/SSH2 /ACCEPTHOSTKEYS "  & host
  end if

  crt.Session.Connect cmd

  crt.Screen.Synchronous = True
  objPartition.RemoveAll
  
  crt.Screen.WaitForString( SysPrompt )
  crt.Screen.Send("term len 0" & vbcr)
  crt.Screen.WaitForString SysPrompt,Timeout

  strPartition = "Shared"
  ParseConfig

  for each strPartition in objPartition
    crt.Screen.Send("active-partition " & strPartition & vbCR )
    crt.Screen.WaitForString SysPrompt,Timeout
    ParseConfig    
  next

  crt.Screen.Synchronous = False
  crt.session.disconnect
  objFileOut.close
  Set objFileOut = Nothing

  Set fso = Nothing

  msgbox "All Done, Cleanup complete"

End Sub

Function ParseConfig

  crt.Screen.Send("show run with-default " & vbCR )
  strResult=trim(crt.Screen.Readstring (SysPrompt,Timeout))

  strResultParts = split (strResult,vbcrlf)
  for x=0 to ubound(strResultParts)
    strLineWords = split (strResultParts(x)," ")
    if ubound(strLineWords)>0 then 
      select case strLineWords(0)
        case "partition"
          objPartition.add strLineWords(1),""
        case "ip"
          if strLineWords(2) = "pool" then 
            objFileOut.writeline strPartition & "," & "NAT," & strLineWords(3) & "," & strLineWords(4) & "," & strLineWords(5) & "," & strLineWords(7)
          end if
        case "slb"
          if strLineWords(1) = "server" then 
            objFileOut.writeline strPartition & "," & "Server," & strLineWords(2) & "," & strLineWords(3) 
          end if
          if strLineWords(1) = "virtual-server" then 
            objFileOut.writeline strPartition & "," & "VIP," & strLineWords(2) & "," & strLineWords(3) 
          end if
      end select 
    end if 
  next

End Function

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
