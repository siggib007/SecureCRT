Option Explicit

'|----------------------------------------------------------------------------------------------------------|
'|  This script will log into each ARG in a specified spreadsheet and confirm ACLs are up to standards.     |
'|  If deviations are found, will generate configuration files for HPNA or manual MOP.                      |
'|                                                                                                          |
'|  Author: Siggi Bjarnason                                                                                 |
'|  Authored: 04/10/2017                                                                                    |
'|  Copyright: Siggi Bjarnason 2017                                                                         |
'|----------------------------------------------------------------------------------------------------------|

 ' User Spefified values, specify values here per your needs

  Const AutoCloseInput = True
  Const ExcelVisible   = True
  const Timeout        = 250 ' Timeout in seconds for each command, if expected results aren't received withing this time, the script moves on.
  const MaxError       = 5 ' If connection error occurs how often to retry

 ' Nothing below here is user values, proceed with caution and at your own risk.
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

  Dim dictVars, dictACLs, dictChange, dictDevAffected, dictFailed
  Dim app, objShell, fso, objFileOut, objLogOut, objACLGen, objChangeOut, wsNames, wsVars, wsACL, wbin
  Dim strOutPath, strOutFile, strLogFile, strAsIsOutPath, strMOPPath, strWBin, strACLNameVar
  Dim iNameCol, iACLCol, iVarCol, iStartPos, iStopPos, iHostCol, iIPCol, iACLNameCol

Sub main
  Dim bComp, bRange, bNewChange, bOut, dkey, dkeys, errcode, errmsg, dItems
  Dim iNameRow, iVarRow, iACLRow, iGSeq, iASeq, iLastLine, iError, iResult, iChangeID, iFailed, iTemp
  Dim strACLVar, strACLVar2, strTempOut, strACLName, strACLID
  Dim strHostname, strIPAddr, strResultParts, strGenOutPath
  Dim strNotes, strNoMatch, strMissing, strErr, strIPVer, strChange

  ' comparison constants
  Const vbBinaryCompare = 0 'Perform a binary comparison, i.e case sensitive
  Const vbTextCompare = 1 'Perform a textual comparison, i.e. case insensitive

  ' Display File Open dialog so that the user can selelct the input workbook.
  strWBin = crt.Dialog.FileOpenDialog("Please select OMW ACL Standard Spreadsheet", "Open", "", "Excel Files (*.xlsx)|*.xlsx||")
  if strWBin = "" then
    msgbox "No file provided, exiting"
    exit Sub
  end if
  ' Hook to the filesystem
  Set fso = CreateObject("Scripting.FileSystemObject")

  ' Doublecheck the input workbook actually does exists
  if not fso.FileExists(strWBin) Then
    msgbox "Input file " & strWBin & " not found, exiting"
    set fso = nothing
    exit sub
  end if

  ' Parse out the path only from the workbook file name, make sure the path ends in a \
  strOutPath = left (strWBin, InStrRev (strWBin,"\"))
  if right(strOutPath,1)<>"\" then
    strOutPath = strOutPath & "\"
  end if

  ' Create file names for log file and results file based on the input file
  strOutFile = left (strWBin, InStrRev (strWBin,".")-1)&"-Results.csv"
  strlogFile = left (strWBin, InStrRev (strWBin,".")-1)&"-log.txt"

  ' Creating some dictionaries
  Set dictVars        = CreateObject("Scripting.Dictionary")
  Set dictACLs        = CreateObject("Scripting.Dictionary")
  Set dictChange      = CreateObject("Scripting.Dictionary")
  Set dictDevAffected = CreateObject("Scripting.Dictionary")
  Set dictFailed      = CreateObject("Scripting.Dictionary")

  ' Get a direct hook into specific sheets in the workbook, as well a command line hook.
  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wbin = app.Workbooks.Open (strWBin,0,true)
  Set wsNames = wbin.Worksheets("ACL Names")
  Set wsVars = wbin.Worksheets("OMW-Vars")
  Set wsACL = wbin.Worksheets("ACL Lines")

  ' Create output files and catch any errros in the process
  set objLogOut = CreateFile(strlogFile)
  set objFileOut = CreateFile(strOutFile)

  ' Now start the real work
  objLogOut.writeline "Starting at " & now()
  objFileOut.writeline "primaryIPAddress,hostName,ACL Name, comment"
  app.visible = ExcelVisible

  ' Grab the column headers of the variable sheet and stick them into a dictionary.
  iVarCol=1
  dictVars.removeall
  Do
    If wsVars.Cells(1,iVarCol).Value = "" Then Exit Do
    If not dictVars.Exists(wsVars.Cells(1,iVarCol).value) then
      dictVars.Add wsVars.Cells(1,iVarCol).value, iVarCol
    End If
    iVarCol = iVarCol + 1
  loop

  ' Get the column headers for the ACL sheet and sticke them into a dicationary.
  iACLCol=1
  dictACLs.removeall
  Do
    If wsACL.Cells(1,iACLCol).Value = "" Then Exit Do
    If not dictACLs.Exists(wsACL.Cells(1,iACLCol).value) then
      dictACLs.Add wsACL.Cells(1,iACLCol).value, iACLCol
    End If
    iACLCol = iACLCol + 1
  loop

  strChange  = ""
  strMOPPath = left (strWBin, InStrRev (strWBin,".")-1)& "-Changes\"
  if not fso.FolderExists(strMOPPath) then
    CreatePath (strMOPPath)
    objLogOut.writeline """" & strMOPPath & """ did not exists so I created it"
  end if

  ' Get the column number for the Router IP address in the Variable sheet, report an error and exit if can't be deteremined.
  if dictVars.Exists("primaryIPAddress") then
    iIPCol = dictVars("primaryIPAddress")
  else
    objLogOut.writeline "couldn't find primaryIPAddress in dictVars :-("
    msgbox "couldn't find primaryIPAddress in dictACLs, exiting :-("
    CleanUp
    exit sub
  end if

  ' Get the column number for the Router hostname in the Variable sheet, report an error and exit if can't be deteremined.
  if dictVars.Exists("hostName") then
    iHostCol = dictVars("hostName")
  else
    objLogOut.writeline "couldn't find hostName in dictVars :-("
    msgbox "couldn't find hostName in dictACLs, exiting :-("
    CleanUp
    exit sub
  end if

  iVarRow=2
  iError = 1
  do ' Now start looping throught the variable sheet.
    iFailed = 0
    strIPAddr = wsVars.Cells(iVarRow,iIPCol).value
    if strHostname <> wsVars.Cells(iVarRow,iHostCol).value then
      strHostname = wsVars.Cells(iVarRow,iHostCol).value
      iError = 1 ' Ensure the error counter is set back to 1, which is default and means no error.
    end if
    if wsVars.Cells(iVarRow,iHostCol).value = "" then
      objLogOut.writeline "reached end of variable rows"
      exit do
    end if


    iNameRow = 4
    do
      strACLID = wsNames.Cells(iNameRow,1).value
      strACLName = wsNames.Cells(iNameRow,2).value
      strACLNameVar = wsNames.Cells(iNameRow,3).value
      strIPVer = wsNames.Cells(iNameRow,4).value
      dictDevAffected.RemoveAll
      dictFailed.RemoveAll

      ' objLogOut.writeline "Working on ACL " & strACLID & " / " & strACLName & " / " & strACLNameVar & " / " & strIPVer
      ' Get the column number for the ACL Name in the Variable sheet, report an error and exit if can't be deteremined.
      if dictVars.Exists(strACLNameVar) then
        iACLNameCol = dictVars(strACLNameVar)
      else
        if strACLNameVar <> "" then
          objLogOut.writeline "couldn't find ACL Name, " & strACLNameVar & ", in dictVars :-("
          msgbox "couldn't find ACL name, " & strACLNameVar & ", in dictACLs, exiting :-("
          CleanUp
          exit sub
        end if
      end if

      ' Setup the output paths to be ACL specific
      ' First Folder for the Generated ACL's
      strGenOutPath = strOutPath & strACLName & "-Gen\"
      if not fso.FolderExists(strGenOutPath) then
        CreatePath (strGenOutPath)
        objLogOut.writeline """" & strGenOutPath & """ did not exists so I created it"
      end if

      ' Then a folder for the ACL's we grab from the router
      strAsIsOutPath = strOutPath & strACLName & "-AsIs\"
      if not fso.FolderExists(strAsIsOutPath) then
        CreatePath (strAsIsOutPath)
        objLogOut.writeline """" & strAsIsOutPath & """ did not exists so I created it"
      end if

      ' There should be column in the ACL sheet whose header is the same as the ACLID we're working on. Get that column number or report an issue and exit.
      if dictACLs.Exists(strACLID) then
        iACLCol = dictACLs(strACLID)
      else
        objLogOut.writeline "couldn't find " & strACLID & " in dictACLs :-("
        msgbox "couldn't find " & strACLID & " in dictACLs, exiting :-("
        CleanUp
        exit sub
      end if

      ' If this ACL has variability in the name find the column number in the variable sheet.
      if strACLNameVar <> "" then
        if dictVars.Exists(strACLNameVar) then
          strACLName = wsVars.Cells(iVarRow,dictVars(strACLNameVar)).value
        end if
        strChange  = strIPVer & " access-list $" & strACLNameVar & "$" &vbcrlf
      else
        strChange  = strIPVer & " access-list " & strACLName & vbcrlf
      end if
      objLogOut.writeline "Starting on router " & strHostname & " with ACL " & strACLName ' Log that we are about to log into a router.

      strNotes   = ""
      strNoMatch = ""
      strMissing = ""

      iGSeq = 0
      iASeq = 0
      bOut = False

      strResultParts = GetAsIsACL(strIPVer, strACLName, strHostname, iError)
      if isArray(strResultParts) then
        iError = 1
        if dictFailed.Exists(strHostname) then dictFailed.remove(strHostname)
        ' objLogOut.writeline "connected to " & strHostname
        ' objLogOut.writeline "Size of the result array: " & ubound(strResultParts)
        if ubound(strResultParts) < 2 then
          redim strResultParts(5)
          strResultParts(1) = "% No response to the command"
        end if
      else
        strNotes = "Failed to connect "
        objLogOut.writeline "Failed to connect to " & strHostname
        if IsNumeric(strResultParts) then
          iError = strResultParts
          if not dictFailed.Exists(strHostname) then dictFailed.add strHostname,iVarRow
        else
          objLogOut.writeline "GetAsIsACL returned neither array nor a number, this shouldn't happen so I'm quitting. I got back: " & strResultParts
          exit Sub
        end if
      end if
      strTempOut = ""
      iACLRow=2
      iResult=1
      if iError = 1 then ' If there has been no connection errors, analyse the results.
        if left(trim(strResultParts(iResult)),1)="%" then
         objLogOut.writeline "encountered error: " & strResultParts(iResult) & " moving on to next ACL"
         exit do
        end if
        set objACLGen = CreateFile(strGenOutPath & strHostname & "-" & strACLName & ".txt")
        objLogOut.writeline "Created " & strGenOutPath & strHostname & "-" & strACLName & ".txt"
        do
          ' bComp = False
          if wsACL.Cells(iACLRow,iACLCol).value <> "" then ' Is current ACL standard line applicable to this ACL (X in the proper column)
            iStartPos = instr (1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare) ' Look for $ which indicates a start of a variable in the ACL standard.
            if iStartPos > 0 then ' If the current line has a variable parse out the variable, and substitute it with the proper value.
              objLogOut.writeline "Looking at standard ACL line " & iACLRow & ": " & wsACL.Cells(iACLRow,1).value & ", bComp:" & bComp & " bOut:" & bOut
              iStopPos = instr (iStartPos+1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare) ' Locate the end of the variable name.
              strACLVar = mid(wsACL.Cells(iACLRow,1).value,iStartPos+1,iStopPos-iStartPos-1) ' Store the name of the variable.
              objLogOut.writeline "Parsed out ACL Variable: " & strACLVar
              iStartPos = instr (iStopPos+1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare) ' Look for $ which indicates a start of a variable in the ACL standard.
              objLogOut.writeline "Second iStartPos: " & iStartPos & " last iStopPos: " & iStopPos
              if iStartPos > 0 then ' you have second variables per line
                iStopPos = instr (iStartPos+1,wsACL.Cells(iACLRow,1).value,"$",vbTextCompare) ' Locate the end of the variable name.
                strACLVar2 = mid(wsACL.Cells(iACLRow,1).value,iStartPos+1,iStopPos-iStartPos-1) ' Store the name of the variable.
                objLogOut.writeline "Parsed out second ACL Variable: " & strACLVar2
              else
                strACLVar2 = ""
              end if
              if strACLVar = "ACLName" then ' If the variable is "ACLName" then substitute it with the actual ACL Name
                strTempOut = trim(replace(wsACL.Cells(iACLRow,1).value,"$ACLName$",strACLName))
                objACLGen.writeline strTempOut
                bOut = True
                bComp = True
                ' objLogOut.writeline "Found ACLName"
              ' else
              '   bComp = False
              end if ' End if the Variable is ACLName.

              if dictVars.Exists(strACLVar) then ' If the variable we found exists in the variable sheet.
                iVarCol = dictVars(strACLVar)
                if bOut = False then ' Only process if not processed before.
                  if wsVars.Cells(iVarRow, iVarCol) <> "" then ' If the varible value is not an empty string do the substitution and write the generate ACL line to file.
                    strTempOut = trim(replace(wsACL.Cells(iACLRow,1).value,"$" & strACLVar & "$",wsVars.Cells(iVarRow, iVarCol)))
                    if strACLVar2 <> "" then
                      if dictVars.Exists(strACLVar2) then ' If the variable we found exists in the variable sheet.
                        iVarCol = dictVars(strACLVar2)
                        if wsVars.Cells(iVarRow, iVarCol) <> "" then ' If the varible value is not an empty string do the substitution and write the generate ACL line to file.
                          strTempOut = trim(replace(wsACL.Cells(iACLRow,1).value,"$" & strACLVar2 & "$",wsVars.Cells(iVarRow, iVarCol)))
                        end if
                      end if
                    end if
                    objACLGen.writeline strTempOut
                    bOut = True
                    bComp = True
                    ' objLogOut.writeline "translated it into: " & strTempOut
                  else ' Variable had no value, skipping.
                    bComp = False
                    bOut = False
                    ' objLogOut.writeline strACLVar & " has no value for " & strHostname & " " & strACLName
                  end if ' End if Variable is not empty string.
                else
                  ' objLogOut.writeline "been here before"
                end if ' if processed before
              else
                ' objLogOut.writeline "Error: Variable " & strACLVar & " not found in dictVars"
              end if ' end if variable exists
            else ' If the current line has no variables, just write it to the file.
              if bOut = False then ' Only process if not processed before.
                strTempOut = trim(wsACL.Cells(iACLRow,1).value)
                objACLGen.writeline strTempOut
                bOut = True
                bComp = True
                ' objLogOut.writeline strTempOut & " has no variable"
              else
                ' bComp = False
                ' objLogOut.writeline "Line already processed"
              end if ' if processed before
            end if ' End if analyzing the current line of the ACL standard.
            ' objLogOut.writeline "Completed analyzing ACL Line, next compare"
            ' objLogOut.writeline "Standard ACL line " & iACLRow & ": " & wsACL.Cells(iACLRow,1).value & ", bComp:" & bComp & " bOut:" & bOut
            ' objLogOut.writeline "Router ACL Line " & iResult & ": " & strResultParts(iResult)
            ' objLogOut.writeline "Old Seq. iGSeq:" & iGSeq & " iASeq:" & iASeq
            if bComp then ' If the ACL line was found applicable, compare the generated line with the same line in the ACL capture.
              ' Grab the sequence number of the generated ACL line we're looking at
              iTemp = GetSeq(strTempOut)
              if iTemp > 0 then iGSeq = iTemp
              iTemp = GetSeq(strResultParts(iResult))
              if iTemp > 0 then iASeq = iTemp
              ' objLogOut.writeline "New Seq. iGSeq:" & iGSeq & " iASeq:" & iASeq & " iResult:" & iResult & " iACLRow:" & iACLRow
              if strTempOut <> trim(strResultParts(iResult)) Then ' If generated and AsIs lines aren't identical, note it.
                ' objLogOut.writeline strHostname & " " & strACLName & " Not identical, analyzing how"
                ' objLogOut.writeline " Gen: " & strTempOut
                ' objLogOut.writeline "AsIs: " & trim(strResultParts(iResult))
                ' objLogOut.writeline "------------End different--------------------"
                if iGSeq > iASeq Then
                  objLogOut.writeline "Line " & iResult & ": Extra line on router not in standard: " & trim(strResultParts(iResult))
                  strMissing = strMissing & iResult & "(only on router) "
                  ' bRange = False
                  strChange = strChange & "no " & iASeq & vbcrlf
                end if
                if iGSeq < iASeq Then
                  objLogOut.writeline "Line " & iResult & ": This standard line missing from router: " & strTempOut
                  strMissing = strMissing & iResult & "(missing from router) "
                  ' bRange = False
                  strChange = strChange & wsACL.Cells(iACLRow,1).value & vbcrlf
                end if
                if iGSeq = iASeq then ' If seqences match report the lines don't match
                  objLogOut.writeline strHostname & " " & strACLName & " NO matchy on line with same Seq " & iResult
                  objLogOut.writeline " Gen: " & strTempOut
                  objLogOut.writeline "AsIs: " & trim(strResultParts(iResult))
                  objLogOut.writeline "------------End NO Match--------------------"
                  strChange = strChange & wsACL.Cells(iACLRow,1).value & vbcrlf
                end if ' end if seqences match
              else
                ' objLogOut.writeline strHostname & " " & strACLName & " matching on line " & iResult
                ' objLogOut.writeline " Gen: " & strTempOut
                ' objLogOut.writeline "AsIs: " & trim(strResultParts(iResult))
                ' objLogOut.writeline "-------------End Match-------------------"
              end if ' End if generated and AsIs are different.
            end if ' End If ACL is applicable
          else ' This line isn't a part of this ACL
            bComp = False
            ' objLogOut.writeline "Not this ACL"
          end if ' Is current ACL standard line applicable to this ACL (X in the proper column).
          ' objLogOut.writeline "Next Line time: bComp:" & bComp & " bOut: " & bOut & " currently on iResult:" & iResult & " and iACLRow:" & iACLRow
          if bComp = True or bOut = True then
            ' If there are lines left in the captured ACL and router ACL sequence is lower or equal move on to the next line.
            ' objLogOut.writeline "iResult:" & iResult & " size:" & ubound(strResultParts)
            if iResult < ubound(strResultParts) then
              if iGSeq >= iASeq then iResult = iResult + 1
            else
              ' objLogOut.writeline "Reached end of AsIs ACL"
              exit do
            end if
            ' objLogOut.writeline "now iResult: " & iResult
          end if
          if iGSeq <= iASeq then
            iACLRow = iACLRow + 1 ' Move down line in the ACL sheet if we are in sync or we are too low
            bOut = False
          end if
        loop until wsACL.Cells(iACLRow,1).Value = "" ' Unless the new line is blank, loop back and repeat.
        objACLGen.Close
        ' objLogOut.writeline "iACLRow " & iACLRow & " is blank, done with ACL " & strACLName & " on " & strHostname
      end if ' End of checking for error prior to analysis
      iNameRow = iNameRow + 1
      if (iError = 1 or iError > MaxError) then
        if strNoMatch = "" and strMissing = "" and strNotes = "" then
          strNotes = "Good"
        else
          if strNotes <> "Failed to connect " then
            if strNoMatch <> "" or strMissing <> "" then strNotes = strNotes & "Updates required"
            bNewChange = True
            if dictChange.Exists(strChange) then
              bNewChange = False
              if not dictDevAffected.Exists(strHostname) then
                dictChange.item(strChange) = dictChange.item(strChange) & vbcrlf & strHostname & " " & strACLName
                dictDevAffected.add strHostname,""
              end if
            else
              dictChange.add strChange,strHostname & " " & strACLName
            end if
          end if
        end if
        objFileOut.writeline strIPAddr & "," & strHostname & "," & strACLName & "," & strNotes
      end if
    loop until wsNames.Cells(iNameRow,1).value = ""
    ' objLogOut.writeline "iError:" & iError & " MaxError:" & MaxError & " iVarRow:" & iVarRow
    if iError > MaxError then objLogOut.writeline "No connection after " & MaxError & " attempts, giving up and moving on."
    if (iError = 1 or iError > MaxError) then
      iVarRow = iVarRow + 1
      ' objLogOut.writeline "iVarRow now:" & iVarRow
    end if
    if wsVars.Cells(iVarRow,iHostCol).Value = "" or iFailed > 0 then
      ' objLogOut.writeline "evalute complete. wsVars.Cells(" & iVarRow & ",1).Value=" & wsVars.Cells(iVarRow,1).Value
      ' objLogOut.writeline "iFailed:" & iFailed
      if dictFailed.count > 0 then
        if iFailed = 0 then
          objLogOut.writeline "There are " & dictFailed.count & " devices I couldn't connect to. Here is the list:"
          dkeys = dictFailed.keys
          for each dkey in dkeys
            objLogOut.writeline dkey & " on line " & dictFailed(dkey)
          next
          objLogOut.writeline "Going to retry those one more time"
          dItems = dictFailed.items
        end if
        if iFailed = dictFailed.count then exit do
        iVarRow = dItems(iFailed)
        if iFailed < dictFailed.count then
          iFailed = iFailed + 1
        else
          exit do
        end if
      else
        exit do
      end if
    end if
  loop  ' This is the end of the loop to go through the Variable sheet
  dkeys = dictChange.keys
  iChangeID = 1
  for each dkey in dkeys
    set objChangeOut = CreateFile(strMOPPath & "HPNAScript-Change" & iChangeID & ".txt")
    objChangeOut.writeline "************ Devices Affected ************" & vbcrlf & dictChange.item(dkey) & vbcrlf & "******************************************"
    objChangeOut.writeline dkey
    CreateCSVs dictChange.item(dkey),dkey,iChangeID
    iChangeID = iChangeID + 1
    objChangeOut.close
  next

  objLogOut.writeline "All done at " & now()
  objLogOut.close
  objFileOut.close
  CleanUp
End Sub ' End of the Main sub

Sub CreateCSVs (strDevlist, strChange, iChangeID)
'-------------------------------------------------------------------------------------------------'
' Sub CreateCSVs (strDevlist, strChange, iChangeID)                                               '
'                                                                                                 '
' This sub takes a Devicelist (CRLF seperated), configuration script and changeID and generates   '
' all the CSV files to be used by both HPNA and PIER to push out those changes.                   '
'-------------------------------------------------------------------------------------------------'
dim strDevListParts, strChangeLines, x, y, strVarCol, dictDevices, iRow, iVarColList, objHPNAout
dim iStartPos, iStopPos, strACLVar, strColHead, iCol, objPIERout, strDevTemp, dictDev

  set objHPNAout  = CreateFile(strMOPPath & "HPNAVars-Change" & iChangeID & ".csv")
  set objPIERout  = CreateFile(strMOPPath & "PIERDuplicate-Change" & iChangeID & ".csv")
  Set dictDevices = CreateObject("Scripting.Dictionary")
  Set dictDev     = CreateObject("Scripting.Dictionary")
  strDevListParts = split(strDevlist,vbcrlf)
  strChangeLines  = split(strChange,vbcrlf)
  strVarCol = ""
  objPIERout.writeline "CR_Id,TaskOrder_to_Duplicate,Config_Item,Copy_Attachment,Copy_Manual_Approver,Copy_Assignee,Copy_Schedules"
  for x = 0 to ubound(strChangeLines)
    iStartPos = instr (1,strChangeLines(x),"$",vbTextCompare) ' Look for $ which indicates a start of a variable in the ACL standard.
    if iStartPos > 0 then ' If the current line has a variable parse out the variable, and substitute it with the proper value.
      iStopPos = instr (iStartPos+1,strChangeLines(x),"$",vbTextCompare) ' Locate the end of the variable name.
      strACLVar = mid(strChangeLines(x),iStartPos+1,iStopPos-iStartPos-1) ' Store the name of the variable.
      if dictVars.Exists(strACLVar) then ' If the variable we found exists in the variable sheet.
        strVarCol = strVarCol & dictVars(strACLVar) & ","
      end if ' end if variable exists
    end if
  next
  ' Grab the device name column of the variable sheet and stick them into a dictionary.
  iRow=2
  dictDevices.removeall
  Do until wsVars.Cells(iRow,iHostCol).Value = ""
    ' objLogOut.writeline "dictDevices.Add " & wsVars.Cells(iRow,iHostCol).value & " " & wsVars.Cells(iRow,iACLNameCol).value & ", " & iRow
    If not dictDevices.Exists(wsVars.Cells(iRow,iHostCol).value & " " & wsVars.Cells(iRow,iACLNameCol).value) then
      dictDevices.Add wsVars.Cells(iRow,iHostCol).value & " " & wsVars.Cells(iRow,iACLNameCol).value, iRow
    End If
    iRow = iRow + 1
  loop
  iRow=2
  dictDev.removeall
  Do until wsVars.Cells(iRow,iHostCol).Value = ""
    ' objLogOut.writeline "dictDev.Add " & wsVars.Cells(iRow,iHostCol).value & ", " & iRow
    If not dictDev.Exists(wsVars.Cells(iRow,iHostCol).value) then
      dictDev.Add wsVars.Cells(iRow,iHostCol).value, iRow
    End If
    iRow = iRow + 1
  loop
  iVarColList = split(strVarCol,",")
  objHPNAout.write wsVars.Cells(1,iIPCol).value
  objHPNAout.write "," & wsVars.Cells(1,iHostCol).value
  for x=0 to ubound(iVarColList)
    if IsNumeric(iVarColList(x)) then
      iCol = cint(iVarColList(x))
      objHPNAout.write "," & wsVars.Cells(1,iCol).value
    end if
  next
  objHPNAout.writeline
  for x=0 to ubound(strDevListParts)
    ' objLogOut.writeline "strDevListParts(" & x & ") = " & strDevListParts(x)
    strDevTemp = split(strDevListParts(x), " ")
    if strACLNameVar = "" then
      if dictDev.Exists(strDevTemp(0)) then
        iRow = dictDev(strDevTemp(0))
        ' objLogOut.writeline "iRow=" & iRow
      else
        objLogOut.writeline "something weird just happened, can't find " & strDevTemp(0) & " in the spreadsheet!"
      end if
    else
      if dictDevices.Exists(strDevListParts(x)) then
        iRow = dictDevices(strDevListParts(x))
        ' objLogOut.writeline "iRow=" & iRow
      else
        objLogOut.writeline "something weird just happened, can't find " & strDevListParts(x) & " in the spreadsheet!"
      end if
    end if
    objHPNAout.write wsVars.Cells(iRow,iIPCol).value
    objHPNAout.write "," & wsVars.Cells(iRow,iHostCol).value
    objPIERout.writeline ",," & wsVars.Cells(iRow,iHostCol).value & ",No,Yes,Yes,No"
    for y=0 to ubound(iVarColList)
      if IsNumeric(iVarColList(y)) then
        iCol = cint(iVarColList(y))
        objHPNAout.write "," & wsVars.Cells(iRow,iCol).value
      end if
    next
    objHPNAout.writeline
  next
  objHPNAout.Close
  objPIERout.Close
  set objHPNAout = Nothing
  set objPIERout = Nothing
  Set dictDevices = Nothing

End Sub ' End of CreateCSVs Sub

Sub CleanUp()
'-------------------------------------------------------------------------------------------------'
' Sub CleanUp()                                                               '
'                                                                                                 '
' This is a cleanup function.             '
'-------------------------------------------------------------------------------------------------'
  crt.Session.Disconnect
  if AutoCloseInput = True then
    if IsObject(wbin) then wbin.Close
  end if

  if IsObject(app) then Set wbin = app.Workbooks.Open (strOutFile,0,False)
  if IsObject(objShell) then objShell.run ("notepad " & strlogFile)

  if IsObject(wbin) then Set wbin = Nothing
  if IsObject(wsNames) then Set wsNames = Nothing
  if IsObject(wsVars) then Set wsVars = Nothing
  if IsObject(wsACL) then Set wsACL = Nothing
  if IsObject(app) then Set app = Nothing
  if IsObject(objShell) then set objShell = Nothing
  if IsObject(objACLGen) then
    objACLGen.close
    set objACLGen = Nothing
  end if
  if IsObject(objChangeOut) then
    objChangeOut.Close
    set objChangeOut = Nothing
  end if
  if IsObject(objLogOut) then
    objLogOut.Close
    set objLogOut = Nothing
  end if
  if IsObject(objFileOut) then
    objFileOut.Close
    set objFileOut = Nothing
  end if
end Sub ' End of CleanUp sub

Function CreatePath (strFullPath)
'-------------------------------------------------------------------------------------------------'
' Function CreatePath (strFullPath)                                                               '
'                                                                                                 '
' This function takes a complete path as input and builds that path out as nessisary.             '
'-------------------------------------------------------------------------------------------------'
dim pathparts, buildpath, part

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
end Function ' End of CreatePath Function

Function CreateFile (strFilePath)
'-----------------------------------------------------------------------------------------------------'
' Function CreateFile (strFilePath)                                                                    '
'                                                                                                     '
' This function takes a filepath and returns a file handle, while doing all nessisary error handling. '
'-----------------------------------------------------------------------------------------------------'
dim objFileOut, iInt, strOrigional, iPos, iLen

  iInt = 1
  strOrigional = strFilePath
  iPos = InStrRev (strFilePath,".") - 1
  iLen = len(strFilePath)
  on error resume next
  set objFileOut = fso.OpenTextFile(strFilePath, ForWriting, True)
  if err.number <> 0 Then
    if err.number = 70 then
      while err.number = 70
        strFilePath = left(strFilePath, iPos) & "-" & iInt & right(strFilePath, iLen-iPos)
        err.clear
        set objFileOut = fso.OpenTextFile(strFilePath, ForWriting, True)
        iInt = iInt + 1
        objLogOut.writeline "trying " & strFilePath
      wend
      objLogOut.writeline "Permission denied error when attempting to create file " & strOrigional & ". Created " & strFilePath & " instead."
    else
      MsgBox ("Create file Error # " & CStr(Err.Number) & " " & Err.Description)
      Exit Function
    end if
  end if
  on error goto 0
  set CreateFile = objFileOut
end Function ' End of CreateFile Function

Function GetAsIsACL(strIPVer, strACLName, strHostname, iError)
'-----------------------------------------------------------------------------------------------------'
' Function GetAsIsACL()                                                                               '
'                                                                                                     '
' This function logs into a router and feteches the current AsIs ACL.                                 '
'-----------------------------------------------------------------------------------------------------'
dim strVerifyCmd, strConnection, strResult, objACLAsIs, szHostName, objTab, objConfig

    strVerifyCmd = "show run " & strIPVer & " access-list " & strACLName ' construct the verification command to run.

    ' If session is connected, disconnect it.
    If crt.Session.Connected Then
      Set objTab = crt.GetScriptTab ' Hook into the current SecureCRT Tab
      Set objConfig = objTab.Session.Config ' Grab the session configuration for the current tab
      szHostName = objConfig.GetOption("Hostname") ' Get the currently connected hostname from the current session configuration
      if szHostName <> strHostname then crt.Session.Disconnect ' Unless the current connection is to the router we need disconnect
    end if

    ' Make new connection unless already connected.
    If not crt.Session.Connected Then
      strConnection = "/SSH2 /ACCEPTHOSTKEYS "  & strHostname ' connect string
      on error resume next
      crt.Session.Connect strConnection
      on error goto 0
    end if

    If crt.Session.Connected Then ' If we have a successful connection, run the verification command, write the ACL to a file and keep it in a variable.
      iError = 1 ' Ensure the error counter is set back to 1, which is default and means no error.
      crt.Screen.Synchronous = True
      crt.Screen.WaitForString "#",Timeout
      crt.Screen.Send("term len 0" & vbcr)
      crt.Screen.WaitForString "#",Timeout
      crt.Screen.Send(strVerifyCmd & vbcr)
      crt.Screen.WaitForString vbcrlf,Timeout
      strResult=trim(crt.Screen.Readstring (vbcrlf&"RP/",Timeout))
      ' crt.Session.Disconnect
      set objACLAsIs = CreateFile(strAsIsOutPath & strHostname & "-" & strACLName & ".txt")
      objACLAsIs.write strResult
      objACLAsIs.close
      set objACLAsIs = Nothing
      GetAsIsACL = split (strResult,vbcrlf)
    else ' If no connection, increase an error counter and note the failure.
      objLogOut.writeline "No connection to " & strHostname & " " & strACLName & ". Attempt #" & iError
      GetAsIsACL = iError + 1
      ' strNotes = "Failed to connect "
    end if ' End of connection verification

end Function ' End GetAsIsACL function

Function GetSeq (strLine)
'-----------------------------------------------------------------------------------------------------'
' Function GetSeq (strLine)                                                                           '
'                                                                                                     '
' This function takes in a string and returns the first token/word which should be the line Seq#      '                                 '
'-----------------------------------------------------------------------------------------------------'
dim iSeq

  if instr(1,trim(strLine)," ",vbTextCompare) > 0 then
    iSeq = trim(left(strLine,instr(1,trim(strLine)," ",vbTextCompare)))
    ' objLogOut.writeline "Found seq:" & iSeq
    if IsNumeric(iSeq) then
      ' objLogOut.writeline "confirmed it's a number"
      GetSeq = CLng(iSeq)
    else
      GetSeq = -1
      ' objLogOut.writeline GetSeq & " is not a number"
    end if
  else
    GetSeq = -1
    ' objLogOut.writeline "No space in: " & strLine
  end if

end Function ' Function GetSeq