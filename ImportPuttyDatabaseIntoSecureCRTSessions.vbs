# $language = "VBScript"
# $interface = "1.0"

' ImportPuttyDatabaseIntoSecureCRTSessions.vbs
' DESCRIPTION:
'
'   **************************************************************
'   *    NOTE: This script must be run from within SecureCRT.    *
'   **************************************************************
'
'    Imports putty session configurations into SecureCRT saved sessions.
'
'    This script will attempt to automatically export existing putty configuration
'    settings to a .reg file in your temporary folder. If this automated attempt
'    fails, putty session info will need to manually be exported to a .reg file,
'    as in running the following command within a CMD.exe shell window:
'        REG EXPORT HKCU\Software\SimonTatham\PuTTY\Sessions %USERPROFILE%\Documents\ptty.reg
'
'    To ensure that any pre-existing sessions are not overwritten, imported
'    sessions are stored SecureCRT's Session Manager in a new folder named:
'       ##_puttyImport(YYYYMMDD_HHMMSS.mmm)
'
' ---------------------------------------------------------------------
'   Last Modified:
'   5 May, 2021
'     - Changed default log location to be a user's "Documents" instead
'       of the user's "Desktop" folder since more and more individuals
'       are running in non-standard environments.
'
'   5 May, 2015
'     - Added block of code to remind users to run the script from within
'       SecureCRT. Added debug line message to log file.
'
'   15 Jan, 2015
'     - Fix importing of Serial-protocol sessions from Putty; some session
'       settings like flow control weren't getting converted properly from
'       putty to SecureCRT session options.
'     - Implement Debug logging for troubleshooting assistance. Set
'       'g_bEnableDebug = True' below to enable debug logging to a file.
'     - Automatically export the putty.reg file for the user, placing it
'       in the user's TEMP folder location.
'     - Support .reg file exports that aren't done with same case
'       as example shows (.reg file exported with 'reg export ...'
'       command line will contain reg key paths with same case (upper
'       vs. lower) as typed in at the command prompt).
'
'   17 Dec, 2013
'     - Some saved putty sessions might have characters that
'       are not allowed in filenames.  Added code to translate
'       such session names to ones that will work for
'       SecureCRT (function named GetSafeFilename).
'
'   22 Oct, 2013
'     - Added support for importing Serial configurations from putty
'
'   23 Mar, 2012
'     - Initial revision
' ---------------------------------------------------------------------

Dim g_fso, g_shell
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

Dim g_strLogFile, g_nStartTime, g_bEnableDebug
g_bEnableDebug = True
g_nStartTime = Timer
g_strLogFile = g_shell.SpecialFolders("MyDocuments") & _
    "\SCRT_Ptty_ImportScriptLog_" & _
    GetTimeDateWithMilliseconds & ".txt"

LogLine "######################################################################"
LogLine "Script Starting"

Dim vRegsArray
Dim nLineCount, strRegFile, strSessionName, strHostName, strPortNumber
Dim strUserName, strProtocol, strProtocolType, strEmulation, strPathForSessions
Dim strSerialSpeed, strSerialDataBits, strSerialStopBits, strSerialParity, strSerialFlowCtrl

Dim g_nSessionsCreated, g_strWarnings
g_nSessionsCreated = 0

Dim g_colSessionsCreated
Set g_colSessionsCreated = CreateObject("Scripting.Dictionary")

Dim g_nTotalSessionsToImport
g_nTotalSessionsToImport = 0

' Script will be looking for the session header in the array.  It will be a pattern
' like:
'   [HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions\local2]
' But not
'   [HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions]
Set g_regexpNewSession = New RegExp
g_regexpNewSession.IgnoreCase = True
g_regexpNewSession.Global = True
g_regexpNewSession.Multiline = True
g_regexpNewSession.Pattern = _
  "\[HKEY_CURRENT_USER\\Software\\SimonTatham\\PuTTY\\Sessions\\(.*)\]"

' Create the regular expression to find the information from the Putty
' registry file
Set g_regexpPuttyInfo = New RegExp
g_regexpPuttyInfo.IgnoreCase = True
g_regexpPuttyInfo.Global = True
g_regexpPuttyInfo.Multiline = False
g_regexpPuttyInfo.Pattern = _
    """(?:(HostName|UserName|Protocol|PortNumber|" & _
    "SerialLine|SerialSpeed|SerialDataBits|SerialStopHalfBits|" & _
    "SerialParity|SerialFlowControl|SshProt|" & _
    "PortForwardings))""=(?:""(.*)""|dword:(.*))"

' GetTimeDateWithMilliseconds is a function defined in this script which
' will return something like:
'    20150114093717.418   [ That is,  YYYYMMDD_HHMMSS.mmm ]
g_strImportDestinationFolder = "##_puttyImport(" & GetTimeDateWithMilliseconds & ")"

ImportFromPutty

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub ImportFromPutty()

    On Error Resume Next
        strVersion = crt.Version
        If Err.Number <> 0 Then
            LogLine "User attempted to execute script directly from VBScript, instead of within SecureCRT. Exiting ..."
            MsgBox "This script must be run from within SecureCRT."
            Exit Sub
        End If
    On Error GoTo 0

    ' Attempt to extract .reg file with existing putty configuration.
    strExportedRegFile = g_fso.GetSpecialFolder(2) & "\scrt_import_ptty.reg"
    If g_fso.FileExists(strExportedRegFile) Then
        LogLine "Removing existing exported putty .reg file: " & strExportedRegFile
        g_fso.DeleteFile(strExportedRegFile)
    End If

    strCmd = "reg export hkcu\Software\simonTatham\PuTTY\sessions """ & strExportedRegFile & """"
    LogLine "Attempting to export putty config to a .reg file we can read in, using command: " & strCmd
    nResult = g_shell.Run(strCmd, 0, True)
    LogLine "  --> Results of 'reg export' command: " & nResult
    If Not g_fso.FileExists(strExportedRegFile) Then
        LogLine "Failed to automatically export putty config to a .reg file. Prompting user for manual specification."
        strMessage = _
            "You will be asked to specify the .reg file exported from the " & _
            "putty registry. If you haven't yet exported the putty config, use " & _
            "a command similar to the following single-line command that may " & _
            "appear wrapped below: " & vbcrlf & vbcrlf & _
            "reg export HKCU\Software\SimonTatham\PuTTY\sessions" & _
            " %USERPROFILE%\Documents\ptty.reg"

        If crt.Dialog.MessageBox(_
            strMessage, _
            "Select Exported Putty .reg File for Import", _
        vbOKCancel) <> vbOK Then
        LogLine "User cancelled import operation. Script exiting."
        Exit Sub
    End If

    LogLine "Prompting user to select existing .reg file for import..."
        'Prompt for the name and location of the registry file
        strExportedRegFile = crt.Dialog.FileOpenDialog("Specify .reg file for import.", _
            "Import", _
            g_shell.SpecialFolders("MyDocuments") & "\*.reg", _
            "Registry Files (*.reg)|*.reg||")

        If strExportedRegFile = "" Then
            LogLine "User cancelled file selection. Script exiting."
            Exit Sub
        End If
    End If

    nLineCount = 0
    nImportStartTime = Timer

    ' Read in information from a file that contains the export of the
    ' registry
    LogLine "Opening exported registry file for reading: " & strExportedRegFile
    Dim objFso, objTextStream
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objTextStream = objFso.OpenTextFile(strExportedRegFile, 1, False)

    LogLine "Reading data from file..."
    Dim strFileText
    strFileText = objTextStream.ReadAll
    objTextStream.close
    LogLine "Read " & Len(strFileText) & " bytes."

    ' If the file is a unicode file, it will be gibberish since we opened it
    ' as ASCII earlier.  If we can't find the .reg file header as we expect,
    ' we'll need to read it in again, but in the proper format:
    If Instr(strFileText, "Windows Registry Editor") = 0 Then
        LogLine "Opening file as UNICODE for re-reading..."
        Set objTextStream = objFso.OpenTextFile(strExportedRegFile, 1, False, -1)

        LogLine "Reading data from file..."
        strFileText = objTextStream.ReadAll
        objTextStream.close
        LogLine "Read " & Len(strFileText) & " bytes."
    End If

    If g_regexpNewSession.Test(strFileText) Then
        Set objSessionMatches = g_regexpNewSession.Execute(strFileText)
        g_nTotalSessionsToImport = objSessionMatches.Count
        LogLine "File contains " & g_nTotalSessionsToImport & " saved connections for importing into SecureCRT."
    Else
        strMsg = "File does not contain any putty configurations exported from the registry"
        LogLine strMsg
        crt.Dialog.MessageBox strMsg
        Exit Sub
    End If

    LogLine "Splitting data we read from the file into an array of lines..."
    Dim n
    vRegsArray = Split(strFileText,vbcrlf)
    LogLine "... We now have an array of " & UBound(vRegsArray) - 1 & " lines."
    LogLine "Beginning our loop inspecting each line of the file until we find the first session definition."
    For n = 0 To UBound(vRegsArray)
        ' Read in each line and find the registry line that contains the
        ' session name so we know that this is the begining of a session
        ' definition.

        ' If the regular expression finds a match, then the session name needs
        ' to be parsed out, along with the settings we care about.
        If g_regexpNewSession.test(vRegsArray(n)) Then
            Set objMatches = g_regexpNewSession.Execute(vRegsArray(n))
            For Each objMatch In objMatches
                strOrigSessionName = objMatch.submatches(0)
                ' Putty likes to store session names in web URL format.  We'll
                ' take care of at least spaces here...
                strSessionName = Replace(strOrigSessionName, "%20", " ")
                strMsg = "Found session definition: " & strOrigSessionName
                If strSessionName <> strOrigSessionName Then
                    strMsg = strMsg & " (sanitized: " & strSessionName & ")"
                End If
                LogLine strMsg

                ' We've already found the line containing the session name.
                ' When we call CreateConfigFromPuttyData, we want to move to the
                ' next line first so that it doesn't mistake the current session
                ' for a new session definition and exit early.
                LogLine "Calling function to create config from putty data for session: " & strSessionName
                n = CreateConfigFromPuttyData(strSessionName, n + 1)
                Exit For
            Next
        End If
    Next

    If g_colSessionsCreated.Count > 0 Then
        strSummaryMsg = _
            "Successfully created " & g_colSessionsCreated.Count & _
            " sessions in " & _
            Round(Timer - nImportStartTime, 2) & " seconds : " & vbcrlf & _
            Join(g_colSessionsCreated.Keys(), ", ") & vbcrlf & vbcrlf & _
            "You may need to close and re-open the Session Manager in order to see the imported sessions."
    Else
        strSummaryMsg = _
            "No putty session configurations were detected in the .reg file" & _
            " you specified: " & strRegFile
    End If

    crt.Dialog.MessageBox strSummaryMsg & vbcrlf & g_strWarnings

    ' Open up the Connect dialog automatically.
    crt.Screen.SendSpecial "MENU_TOGGLE_SESSION_MANAGER"

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function CreateConfigFromPuttyData(strSessionName, nStartIndex)
    ' Found the header information and now need to find:
    ' Hostname ("HostName"=)
    ' Port ("PortNumber"=)
    ' Username ("UserName"=)
    ' Protocol ("Protocol"=)
    ' Emulation ("TerminalType"=)
    ' Tunnels ("PortForwardings"=)


    Dim x
    LogLine "CreateConfigFromPuttyData::Beginning loop at x=" & nStartIndex
    For x = nStartIndex To uBound(vRegsArray)
        If g_regexpNewSession.Test(vRegsArray(x)) Then
            ' We've found another session definition, so we're done parsing
            ' the info for the current session.
            LogLine "    ######> Found a new session header on this line, so we're done. Exiting function."
            Exit For
        End If

        ' Now set the submatches from the regular expression as variables that
        ' can be used to write the new session file
        If g_regexpPuttyInfo.Test(vRegsArray(x)) Then
            Set Matches = g_regexpPuttyInfo.Execute(vRegsArray(x))
            For Each match In Matches
                Select Case match.submatches(0)
                    Case "HostName"
                        strHostName = match.submatches(1)
                        LogLine "Found hostname line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "PortNumber"
                        strPortNumber = match.submatches(2)
                        LogLine "Found port line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "UserName"
                        strUserName = match.submatches(1)
                        LogLine "Found username line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "Protocol"
                        strProtocol = match.submatches(1)
                        LogLine "Found protocol name line:(" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "SshProt"
                        strProtocolType = match.submatches(2)
                        LogLine "Found SSH protocol type line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "PortForwardings"
                        strPortForwardings = match.submatches(1)
                        LogLine "Found port-forward line (" & NN(x, 4) & "): " & vRegsArray(x)
                        strSerialPort = match.submatches(1)
                    Case "SerialLine"
                        strSerialPort = match.submatches(1)
                        LogLine "Found serial com port line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "SerialSpeed"
                        strSerialSpeed = match.submatches(2)
                        LogLine "Found serial speed line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "SerialDataBits"
                        strSerialDataBits = match.submatches(2)
                        LogLine "Found serial data-bits line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "SerialStopHalfbits"
                        strSerialStopBits = match.submatches(2)
                        LogLine "Found serial stop-bits line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "SerialParity"
                        strSerialParity = match.submatches(2)
                        LogLine "Found serial parity line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case "SerialFlowControl"
                        strSerialFlowCtrl = match.submatches(2)
                        LogLine "Found serial flow control line (" & NN(x, 4) & "): " & vRegsArray(x)
                    Case Else
                        LogLine "    ##########> RegExp matched value, but submatch unknown/unhandled:" & match.submatches(0)
                End Select
            Next
        Else
            ' LogLine "Skipping Line (Importing this data is not supported)"
        End If
    Next

    ' Set correct SSH protocol type. If strProtcolType is 00000003 or 00000002
    ' set protocol to SSH2 If it is 00000001 or 00000000 then use SSH1. In both
    ' cases, this should only occur when strProtocol is "ssh"
    If strProtocol = "ssh" Then
        Select Case strProtocolType
            Case "00000003", "00000002"
                strProtocol = "SSH2"
            Case "00000001", "00000000"
                strProtocol = "SSH1"
            Case Else
                ' Default to SSH2 if we don't recognize the protocol from putty
                strProtocol = "SSH2"
        End Select
    ElseIf strProtocol = "serial" Then
        ' Guard against session name being a reserved file handle in Windows (e.g. COM1, COM2)
        Set regexpCom = New RegExp
        regexpCom.IgnoreCase = True
        regexpCom.Global = True
        regexpCom.Multiline = False

        ' Create the regular expression to find reserved names
        regexpCom.Pattern = "^(?:COM\d+|PRN$|LPT\d+|CON$|AUX$|NUL$|CLOCK\$$)"
        If regexpCom.Test(strSessionName) Then
            strSessionName = "serial-" & strSessionName
        End If
    Else
        strProtocol = "Telnet"
    End If

    ' Copy the default session settings into new session name and set the
    ' protocol.  Setting protocol protocol is essential since some variables
    ' within a config are only available with certain protocols.  For example,
    ' a telnet configuration will not be allowed to set any port forwarding
    ' settings since port forwarding settings are specific to SSH.
    LogLine "Creating a brand new '" & strProtocol & "' session from the 'Default' config..."
    Set objConfig = crt.OpenSessionConfiguration("Default")
    objConfig.SetOption "Protocol Name", strProtocol

    ' Since we're placing all imported sessions into a uniquely-named folder
    ' within the Connection Manager (aka Connect dialog), we don't have to worry
    ' about duplicate configurations that might already exist in SecureCRT.
    LogLine "Current Session name: " & strSessionName
    objConfig.Save g_strImportDestinationFolder & "/" & GetSafeFilename(strSessionName)
    LogLine "Saved new session as: " & g_strImportDestinationFolder & "/" & GetSafeFilename(strSessionName)

    LogLine "Loading the new session so that we can populate settings..."
    Set objConfig = crt.OpenSessionConfiguration(g_strImportDestinationFolder & "/" & GetSafeFilename(strSessionName))

    Select Case LCase(strProtocol)
        Case "ssh2", "ssh1", "telnet"
            objConfig.SetOption "Hostname", strHostName
            objConfig.SetOption "Username", strUserName
            objConfig.SetOption "Emulation", "Xterm"
            objConfig.SetOption "ANSI Color", True

            If LCase(strProtocol) = "ssh1" Then
                objConfig.SetOption "[SSH1] Port", cInt("&H" & strPortNumber)
            ElseIf LCase(strProtocol) = "telnet" Then
                objConfig.SetOption "Port", cInt("&H" & strPortNumber)
            Else
                objConfig.SetOption "[SSH2] Port", cInt("&H" & strPortNumber)
            End If

            ' SecureCRT Port Forwarding Configuration format:
            ' Name|ListenHost,ListenPort|TargetHostDiff?|TargetHost|TargetPort||
            ' Z:"Port Forward Table V2"=00000005
            ' test1|1.1.1.1,22|1|remote|2222||
            ' test2|22|1|remote|2222||
            ' test3|22|0||2222||
            ' DynName|ListenHost,ListenPort|0|socks,|TargetPort(notusedindynamic)||
            ' test7|7.7.7.7,22|0|socks,|22||
            ' test8|22|0|socks,|22||

            ' SecureCRT Reverse/Remote Port Forwarding Configuration format:
            ' Name|ListenHost,ListenPort|TargetHostDiff?|TargetHost|TargetPort||
            ' Z:"Reverse Forward Table V2"=00000003
            ' test4|4.4.4.4,22|1|remote4|2222||
            ' test5|225|1|remote5|2222||
            ' test6|226|0||226||

            ' Putty Port Forwarding Configuration
            ' "PortForwardings"="L2222=127.0.0.1:22,R2222=localhost:22,D2222,"

            If strPortForwardings <> "" Then
                strPFarray = Split(strPortForwardings, ",")

                nLocalIndex = 1
                nRemoteIndex = 1
                nDynamicIndex = 1

                Set regexp = New RegExp
                regexp.IgnoreCase = True
                regexp.Global = True
                regexp.Multiline = False
                regexp.Pattern = "(L|R)(?:([^\:]+?)\:)*?(\d+)=(?:([^\:]+?)\:)*?(\d+)$"

                For Each Forward in strPFarray
                    If Forward <> "" Then
                        If Left(Forward, 1) = "D" Or Left(Forward, 2) = "4D" OR _
                           Left(Forward, 2) = "6D" Then
                            ' Handle dynamic pf
                            ' Handle case where PF entry from Putty export does not
                            ' match a known pattern. We will store the warnings in a
                            ' global variable to be displayed
                            If g_strWarnings = "" Then
                                g_strWarnings = vbcrlf & string(40, "_") & vbcrlf
                            End If
                            g_strWarnings = g_strWarnings & _
                                "Warning: Import of dynamic port fwd settings is " & _
                                "not currently supported (session """ & _
                                strSessionName & """)"
                        Else
                            If regexp.Test(Forward) Then
                                Set objMatch = regExp.Execute(Forward)(0)
                                strType = objMatch.submatches(0)
                                strListenHost = objMatch.submatches(1)
                                strListenPort = objMatch.submatches(2)
                                strTargetHost = objMatch.submatches(3)
                                strTargetPort = objMatch.submatches(4)
                                strListenSpec = strListenPort
                                ' Prepend host info if present
                                If strListenHost <> "" Then strListenSpec = _
                                    strListenHost & "," & strListenSpec

                                nTargetIsDiff = 0
                                If strTargetHost <> "" Then nTargetIsDiff = 1
                                strCommon = "|" & strListenSpec & _
                                            "|" & nTargetIsDiff & _
                                            "|" & strTargetHost & _
                                            "|" & strTargetPort & _
                                            "||"

                                Select Case strType
                                    Case "L"
                                        If strLocalPF <> "" Then
                                            strLocalPF = strLocalPF & vbcrlf
                                        End If
                                        strLocalPF = strLocalPF & _
                                            "Local " & nLocalIndex & strCommon

                                        nLocalIndex = nLocalIndex + 1

                                    Case "R"
                                        If strRemotePF <> "" Then
                                            strRemotePF = strRemotePF & vbcrlf
                                        End If
                                        strRemotePF = strRemotePF & _
                                            "Remote " & nRemoteIndex & strCommon

                                        nRemoteIndex = nRemoteIndex + 1

                                End Select
                            Else
                                ' Handle case where PF entry from Putty export does not
                                ' match a known pattern. We will store the warnings in a
                                ' global variable to be displayed
                                If g_strWarnings = "" Then
                                    g_strWarnings = vbcrlf & string(40, "_") & vbcrlf
                                End If
                                g_strWarnings = g_strWarnings & _
                                    "Warning: """ & strSessionName & _
                                    """'s port forward setting unrecognized: " & _
                                    vbcrlf & vbtab & Forward & vbcrlf
                            End If
                        End If
                    End If
                Next

                If InStr(strProtocol, "SSH") > 0 Then
                    If strLocalPF <> "" Then
                        objConfig.SetOption _
                            "Port Forward Table V2", Split(strLocalPF, vbcrlf)
                    End If
                    If strRemotePF <> "" Then
                        objConfig.SetOption _
                            "Reverse Forward Table V2", Split(strRemotePF, vbcrlf)
                    End If
                End If
            End If

        Case "serial"
            LogLine "Setting COM port for serial session to: " & strSerialPort
            objConfig.SetOption "Com Port", strSerialPort
            ' Speed is hex, ie: 2580 = 9600, so use the hex digits appropriately
            ' to create a decimal value that can be passed to SetOption(), which
            ' takes decimal values.
            LogLine "             Setting Baud to: " & strSerialSpeed
            objConfig.SetOption "Baud Rate", CDbl("&H" & strSerialSpeed)
            LogLine "        Setting Data bits to: " & strSerialDataBits
            objConfig.SetOption "Data Bits", strSerialDataBits

            ' Putty: stop bits 0=0(not used in SCRT); 1=2; 1.5=3; 2=4
            ' In SecureCRT ini file: 0 (for 1 StopBit), 1 (for 1.5 StopBits), 2 (for 2 StopBits)
            nStopBits = -1
            Select Case strSerialStopBits
                Case 2
                    nStopBits = 0
                Case 3
                    nStopBits = 1
                Case 4
                    nStopBits = 2
                Case Else
                    nStopBits = 0
            End Select
            LogLine "        Setting stop bits to: " & nStopBits
            objConfig.SetOption "Stop Bits", nStopBits

            ' Putty: parity none=0; odd=1; even=2; mark=3; space=4
            LogLine "           Setting parity to: " & strSerialParity
            objConfig.SetOption "Parity", strSerialParity

            ' Putty: flow control none=0; xon/xoff=1; RTS=2; DSR=3
            ' Default for SecureCRT -- RTS/CTS is CTSFlow=1; RTSFlowCtrl=2)
            LogLine "Setting flow control for serial session based on value from putty: " & strSerialFlowCtrl
            '~ All Off (no flow control)
            '~ Putty: SerialFlowControl=0
            '~ ---------------------------------------------------------------------
            '~ D:"DSR Flow"=00000000
            '~ D:"DTR Flow Control"=00000001
            '~ D:"CTS Flow"=00000000
            '~ D:"RTS Flow Control"=00000001
            '~ D:"XON Flow"=00000000

            '~ XON/XOFF enabled
            '~ Putty: SerialFlowControl=1
            '~ ---------------------------------------------------------------------
            '~ D:"DSR Flow"=00000000
            '~ D:"DTR Flow Control"=00000001
            '~ D:"CTS Flow"=00000000
            '~ D:"RTS Flow Control"=00000001
            '~ D:"XON Flow"=00000001

            '~ RTS/CTS enabled
            '~ Putty: SerialFlowControl=2
            '~ ---------------------------------------------------------------------
            '~ D:"DSR Flow"=00000000
            '~ D:"DTR Flow Control"=00000001
            '~ D:"CTS Flow"=00000001
            '~ D:"RTS Flow Control"=00000002
            '~ D:"XON Flow"=00000000

            '~ DTR/DSR enabled
            '~ Putty: SerialFlowControl=3
            '~ ---------------------------------------------------------------------
            '~ D:"DSR Flow"=00000001
            '~ D:"DTR Flow Control"=00000002
            '~ D:"CTS Flow"=00000000
            '~ D:"RTS Flow Control"=00000001
            '~ D:"XON Flow"=00000000

            Select Case strSerialFlowCtrl
                Case 0
                    objConfig.SetOption "DSR Flow", 0
                    objConfig.SetOption "DTR Flow Control", 1
                    objConfig.SetOption "CTS Flow", 0
                    objConfig.SetOption "RTS Flow Control", 1
                    objConfig.SetOption "XON Flow", 0

                Case 1
                    objConfig.SetOption "DSR Flow", 0
                    objConfig.SetOption "DTR Flow Control", 1
                    objConfig.SetOption "CTS Flow", 0
                    objConfig.SetOption "RTS Flow Control", 1
                    objConfig.SetOption "XON Flow", 1

                Case 2
                    objConfig.SetOption "DSR Flow", 0
                    objConfig.SetOption "DTR Flow Control", 1
                    objConfig.SetOption "CTS Flow", 1
                    objConfig.SetOption "RTS Flow Control", 2
                    objConfig.SetOption "XON Flow", 0

                Case 3
                    objConfig.SetOption "DSR Flow", 1
                    objConfig.SetOption "DTR Flow Control", 2
                    objConfig.SetOption "CTS Flow", 0
                    objConfig.SetOption "RTS Flow Control", 1
                    objConfig.SetOption "XON Flow", 0

            End Select

    End Select

    ' Place session in a folder that indicates sessions were imported from
    ' putty (folder name already has unique time-stamp)
    LogLine "Saving changes to session configuration..."
    objConfig.Save g_strImportDestinationFolder & "/" & GetSafeFilename(strSessionName)

    On Error Resume Next
    g_colSessionsCreated.Add strSessionName, 1
    On Error Goto 0

    g_nSessionsCreated = g_nSessionsCreated + 1

    ' Since we found another session definition, we want to back up one line
    ' because the calling loop in Main() above will have already gone to the
    ' next line (if we don't back up, we'll potentially miss sessions).  This
    ' Function returns the line number that the caller should continue working
    ' on.
    CreateConfigFromPuttyData = x - 1

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function GetSafeFilename(strName)
    ' Replace any illegal characters that might have been introduced by the
    ' command we are running or the date.  Since Python can run on multiple
    ' platforms, replace illegal characters for all platforms.
    strName = Replace(strName, ":",  "-")
    strName = Replace(strName, "/",  "[SLASH]")
    strName = Replace(strName, "\",  "[BKSLASH]")
    strName = Replace(strName, ":",  "[COLON]")
    strName = Replace(strName, "*",  "[STAR]")
    strName = Replace(strName, "?",  "[QUESTION]")
    strName = Replace(strName, """", "[QUOTE]")
    strName = Replace(strName, "<",  "[LT]")
    strName = Replace(strName, ">",  "[GT]")
    strName = Replace(strName, "|",  "[PIPE]")
    GetSafeFilename = strName
End Function

'-----------------------------------------------------------------------------
Function GetTimeDateWithMilliseconds()
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objItem In colItems
        strLocalDateTime = objItem.LocalDateTime
        Exit For
    Next
    ' strLocalDateTime has the following pattern:
    ' 20111013093717.418000-360   [ That is,  YYYYMMDDHHMMSS.MILLIS(zone) ]
    ' Take the left-most 18 digits...
    strTimeDate = Left(strLocalDateTime, 18)
    strTimeDate = Left(strTimeDate, 8) & "_" & Mid(strTimeDate, 9)

    ' ... return the time/date string as value of the function in the following
    ' format:  20111013_093717.418
    GetTimeDateWithMilliseconds = strTimeDate
End Function

'-----------------------------------------------------------------------------
Function LogLine(strText)
    If Not g_bEnableDebug Then Exit Function
    nTimeElapsed = CStr(Round(Timer - g_nStartTime, 3))
    nPos = Instr(nTimeElapsed, ".")
    If nPos > 0 Then
        nNumDigits = Len(Mid(CStr(nTimeElapsed), nPos))
        If nNumDigits < 4 Then
            nDiff = 4 - nNumDigits
            nTimeElapsed = nTimeElapsed & String(nDiff, "0")
        End If
    Else
        nTimeElapsed = nTimeElapsed & ".000"
    End If
    strLine = Now & " (" & nTimeElapsed & "): " & strText

    Set objFile = g_fso.OpenTextFile(g_strLogFile, 8, True)
    objFile.WriteLine strLine
    objFile.Close
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function NormalizeNumber(nNumber, nDesiredDigits)
' Normalizes a single digit number to have a 0 in front of it
    Dim nIndex, nOffbyDigits, strResult
    nOffbyDigits = nDesiredDigits - len(nNumber)

    strResult = nNumber

    For nIndex = 1 to nOffByDigits
        strResult = "0" & strResult
    Next
    NormalizeNumber = strResult
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function NN(nNumber, nDigits)
' Wrapper for NormalizeNumber function
    NN = NormalizeNumber(nNumber, nDigits)
End Function
