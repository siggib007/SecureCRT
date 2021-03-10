#$Language="VBScript" 
#$Interface="1.0" 
' ImportArbitraryDataFromFileToSecureCRTSessions.txt
'   (Designed for use with SecureCRT 6.7 and later)
'   Last Modified: 21 Mar, 2012
'
' DESCRIPTION
' This sample script is designed to create sessions from a text file (.csv
' format by default, but this can be edited to fit the format you have).
'
' To launch this script, map a button on the button bar to run this script:
'    http://www.vandyke.com/support/tips/buttonbar.html
'
' The first line of your data file should contain a comma-separated (or whatever
' you define as the g_strDelimiter below) list of supported "fields" designated
' by the following keywords:
' -----------------------------------------------------------------------------
' session_name: The name that should be used for the session. If this field
'               does not exist, the hostname field is used as the session_name.
'       folder: Relative path for session as displayed in the Connect dialog.
'     hostname: The hostname or IP for the remote server (For TAPI protocol,
'               specify the phone number here in the hostname field, for
'               example: 505-555-1212 or 5055551212 or 555-1212 or 5551212).
'               Alternatively, you could use the "port" field to specify the
'               area code, and use the "hostname" field to store the 7-digit
'               phone number.  If you choose to store all 10-digits in the
'               hostname field, make sure that you do not include a port field
'               in your datafile.
'     protocol: The protocol (SSH2, SSH1, telnet, rlogin, TAPI)
'         port: The port on which remote server is listening
'     username: The username for the account on the remote server
'    emulation: The emulation (vt100, xterm, etc.)
'  description: The comment/description. Multiple lines are separated with '\r'
' =============================================================================
'
'
' As mentioned above, the first line of the data file instructs this script as
' to the format of the fields in your data file and their meaning.  It is not a
' requirement that all the options be used. For example, notice the first line
' of the following file only uses the "hostname", "username", and "protocol"
' fields.  Note also that the "protocol" field can be defaulted so that if a
' protocol field is empty it will use the default value.
' -----------------------------------------------------------------------------
'   hostname,username,protocol=SSH2
'   192.168.0.1,root,SSH1
'   192.168.0.2,administrator,SSH2
'   192.168.0.3,root,
'   192.168.0.4,root,
'   192.168.0.5,administrator,telnet
'   ... and so on
' =============================================================================
'
'
' The g_strDefaultProtocol variable will only be defined within the
' ValidateFieldDesignations function if the protocol field has a default value
' (e.g., protocol=TAPI), as read in from the first line of the data file.
Dim g_strDefaultProtocol


' If your data file uses spaces or a character other than comma as the
' delimiter, you would also need to edit the g_strDelimiter value a few lines
' below to indicate that fields are separated by spaces, rather than by commas.
' For example:
'   g_strDelimiter = " "

' If you are importing TAPI sessions, beware of using "," as your delimiter
' because the comma character is actually a valid dialing instruction (pause).
' Using a ";" might be a good alternative for a file that includes TAPI info.
Dim g_strDelimiter
g_strDelimiter = ","     ' comma
' g_strDelimiter = " "    ' space
' g_strDelimiter = ";"    ' semi-colon
' g_strDelimiter = chr(9) ' tab
' g_strDelimiter = "|||"  ' a more unique example of a delimiter.

' The g_strSupportedFields indicates which of all the possible fields, are
' supported in this example script.  If a field designation is found in a data
' file that is not listed in this variable, it will not be imported into the
' session configuration.
Dim g_strSupportedFields
g_strSupportedFields = _
    "description,emulation,folder,hostname,port,protocol,session_name,username"

' If you wish to overwrite existing sessions, set the
' g_bOverwriteExistingSessions to True; for this example script, we're playing
' it safe and leaving any existing sessions in place :).
Dim g_bOverwriteExistingSessions
g_bOverwriteExistingSessions = False

Dim g_fso, g_shell
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_shell = CreateObject("WScript.Shell")

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim g_strHostsFile, g_strExampleHostsFile, g_strMyDocs
g_strMyDocs = g_shell.SpecialFolders("MyDocuments")
g_strHostsFile = g_strMyDocs & "\MyDataFile.csv"
g_strExampleHostsFile = _
    vbtab & "hostname,protocol,username,folder,emulation" & vbcrlf & _
    vbtab & "192.168.0.1,SSH2,root,Linux Machines,XTerm" & vbcrlf & _
    vbtab & "192.168.0.2,SSH2,root,Linux Machines,XTerm" & vbcrlf & _
    vbtab & "..." & vbcrlf & _
    vbtab & "10.0.100.1,SSH1,admin,CISCO Routers,VT100" & vbcrlf & _
    vbtab & "10.0.101.1,SSH1,admin,CISCO Routers,VT100" & vbcrlf & _
    vbtab & "..." & vbcrlf & _
    vbtab & "myhost.domain.com,SSH2,administrator,Windows Servers,VShell" & _
    vbtab & "..." & vbcrlf & _
g_strExampleHostsFile = Replace(g_strExampleHostsFile, ",", g_strDelimiter)

Dim g_strConfigFolder, strFieldDesignations, vFieldsArray, vSessionInfo

Dim strSessionName, strHostName, strPort
Dim strUserName, strProtocol, strEmulation
Dim strPathForSessions, strLine, nFieldIndex
Dim strSessionFileName, strFolder, nDescriptionLineCount, strDescription

Dim g_strLastError, g_strErrors, g_strSessionsCreated
Dim g_nSessionsCreated, g_nDataLines

' Use WMI to get at the current time values.  This info will be used
' to avoid overwriting existing sessions by naming new sessions with
' the current (unique) timestamp.
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objItem In colItems
    strLocalDateTime = objItem.LocalDateTime
    Exit For
Next
' strLocalDateTime has the following pattern:
' 20111013093717.418000-360   [ That is,  YYYYMMDDHHMMSS.MILLIS(zone) ]
g_strDateTimeTag = Left(strLocalDateTime, 18)


Import

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub Import()

    g_strHostsFile = crt.Dialog.FileOpenDialog( _
        "Please select the host data file to be imported.", _
        "Open", _
        g_strHostsFile, _
        "Text Files (*.txt)|*.txt|CSV File (*.csv)|*.csv||")

    If g_strHostsFile = "" Then
        Exit Sub
    End If

    ' Open our data file for reading
    Dim objDataFile
    Set objDataFile = g_fso.OpenTextFile(g_strHostsFile, ForReading, False)

    ' Now read the first line of the data file to determine the field
    ' designations
    strFieldDesignations = LCase(objDataFile.ReadLine)

    ' Validate the data file
    If Not ValidateFieldDesignations(strFieldDesignations) Then
        objDataFile.Close
        Exit Sub
    End If

    ' Get a timer reading so that we can calculate how long it takes to import.
    nStartTime = Timer

    ' Here we create an array of the items that will be used to create the new
    ' session, based on the fields separated by the delimiter specified in
    ' g_strDelimiter
    vFieldsArray = Split(strFieldDesignations, g_strDelimiter)

    ' Loop through reading each line in the data file and creating a session
    ' based on the information contained on each line.
    Do While Not objDataFile.AtEndOfStream
        strLine = ""
        strLine = objDataFile.ReadLine

        ' This sets v_File Data array elements to each section of strLine,
        ' separated by the delimiter
        vSessionInfo = Split(strLine, g_strDelimiter)
        If UBound(vSessionInfo) < UBound(vFieldsArray) Then
            If Trim(strLine) <> "" Then
                g_strErrors = g_strErrors & vbcrlf & _
                    "Insufficient data on line #" & _
                    NN(objDataFile.Line - 1, 4) & ": " & strLine
            Else
                g_strErrors = g_strErrors & vbcrlf & _
                    "Insufficient data on line #" & _
                    NN(objDataFile.Line - 1, 4) & ": [Empty Line]"
            End If
        Else

            ' Variable used to determine if a session file should actually be
            ' created, or if there was an unrecoverable error (and the session
            ' should be skipped).
            Dim bSaveSession
            bSaveSession = True

            ' Now we will match the items from the new file array to the correct
            ' variable for the session's ini file
            For nFieldIndex = 0 To UBound(vSessionInfo)

                Select Case vFieldsArray(nFieldIndex)
                    Case "session_name"
                        strSessionName = vSessionInfo(nFieldIndex)
                        ' Check folder name for any invalid characters
                        Dim re
                        Set re = New RegExp
                        re.Pattern = "[\\\|\/\:\*\?\""\<\>]"
                        If re.Test(strSessionName) Then
                            bSaveSession = False
                            If g_strErrors <> "" Then g_strErrors = _
                                vbcrlf & g_strErrors

                            g_strErrors = _
                                "Error: " & _
                                "Invalid characters found in SessionName """ & _
                                strSessionName & """ specified on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                ": " & strLine & g_strErrors
                        End If

                    Case "port"
                        strPort = Trim(vSessionInfo(nFieldIndex))
                        If Not IsNumeric(strPort) Then
                            bSaveSession = False
                            If g_strErrors <> "" Then g_strErrors = _
                                vbcrlf & g_strErrors

                            g_strErrors = _
                                "Error: Invalid port """ & strPort & _
                                """ specified on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                ": " & strLine & g_strErrors
                        End If

                    Case "protocol"
                        strProtocol = Trim(lcase(vSessionInfo(nFieldIndex)))

                        Select Case strProtocol
                            Case "ssh2"
                                strProtocol = "SSH2"
                            Case "ssh1"
                                strProtocol = "SSH1"
                            Case "telnet"
                                strProtocol = "Telnet"
                            Case "serial"
                                bSaveSession = False
                                g_strErrors = g_strErrors & vbcrlf & _
                                    "Warning: This sample script does " & _
                                    "not support creating sessions " & _
                                    "with protocol """ & _
                                    vSessionInfo(nFieldIndex) & _
                                    """ specified on line #" & _
                                    NN(objDataFile.Line - 1, 4) & _
                                    ": " & strLine
                            Case "tapi"
                                strProtocol = "TAPI"
                                bSaveSession = False
                                If g_strErrors <> "" Then g_strErrors = _
                                    vbcrlf & g_strErrors

                                g_strErrors = _
                                    "Error: Unsupported protocol """ & _
                                    vSessionInfo(nFieldIndex) & _
                                    """ specified on line #" & _
                                    NN(objDataFile.Line - 1, 4) & _
                                    ": " & strLine & g_strErrors

                            Case "rlogin"
                                strProtocol = "RLogin"
                            Case Else
                                If g_strDefaultProtocol <> "" Then
                                    strProtocol = g_strDefaultProtocol
                                Else
                                    bSaveSession = False
                                    If g_strErrors <> "" Then g_strErrors = _
                                        vbcrlf & g_strErrors

                                    g_strErrors = _
                                        "Error: Invalid protocol """ & _
                                        vSessionInfo(nFieldIndex) & _
                                        """ specified on line #" & _
                                        NN(objDataFile.Line - 1, 4) & _
                                        ": " & strLine & g_strErrors
                                End If
                        End Select ' for protocols

                    Case "hostname"
                        strHostName = Trim(vSessionInfo(nFieldIndex))
                        If strHostName = "" Then
                            bSaveSession = False
                            g_strErrors = g_strErrors & vbcrlf & _
                                "Warning: 'hostname' field on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                " is empty: " & strLine
                        End If

                    Case "username"
                        strUserName = Trim(vSessionInfo(nFieldIndex))

                    Case "emulation"
                        strEmulation = LCase(Trim(vSessionInfo(nFieldIndex)))
                        Select Case strEmulation
                            Case "xterm"
                                strEmulation = "Xterm"
                            Case "vt100"
                                strEmulation = "VT100"
                            Case "vt102"
                                strEmulation = "VT102"
                            Case "vt220"
                                strEmulation = "VT220"
                            Case "ansi"
                                strEmulation = "ANSI"
                            Case "linux"
                                strEmulation = "Linux"
                            Case "scoansi"
                                strEmulation = "SCOANSI"
                            Case "vshell"
                                strEmulation = "VShell"
                            Case "wyse50"
                                strEmulation = "WYSE50"
                            Case "wyse60"
                                strEmulation = "WYSE60"
                            Case Else
                                bSaveSession = False
                                g_strErrors = g_strErrors & vbcrlf & _
                                    "Warning: Invalid emulation """ & _
                                    strEmulation & """ specified on line #" & _
                                    NN(objDataFile.Line - 1, 4) & _
                                    ": " & strLine
                        End Select

                    Case "folder"
                        strFolder = Trim(vSessionInfo(nFieldIndex))

                        ' Check folder name for any invalid characters
                        ' Note that a folder can have subfolder designations,
                        ' so '/' is a valid character for the folder (path).
                        Set re = New RegExp
                        re.Pattern = "[\|\:\*\?\""\<\>]"
                        If re.Test(strFolder) Then
                            bSaveSession = False
                            If g_strErrors <> "" Then g_strErrors = _
                                vbcrlf & g_strErrors

                            g_strErrors = _
                                "Error: Invalid characters in folder """ & _
                                strFolder & """ specified on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                ": " & strLine & g_strErrors
                        End If

                    Case "description"
                        strDescription = Trim(vSessionInfo(nFieldIndex))
                        If strDescription = "" Then
                            g_strErrors = g_strErrors & vbcrlf & _
                                "Warning: 'description' field on line #" & _
                                NN(objDataFile.Line - 1, 4) & _
                                " is empty: " & strLine
                        End If

                    Case Else
                        ' If there is an entry that the script is not set to use
                        ' in strFieldDesignations, stop the script and display a
                        ' message
                        Dim strMsg1
                        strMsg1 = "Error: Unknown field designation: " & _
                            vFieldsArray(nFieldIndex) & vbcrlf & vbcrlf & _
                            "       Supported fields are as follows: " & _
                            vbcrlf & vbcrlf & vbtab & g_strSupportedFields & _
                            vbcrlf & _
                            vbcrlf & "       For a description of " & _
                            "supported fields, please see the comments in " & _
                            "the sample script file."

                        If Trim(g_strErrors) <> "" Then
                            strMsg1 = strMsg1 & vbcrlf & vbcrlf & _
                                "Other errors found so far include: " & _
                                g_strErrors
                        End If

                        MsgBox strMsg1, _
                            vbOkOnly, _
                            "Import Data To SecureCRT Sessions: Data File Error"
                        Exit Sub
                End Select
            Next

            If bSaveSession Then
                ' Use hostname if a session_name field wasn't present
                If strSessionName = "" Then
                    strSessionName = strHostName
                End If

                ' Canonicalize the path to the session, as needed
                strSessionPath = strSessionName
                If strFolder <> "" Then
                    strSessionPath = strFolder & "/" & strSessionName
                End If
                ' Strip any leading '/' characters from the session path
                If Left(strSessionPath, 1) = "/" Then
                    strSessionPath = Mid(strSessionPath, 2)
                End If

                If SessionExists(strSessionPath) Then
                    If Not g_bOverwriteExistingSessions Then
                        ' Append a unique tag to the session name, if it already exists
                        strSessionPath = strSessionPath & _
                            "(import_" & g_strDateTimeTag & ")"
                    End If
                End If

                ' Now: Create the session.

                ' Copy the default session settings into new session name and set the
                ' protocol.  Setting protocol protocol is essential since some variables
                ' within a config are only available with certain protocols.  For example,
                ' a telnet configuration will not be allowed to set any port forwarding
                ' settings since port forwarding settings are specific to SSH.
                Set objConfig = crt.OpenSessionConfiguration("Default")
                objConfig.SetOption "Protocol Name", strProtocol

                ' We opened a default session & changed the protocol, now we save the
                ' config to the new session path:
                objConfig.Save strSessionPath

                ' Now, let's open the new session configuration we've saved, and set
                ' up the various parameters that were specified in the file.
                Set objConfig = crt.OpenSessionConfiguration(strSessionPath)

                objConfig.SetOption "Emulation", strEmulation

                If LCase(strProtocol) <> "serial" And LCase(strProtocol) <> "tapi" Then
                    If strHostName <> "" Then
                        objConfig.SetOption "Hostname", strHostName
                    End If

                    If strUserName <> "" Then
                        objConfig.SetOption "Username", strUserName
                    End If
                End If

                If strDescription <> "" Then
                    objConfig.SetOption "Description", Split(strDescription, "\r")
                End If

                If UCase(strProtocol) = "SSH2" Then
                    If strPort = "" Then strPort = 22
                    objConfig.SetOption "[SSH2] Port", CInt(strPort)
                End If
                If UCase(strProtocol) = "SSH1" Then
                    If strPort = "" Then strPort = 22
                    objConfig.SetOption "[SSH1] Port", CInt(strPort)
                End If
                If UCase(strProtocol) = "TAPI" Then
                    ' TAPI is not currently supported because SecureCRT
                    ' (as of 7.2.x and earlier) doesn't support specification
                    ' of dialing list for TAPI sessions by way of a script
                    ' modifying the "Dialing List" cofiguration setting.
                    ' If SecureCRT's SessionConfiguration API did support
                    ' it, this is what it would look like:

                    ' For TAPI sessions, we use the "hostname" field for the
                    ' phone number to dial, and we use the port field for
                    ' the area code (if port field is not empty)
                    If strPort <> "" Then
                        strHostName = "[1]+1 (" & strPort & ") " & _
                            strHostName
                    ElseIf Instr(strHostname, "[1]") = 0 Then
                        strHostName = "[1]+1 " & strHostname
                    End If
                    objConfig.SetOption "Dialing List", strHostname
                End If
                If UCase(strProtocol) = "TELNET" Then
                    If strPort = "" Then strPort = 23
                    objConfig.SetOption "Port", CInt(strPort)
                End If

                ' If you would like ANSI Color enabled for all imported sessions (regardless
                ' of value in Default session, remove comment from following line)
                ' objConfig.SetOption "ANSI Color", True

                ' Add other "SetOption" calls desired here...
                ' objConfig.SetOption "Auto Reconnect", True
                ' objConfig.SetOption "Color Scheme", "Traditional"
                ' objConfig.SetOption "Color Scheme Overrides Ansi Color", True
                ' objConfig.SetOption "Copy to clipboard as RTF and plain text", True
                ' objConfig.SetOption "Description", Array("This session was imported from a script on " & Now)
                ' objConfig.SetOption "Firewall Name", "YOUR CUSTOM FIREWALL NAME HERE"
                ' objConfig.SetOption "Line Send Delay", 15
                ' objConfig.SetOption "Log Filename V2", "${VDS_USER_DATA_PATH}\_ScrtLog(%S)_%Y%M%D_%h%m%s.%t.txt"
                ' objConfig.SetOption "Rows", 60
                ' objConfig.SetOption "Cols", 140
                ' objConfig.SetOption "Start Tftp Server", True
                ' objConfig.SetOption "Use Word Delimiter Chars", True
                ' objConfig.SetOption "Word Delimiter Chars", " <>()+=$%!#*"
                ' objConfig.SetOption "X Position", 100
                ' objConfig.SetOption "Y Position", 50

                objConfig.Save

                If g_strSessionsCreated <> "" Then
                    g_strSessionsCreated = g_strSessionsCreated & vbcrlf
                End If
                g_strSessionsCreated = g_strSessionsCreated & "    " & strSessionPath

                g_nSessionsCreated = g_nSessionsCreated + 1

            End If

            ' Reset all variables in preparation for reading in the next line of
            ' the hosts info file.
            strEmulation = ""
            strPort = ""
            strHostName = ""
            strFolder = ""
            strUserName = ""
            strSessionName = ""
            strDescription = ""
            nDescriptionLineCount = 0
        End If

    Loop

    g_nDataLines = objDataFile.Line
    objDataFile.Close

    Dim strResults
    strResults = "Import operation completed in " & _
        GetMinutesAndSeconds(Timer - nStartTime)

    If g_nSessionsCreated > 0 Then
        strResults = strResults & _
            vbcrlf & _
            String(70, "-") & vbcrlf & _
            "Number of Sessions created: " & g_nSessionsCreated & vbcrlf & _
            g_strSessionsCreated
    Else
        strResults = strResults & vbcrlf & _
            String(70, "-") & vbcrlf & _
            "No sessions were created from " & g_nDataLines & " lines of data."
    End If

    ' Log activity information to a file for debugging purposes...
    strFilename = g_strMyDocs & "\__SecureCRT-Session-ImportLog-" & _
        Year(Now) & "-" & _
        NN(Month(Now), 2) & "-" & _
        NN(Day(Now), 2) & "--" & _
        NN(Hour(Now), 2) & "-" & _
        NN(Minute(Now), 2) & "-" & _
        NN(Second(Now), 2) & ".txt"
    Set objFile = g_fso.OpenTextFile(strFilename, ForWriting, True)
    objFile.Write _
        "Errors/warnings from this operation include:" & _
        g_strErrors & vbcrlf & _
        String(70, "-") & vbcrlf & _
        strResults & vbcrlf & vbcrlf & _
        ""
    objFile.Close

    ' Display the log file as an indication that the information has been
    ' imported.
    g_shell.Run chr(34) & strFilename & chr(34), 5, False
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'      Helper Methods and Functions
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ValidateFieldDesignations(ByRef strFields)
    If Instr(strFieldDesignations, g_strDelimiter) = 0 Then
        Dim strErrorMsg, strDelimiterDisplay
        strErrorMsg = "Invalid header line in data file. " & _
            "Delimiter character not found: "
        If Len(g_strDelimiter) > 1 Then
            strDelimiterDisplay = g_strDelimiter
        Else
            If Asc(g_strDelimiter) < 33 or Asc(g_strDelimiter) > 126 Then
                strDelimiterDisplay = "ASCII[" & Asc(g_strDelimiter) & "]"
            Else
                strDelimiterDisplay = g_strDelimiter
            End If
        End If
        strErrorMsg = strErrorMsg & strDelimiterDisplay & vbcrlf & vbcrlf & _
            "The first line of the data file is a header line " & _
            "that must include" & vbcrlf & _
            "a '" & strDelimiterDisplay & _
            "' separated list of field keywords." & vbcrlf & _
            vbcrlf & "'hostname' and 'protocol' are required keywords." & _
            vbcrlf & vbcrlf & _
            "The remainder of the lines in the file should follow the " & _
            vbcrlf & _
            "pattern established by the header line " & _
            "(first line in the file)." & vbcrlf & "For example:" & vbcrlf & _
            g_strExampleHostsFile
        MsgBox strErrorMsg, _
               vbOkOnly, _
               "Import Data To SecureCRT Sessions"
        Exit Function
    End If

    If Instr(strFieldDesignations, "hostname") = 0 Then
        strErrorMsg = "Invalid header line in data file. " & _
            "'hostname' field is required."
        If Len(g_strDelimiter) > 1 Then
            strDelimiterDisplay = g_strDelimiter
        Else
            If Asc(g_strDelimiter) < 33 Or Asc(g_strDelimiter) > 126 Then
                strDelimiterDisplay = "ASCII[" & Asc(g_strDelimiter) & "]"
            Else
                strDelimiterDisplay = g_strDelimiter
            End If
        End If

        MsgBox strErrorMsg & vbcrlf & _
            "The first line of the data file is a header line " & _
            "that must include" & vbcrlf & _
            "a '" & strDelimiterDisplay & _
            "' separated list of field keywords." & vbcrlf & _
            vbcrlf & "'hostname' and 'protocol' are required keywords." & _
            vbcrlf & vbcrlf & _
            "The remainder of the lines in the file should follow the " & _
            vbcrlf & _
            "pattern established by the header line " & _
            "(first line in the file)." & vbcrlf & "For example:" & vbcrlf & _
            g_strExampleHostsFile, _
            vbOkOnly, _
            "Import Data To SecureCRT Sessions"
        Exit Function
    End If

    If Instr(strFieldDesignations, "protocol") = 0 Then
        MsgBox "Invalid data file header line: " & vbcrlf & vbcrlf & _
            vbtab & strFieldDesignations & vbcrlf & _
            vbcrlf & "--> 'protocol' field is required.", _
            vbOkOnly, _
            "Import Data To SecureCRT Sessions"
        Exit Function
    Else
        ' We found "protocol", now look for a default protocol designation
        vFields = Split(strFields,g_strDelimiter)
        For each strField In vFields
            If (InStr(strField, "protocol") > 0) And _
               (Instr(strField, "=") >0) Then
                    g_strDefaultProtocol = UCase(Split(strField, "=")(1))

                    ' Fix the protocol field since we know the default protocol
                    ' value
                    strFields = Replace(strFields, strField, "protocol")
            End If
        Next
    End If

    ValidateFieldDesignations = True
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ReadRegKey(strKeyPath)
    On Error Resume Next
    Err.Clear
    ReadRegKey = g_shell.RegRead(strKeyPath)
    If Err.Number <> 0 Then
        ' Registry key must not have existed.
        ' ReadRegKey will already be empty, but for the sake of clarity, we'll
        ' set it to an empty string explicitly.
        ReadRegKey = ""
    End If
    On Error Goto 0
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function CreateFolderPath(strPath)
' Recursive function
    If g_fso.FolderExists(strPath) Then
        CreateFolderPath = True
        Exit Function
    End If

    ' Check to see if we've reached the drive root
    If Right(strPath, 2) = ":\" Then
        CreateFolderPath = True
        Exit Function
    End If

    ' None of the other two cases were successful, so attempt to create the
    ' folder
    On Error Resume Next
    g_fso.CreateFolder strPath
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    If nError <> 0 Then
        ' Error 76 = Path not found, meaning that the full path doesn't exist.
        ' Call ourselves recursively until all the parent folders have been
        ' created:
        If nError = 76 Then _
            CreateFolderPath(g_fso.GetParentFolderName(strPath))

        On Error Resume Next
        g_fso.CreateFolder strPath
        nError = Err.Number
        strErr = Err.Description
        On Error Goto 0

        ' If the Error is not = 76, then we have to bail since we no longer have
        ' any hope of successfully creating each folder in the tree
        If nError <> 0 Then
            g_strLastError = strErr
            Exit Function
        End If
    End If

    CreateFolderPath = True
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function NN(nNumber, nDesiredDigits)
' Normalizes a number to have a number of zeros in front of it so that the
' total length of the number (displayed as a string) is nDesiredDigits.
    Dim nIndex, nOffbyDigits, strResult
    nOffbyDigits = nDesiredDigits - Len(nNumber)

    NN = nNumber

    If nOffByDigits = 0 Then Exit Function

    If nOffByDigits > 0 Then
        ' The number provided doesn't have enough digits
        strResult = String(nOffbyDigits, "0") & nNumber
    Else
        ' The number provided has too many digits.

        nOffByDigits = Abs(nOffByDigits)

        ' Only remove leading digits if they're all insignificant (0).
        If Left(nNumber, nOffByDigits) = String(nOffByDigits, "0") Then
            strResult = Mid(nNumber, nOffByDigits + 1)
        Else
            ' If leading digits beyond desired number length aren't 0, we'll
            ' return the number as originally passed in.
            strResult = nNumber
        End If
    End If

    NN = strResult
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function GetMinutesAndSeconds(nTotalSecondsElapsed)
    Dim nMinutesElapsed, nSecondsValue, nSecondsElapsed

    If nTotalSecondsElapsed = 0 Then
        GetMinutesAndSeconds = "less than a second."
        Exit Function
    End If

    ' Convert seconds into a fractional minutes value.
    nMinutesElapsed = nTotalSecondsElapsed / 60

    ' Convert the decimal portion into the number of remaining seconds.
    nSecondsValue = nMinutesElapsed - Fix(nMinutesElapsed)
    nSecondsElapsed = Fix(nSecondsValue * 60)

    ' Remove the fraction portion of minutes value, keeping only the digits to
    ' the left of the decimal point.
    nMinutesElapsed = Fix(nMinutesElapsed)

    ' Calculate the number of milliseconds using the four most significant
    ' digits of only the decimal fraction portion of the number of seconds
    ' elapsed.
    nMSeconds = Fix(1000 * (nTotalSecondsElapsed - Fix(nTotalSecondsElapsed)))

    ' Form the final string to be returned and set it as the value of our
    ' function.
    GetMinutesAndSeconds = nMinutesElapsed & " minutes, " & _
        nSecondsElapsed & " seconds, and " & _
        nMSeconds & " ms"
End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function SessionExists(strSessionPath)
' Returns True if a session specified as value for strSessionPath already
' exists within the SecureCRT configuration.
' Returns False otherwise.
    On Error Resume Next
    Set objTosserConfig = crt.OpenSessionConfiguration(strSessionPath)
    nError = Err.Number
    strErr = Err.Description
    On Error Goto 0
    ' We only used this to detect an error indicating non-existance of session.
    ' Let's get rid of the reference now since we won't be using it:
    Set objTosserConfig = Nothing
    ' If there wasn't any error opening the session, then it's a 100% indication
    ' that the session named in strSessionPath already exists
    If nError = 0 Then
        SessionExists = True
    Else
        SessionExists = False
    End If
End Function