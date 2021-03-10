Sub main
  Dim host, user, passwd, app, wb, ws, wbName, wbCount, waitStrs, strHeaders
  Dim row, screenrow, readline, items, x, objShell, strResult, strCompleteOutput
  
  wbName = "C:\temp\ARGInts.xls"
  host = "argsnq01"
  user = "siggib"
  
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160

  ' Prompt for a password instead of embedding it in a script...
  ' 
  'passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

  ' Build a command-line string to pass to the Connect method.
  '
  'cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " /C 3DES /M MD5 " & host
  cmd = "/SSH2 "  & host
  If crt.Session.Connected Then crt.Session.Disconnect
  crt.Session.Connect cmd

  crt.screen.synchronous = true

  
  ' Create an Excel workbook/worksheet
  '

  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)
  ws.Name = "Interfaces"
	strCompleteOutput = ""
  ' Send the initial command to run and wait for the first linefeed
  '
  crt.Screen.WaitForString "#"
  crt.Screen.Send("show ip interface brief"  & vbcr )
  crt.Screen.WaitForString vbcr
  do
  	strResult = crt.screen.ReadString("#", "--More--")
  	strCompleteOutput = strCompleteOutput & strResult
  	If crt.Screen.MatchIndex = 2 Then crt.Screen.Send " "
  	If crt.Screen.MatchIndex = 1 Then Exit Do
	loop
  ws.Cells(1, 1).Value = strCompleteOutput
  ws.select
  wb.SaveAs(wbName)
  wb.Close
  app.Quit

  Set ws = nothing
  Set wb = nothing
  Set app = nothing

  crt.Screen.Send "exit" & vbcr
  crt.screen.synchronous = false
  objShell.Run(wbName)
End Sub