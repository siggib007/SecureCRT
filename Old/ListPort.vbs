Sub main
  Dim host, user, passwd, app, wb, ws, wbName, wbCount, waitStrs, strHeaders
  Dim row, screenrow, readline, items, x, objShell, strResult, strCompleteOutput
  
  wbName = "C:\temp\CatOS Ports.xlsx"
  host = "ASEATL13"
  user = "sbjarna"
  
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160

  ' Prompt for a password instead of embedding it in a script... 
  passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

	If crt.Session.Connected Then crt.Session.Disconnect
	crt.Session.Connect "/TELNET " & host & " 23"
	crt.Session.log false
  crt.Screen.WaitForString "name:"
	crt.Screen.Send user & vbCr
  crt.Screen.WaitForString "assword:"
  crt.Screen.Send passwd & vbCr
  
  crt.screen.synchronous = true

  
  ' Create an Excel workbook/worksheet
  '

  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)
  ws.Name = "Ports"
	strCompleteOutput = ""
  ' Send the initial command to run and wait for the first linefeed
  '
  crt.Screen.WaitForString "(enable)"
  crt.Screen.Send("show port"  & vbcr )
  crt.Screen.WaitForString vbcr
  do
  	strResult = crt.screen.ReadString("(enable)", "--More--")
  	strCompleteOutput = strCompleteOutput & strResult
  	If crt.Screen.MatchIndex = 2 Then crt.Screen.Send " "
  	If crt.Screen.MatchIndex = 2 Then Exit Do
	loop
  ws.Cells(1, 1).Value = strCompleteOutput
  ws.select
  wb.SaveAs(wbName)
  wb.Close
  app.Quit

  Set ws = nothing
  Set wb = nothing
  Set app = nothing
  
  crt.Screen.Send "q" & vbcr
  crt.Screen.Send "exit" & vbcr
  crt.screen.synchronous = false
  'objShell.Run(wbName)
End Sub