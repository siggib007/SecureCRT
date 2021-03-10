option explicit

Sub main
  Dim host, user, passwd, app, wb, ws, wbName, waitStrs, HostArray(14), lineparts
  Dim row, screenrow, readline, objShell, cmd, intname, Hostline, strHeaders, x
  
  wbName = "C:\temp\Ints.xlsx"
  host = "argsnq01"
  intname = "g7/43"
  user = "siggib"
  
	HostArray(1)  = "ASACHR01 E2/36"
	HostArray(2)  = "ASACHR01 E2/37"
	HostArray(3)  = "ASACHR02 E2/36"
	HostArray(4)  = "ASACHR02 E2/37"
	HostArray(5)  = "ASAWSC01 E1/43"
	HostArray(6)  = "ASAWSC01 E1/44"
	HostArray(7)  = "ASAWSC02 E1/43"
	HostArray(8)  = "ASAWSC02 E1/44"
	HostArray(9)  = "ASACHR01 E2/38"
	HostArray(10) = "ASACHR02 E2/38"
	HostArray(11) = "ASAWSC01 E1/45"
	HostArray(12) = "ASAWSC02 E1/45"
	HostArray(13) = "ASADET01 e3/34"
	HostArray(14) = "ASADET02 e3/37"
  
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160

  crt.screen.synchronous = true
  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)
  ws.Name = "Interfaces"
  strHeaders = Array("Device Name", "Interface", "Status")
  For x = 0 To UBound(strHeaders)
  	ws.Cells(1, x+1).Value = strHeaders(x)
  Next 

  row = 2
  
  ' Prompt for a password instead of embedding it in a script...
  ' 
  'passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

  for each hostline in hostarray
  	' Build a command-line string to pass to the Connect method.
  	'
  	'cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " /C 3DES /M MD5 " & host
  	if hostline <> "" then
  		lineparts = split(hostline, " ")
  		host = lineparts(0)
  		intname = lineparts(1)
  		ws.Cells(row,1).Value = host
  		ws.Cells(row,2).Value = intname
  		cmd = "/SSH2 "  & host
  		If crt.Session.Connected Then crt.Session.Disconnect
  		crt.Session.Connect cmd
  		
  		crt.Screen.WaitForString "#"
  		crt.Screen.Send("show int " & intname & vbcr )
  		crt.Screen.WaitForString (vbcr)
  		crt.Screen.WaitForString (vbcr)
  		crt.Screen.WaitForString (vbcr)
  		
  		screenrow = crt.screen.CurrentRow
  		readline = trim(crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns ))
			ws.Cells(row, 3).Value = readline
'			crt.Screen.WaitForString (vbcr)
	  	crt.Screen.Send "exit" & vbcr
  		row = row + 1
  	end if
  next
  With ws.Range(ws.Cells(1, 1), ws.Cells(1, 3))
      .Font.Bold = True
      .HorizontalAlignment = xlHAlignCenter 
      .VerticalAlignment = xlVAlignBottom 
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .MergeCells = False
      .EntireColumn.AutoFit
  End With
  
  ws.select
  wb.SaveAs(wbName)
  wb.Close
  app.Quit

  Set ws = nothing
  Set wb = nothing
  Set app = nothing

  crt.screen.synchronous = false
  objShell.Run(wbName)
End Sub