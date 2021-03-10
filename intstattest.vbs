option explicit

Sub main
  Dim host, user, passwd, app, wb, ws, wbName, waitStrs, HostArray(14), lineparts
  Dim row, screenrow, readline, objShell, cmd, intname, Hostline, strHeaders, x, IntShort
  
  wbName = "C:\temp\IntTest.xlsx"
  user = "sbjarna"
  
	HostArray(0) 	= "DRCBOT11 FA3/82"
	HostArray(1) 	= "DRCAUS11 FA3/72"
	HostArray(2) 	= "DRCSTL11 F3/77"
	HostArray(3) 	= "DRCWAY11 F3/52"

	 
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160

  crt.screen.synchronous = true
  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)
  app.visible = true
  ws.Name = "Interfaces"
  strHeaders = Array("Device Name", "Interface", "Status")
  For x = 0 To UBound(strHeaders)
  	ws.Cells(1, x+1).Value = strHeaders(x)
  Next 

  row = 2
  
  passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

  for each hostline in hostarray
  	if hostline <> "" then
  		lineparts = split(hostline, " ")
  		host = lineparts(0)
  		intname = lineparts(1)
  			x = 1
			while (not IsNumeric(mid(intname,x,1))) and (x < len(intname))
				x = x + 1
			wend
			IntShort = mid(intname,x)
 
  		ws.Cells(row,1).Value = host
  		ws.Cells(row,2).Value = intname
  		If crt.Session.Connected Then crt.Session.Disconnect
	  	cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " " & host
  		crt.Session.Connect cmd
  		
  		crt.Screen.WaitForString "#"
  		crt.Screen.Send("show int " & intname & vbcr )
  		crt.Screen.WaitForString (vbcr)
  		crt.Screen.WaitForString (IntShort)
  		crt.Screen.WaitForString (vbcr)
  		
  		screenrow = crt.screen.CurrentRow
  		readline = trim(crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns ))
			ws.Cells(row, 3).Value = readline
  		row = row + 1
  		crt.Session.Disconnect
  		'crt.Screen.Send "exit" & vbcr
  		crt.dialog.messagebox "Disconnected from " & host
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


  Set ws = nothing
  Set wb = nothing
  Set app = nothing

  crt.screen.synchronous = false
End Sub