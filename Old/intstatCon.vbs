option explicit

Sub main
  Dim host, app, wb, ws, wbName, HostArray(14), lineparts, linesread, AscValue, PortStat
  Dim row, screenrow, readline, objShell, cmd, intname, Hostline, strHeaders, x, IntShort

  wbName = "C:\temp\ConTest3.xlsx"

HostArray(0) = "CSMBOT11 31"
HostArray(1) = "CSMAUS12 31"
HostArray(2) = "CSMSTL12 31"
HostArray(3) = "CSMWAY14 25"
HostArray(4) = "CSMWAY14 27"
HostArray(5) = "CSMWAY14 9"
HostArray(6) = "CSMWAY14 13"
HostArray(7) = "CSMWAY14 16"
HostArray(8) = "CSMWAY14 5"
HostArray(9) = "CSMWAY14 6"

	
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160

  crt.screen.synchronous = true
  crt.screen.IgnoreCase = true
  crt.screen.IgnoreEscape = True

  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)
  app.visible = true
  ws.Name = "Interfaces"
  strHeaders = Array("Device Name", "Interface", "code", "Status", "String")
  For x = 0 To UBound(strHeaders)
  	ws.Cells(1, x+1).Value = strHeaders(x)
  Next

  row = 2

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
			
		  PortStat = ""
  		ws.Cells(row,1).Value = host
  		ws.Cells(row,2).Value = intname
  		cmd = "/SSH2 "  & host
  		If crt.Session.Connected Then crt.Session.Disconnect
  		crt.Session.Connect cmd

  		crt.Screen.WaitForString "]>"
  		crt.Screen.Send("c d d " & intname & vbcr  )
  		crt.Screen.WaitForString "Escape sequence is ESC A" & vbcr, 30
  		crt.Screen.Send(vbcr)
 			crt.Screen.WaitForString vbcr, 2
  		while crt.Screen.WaitForString (vbcr, 2)
  		wend
 			screenrow = crt.screen.CurrentRow
 			readline = trim(crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns ))
 			if readline <> "" then
 				AscValue = asc(mid(readline,1,1))
 				if AscValue > 32 and AscValue < 127 then
 					PortStat="OK"
 				else
 					Portstat = "Fail"
 				end if
				ws.Cells(row, 3).Value = AscValue
			else
				Portstat = "Fail"
			end if
			ws.Cells(row, 4).Value = Portstat
  		ws.Cells(row, 5).Value = readline
 			row = row + 1
	 		crt.Screen.Send chr(27) & "A"
	 		crt.Screen.WaitForString "]>"
	 		crt.Screen.Send "logout" & vbcr
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
'  wb.Close
'  app.Quit

  Set ws = nothing
  Set wb = nothing
  Set app = nothing

  crt.screen.synchronous = false
'  objShell.Run(wbName)
End Sub