' This script demonstrates how ActiveX scripting can be used to
' interact with CRT and manipulate other programs such as Microsoft Excel
' through an OLE automation interface. This script creates an instance of Excel, 
' then it sends a command to a remote server (assuming we're already 
' connected). It reads the output, parses it and writes out some of the
' data to an Excel spreadsheet and saves it. This script also demonstrates
' how the WaitForStrings function can be used to wait for more than one
' output string.

Sub main

  crt.screen.synchronous = true

  ' Create an Excel workbook/worksheet
  '
  Dim app, wb, ws, wbName, wbCount
  Dim waitStrs, strHeaders
  Dim row, screenrow, readline, items, x
  Dim objShell
  Const xlHAlignCenter  = -4108
  Const xlVAlignBottom  = -4107
  Const xlVAlignCenter = -4108
  Const xlVAlignTop = -4160
  
  wbName = "C:\temp\chart3.xls"
  Set objShell = CreateObject("WScript.Shell")
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set ws = wb.Worksheets(1)

  ' Send the initial command to run and wait for the first linefeed
  '
  crt.Screen.Send("cat /etc/passwd"  & Chr(10) )
  crt.Screen.WaitForString Chr(10)

  ' Create an array of strings to wait for.
  '
  waitStrs = Array( Chr(10), "]$" )
  
  'Header Array
  strHeaders = Array("Username", "Password", "User ID (UID)", "Group ID (GID)", "User ID Info", "Home directory", "Command/shell")
  For x = 0 To UBound(strHeaders)
  	ws.Cells(1, x+1).Value = strHeaders(x)
  Next 

  row = 2

  Do
    While True

      ' Wait for the linefeed at the end of each line, or the shell prompt
      ' that indicates we're done.
      '	
      result = crt.Screen.WaitForStrings( waitStrs )

      ' We saw the prompt, we're done.
      '
      If result = 2 Then
        Exit Do
      End If

      ' Fetch current row and read the whole line from the screen
      ' on that row. Note, since we read a linefeed character subtract 1 
      ' from the return value of CurrentRow to read the actual line.
      '
      screenrow = crt.screen.CurrentRow -1
      readline = crt.Screen.Get(screenrow, 1, screenrow, crt.Screen.Columns )

      ' Split the line ( ":" delimited) and put some fields into Excel
      '
      items = Split( readline, ":", -1 )

      For x = 0 To UBound(items)
      	ws.Cells(row, x+1).Value = items(x)
      Next 
      row = row + 1
    Wend
  Loop
  
  With ws.Range(ws.Cells(1, 1), ws.Cells(1, x))
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
  ws.Name = "passwd"
  
'  MsgBox "Name first tab, total number of tabs: " & wbCount
  While wb.worksheets.count > 1
  	wb.Worksheets(2).delete
  Wend

  wb.Sheets.Add ,ws
'  MsgBox "Name first tab, total number of tabs: " & wb.worksheets.count
  Set ws2 = wb.Worksheets(2)
  ws2.Name = "Demo"
  With ws2.Cells(5,2)
  	.value="Just a Silly Demo"
  	    With .Font
        	.Name = "Calibri"
        	.Size = 48
        	.Strikethrough = False
        	.Superscript = False
        	.Subscript = False
        	.OutlineFont = False
        	.Shadow = False
        	.Color = -16776961
        End With
     .Interior.Color = 65535    
  End With
  ws2.Range(ws2.Cells(5, 2), ws2.Cells(5, 8)).Merge
'  ws.move ws2
  wb.Sheets.Add ,ws2
  Set ws3 = wb.Worksheets(3)
'  ws.move ws3
'  ws2.move ws3
  ws3.name="Demo2"
  ws3.cells(3,3).value ="more silly testing"
  ws3.Cells(3, 3).EntireColumn.AutoFit

  ws.select
  wb.SaveAs(wbName)
  wb.Close
  app.Quit

  Set ws = nothing
  Set wb = nothing
  Set app = nothing

  crt.screen.synchronous = false
  objShell.Run(wbName)
'  MsgBox "Done"
End Sub