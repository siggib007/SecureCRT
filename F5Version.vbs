' Connect to an SSH server using the SSH2 protocol. Specify the
' username and password and hostname on the command line as well as
' some SSH2 protocol specific options.
option Explicit

Sub Main

  Dim host, cmd, strVersion, user, result, screenrow
  host = "serou042"
   
  user = "sbjarna"

  ' Prompt for a password instead of embedding it in a script...
  '
  Dim passwd
  'passwd = crt.Dialog.Prompt("Enter password for " & host, "Login", "", True)

  ' Build a command-line string to pass to the Connect method.
  '
  'cmd = "/SSH2 /L " & user & " /PASSWORD " & passwd & " /C 3DES /M MD5 " & host
  cmd = "/SSH2 "  & host
  crt.Session.Connect cmd
   '
  crt.Screen.Synchronous = True
  crt.Screen.WaitForString( "#" )
  crt.Screen.Send("show /sys version" & vbCR )
  crt.Screen.WaitForString "Version ",5 
   
  strVersion = trim(crt.screen.Readstring(vbCR,5))
  msgbox "Version: " & strVersion

  ' When we get here the cursor should be one line below the output
  ' of the 'show clock' command. Subtract one line and use that value to
  ' read a chunk of text (1 row, 40 characters) from the screen.
  ' 
  screenrow = crt.screen.CurrentRow - 5
  result = crt.Screen.Get(screenrow, 11, screenrow, 20 )

  ' Get() reads a fixed size of the screen. So you may need to use
  ' VBScript's regular expression functions or the Split() function to
  ' do some simple parsing if necessary. Just print it out here.
  '
  msgbox "results are: " & result
  crt.Screen.Synchronous = False
  crt.session.disconnect

End Sub
