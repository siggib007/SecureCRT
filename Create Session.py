# $language = "Python"
# $interface = "1.0"

def createSession(strSessionPath):
  try:
    objSession = crt.OpenSessionConfiguration(strSessionPath)
  except Exception as err:
    objSession = crt.OpenSessionConfiguration()
  objSession.SetOption("Hostname","test5.supergeek.is")
  objSession.SetOption("Username","siggib")
  # objSession.SetOption("[SSH2] Port",22)
  #objSession.SetOption("Session Password Saved", 1)
  #objSession.SetOption("Password","MyTesting123!")
  objSession.Save(strSessionPath)

def readSession(strSessionPath):
  objSession = crt.OpenSessionConfiguration(strSessionPath)
  strHostname = objSession.GetOption("Hostname")
  crt.Dialog.MessageBox("Hostname:\r\n{0}".format(strHostname))
  strPassword = objSession.GetOption("Password")
  crt.Dialog.MessageBox("Password:\r\n{0}".format(strPassword))

def ListSession():
  objSessions = crt.Session
  if isinstance(objSessions,list):
    crt.Dialog.MessageBox("Sessions object is a list")
  else:
    crt.Dialog.MessageBox("Sessions object is not a list")

#ListSession()
createSession("test/test5")
#readSession("test/Second Test")
crt.Dialog.MessageBox("Done")