# $language = "Python"
# $interface = "1.0"

def createSession(dictSession):
  if "Path" not in dictSession:
    return "No Path"
  if dictSession["Path"] == "":
    return "No Path"

  if "HostName" in dictSession:
    strHostName = dictSession["HostName"]
  else:
    strHostName = ""
  if "User" in dictSession:
    strUser = dictSession["User"]
  else:
    strUser = ""
  if "Cred" in dictSession:
    strCred = dictSession["Cred"]
  else:
    strCred = ""
  if "FW" in dictSession:
    strFW = dictSession["FW"]
  else:
    strFW = ""
  if "Port" in dictSession:
    strPort = dictSession["Port"]
  else:
    strPort = 22

  try:
    objSession = crt.OpenSessionConfiguration(dictSession["Path"])
  except Exception as err:
    objSession = crt.OpenSessionConfiguration()

  objSession.SetOption("Hostname",strHostName)
  objSession.SetOption("Username",strUser)
  objSession.SetOption("Credential Title",strCred)
  objSession.SetOption("Firewall Name",strFW)
  objSession.SetOption("[SSH2] Port",strPort)
  objSession.Save(dictSession["Path"])
  return "Success"

dictTemp = {}
dictTemp["Path"] = "mytest"
dictTemp["HostName"] = "mytest.supergeek.is"
dictTemp["Cred"] = "siggi-key"
dictTemp["FW"] = "Session:Nanitor\Nanitor Jump"
createSession(dictTemp)
crt.Dialog.MessageBox("Done")