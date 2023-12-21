# $language = "Python"
# $interface = "1.0"

'''
Script that reads in a csv of session details
Creates session profiles in SecureCRT

Author Siggi Bjarnason Dec 2023
Copyright 2023 Siggi Bjarnason

'''
# Import libraries
import os
import time
import sys

iLogLevel = 5  # How much logging should be done. Level 10 is debug level, 0 is none

def CleanExit(strCause):
  """
  Handles cleaning things up before unexpected exit in case of an error.
  Things such as closing down open file handles, open database connections, etc.
  Logs any cause given, closes everything down then terminates the script.
  Remember to add things here that need to be cleaned up
  Parameters:
    Cause: simple string indicating cause of the termination, can be blank
  Returns:
    nothing as it terminates the script
  """
  if strCause != "":
    strMsg = "{} is exiting abnormally because: {}".format(
        strScriptName, strCause)
    strTimeStamp = time.strftime("%m-%d-%Y %H:%M:%S")
    objLogOut.write("{0} : {1}\n".format(strTimeStamp, strMsg))

  objLogOut.close()
  sys.exit(9)


def LogEntry(strMsg, iMsgLevel, bAbort=False):
  """
  This handles writing all event logs into the appropriate log facilities
  This could be a simple text log file, a database connection, etc.
  Needs to be customized as needed
  Parameters:
    Message: Simple string with the event to be logged
    iMsgLevel: How detailed is this message, debug level or general. Will be matched against Loglevel
    Abort: Optional, defaults to false. A boolean to indicate if CleanExit should be called.
  Returns:
    Nothing
  """
  if iLogLevel > iMsgLevel:
    strTimeStamp = time.strftime("%m-%d-%Y %H:%M:%S")
    objLogOut.write("{0} : {1}\n".format(strTimeStamp, strMsg))
  else:
    if bAbort:
        strTimeStamp = time.strftime("%m-%d-%Y %H:%M:%S")
        objLogOut.write("{0} : {1}\n".format(strTimeStamp, strMsg))

  if bAbort:
    CleanExit("")


def createSession(dictSession):
  """
  This handles the actual creation of the session in SecureCRT
  Parameters:
    dictSession: dictionary of the session elements: Path, Hostname, User, Cred, FW and Port
  Returns:
    string with error or success
  """
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

def main():
  global objLogOut
  global strScriptName

  ISO = time.strftime("-%Y-%m-%d-%H-%M-%S")
  strBaseDir = os.path.dirname(os.path.abspath(__file__))

  if strBaseDir[-1:] != "/":
      strBaseDir += "/"
  strLogDir = strBaseDir + "Logs/"
  if strLogDir[-1:] != "/":
      strLogDir += "/"
  if not os.path.exists(strLogDir):
      crt.Dialog.MessageBox("Attempting to create log directory: {}".format(strLogDir))
      os.makedirs(strLogDir)

  strScriptName = os.path.basename(sys.argv[0])
  iLoc = strScriptName.rfind(".")
  strLogFile = strLogDir + strScriptName[:iLoc] + ISO + ".log"
  objLogOut = open(strLogFile, "w", 1)
  LogEntry("Starting up",3)

  dictTemp = {}
  dictTemp["Path"] = "mytest"
  dictTemp["HostName"] = "mytest.supergeek.is"
  dictTemp["Cred"] = "siggi-key"
  dictTemp["FW"] = "Session:Nanitor\Nanitor Jump"
  strRet = createSession(dictTemp)
  LogEntry("create session returned:{}".format(strRet),4)
  LogEntry("Done",3)
  objLogOut.close()
  crt.Dialog.MessageBox("Done")

main()