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
import csv

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

def isInt(CheckValue):
    """
    function to safely check if a value can be interpreded as an int
    Parameter:
      Value: A object to be evaluated
    Returns:
      Boolean indicating if the object is an integer or not.
    """
    if isinstance(CheckValue, (float, int, str)):
        try:
            fTemp = int(CheckValue)
        except ValueError:
            fTemp = "NULL"
    else:
        fTemp = "NULL"
    return fTemp != "NULL"

def createSession(dictSession):
  """
  This handles the actual creation of the session in SecureCRT
  Parameters:
    dictSession: dictionary of the session elements: Path, Hostname, User, Cred, FW and Port
  Returns:
    string with error or success
  """
  if "Label" not in dictSession:
    return "No Label"
  else:
    strLabel = dictSession["Label"] or ""
  if strLabel == "":
    return "No Label"

  if "Address" in dictSession:
    strAddress = dictSession["Address"] or ""
  else:
    strAddress = ""
  if strAddress == "":
    return "No Address"

  if "Group" in dictSession:
    strGroup = dictSession["Group"] or ""
  else:
    strGroup = ""
  if "User" in dictSession:
    strUser = dictSession["User"] or ""
  else:
    strUser = ""
  if "Credential" in dictSession:
    strCredential = dictSession["Credential"] or ""
  else:
    strCredential = ""
  if "Jump" in dictSession:
    strJump = dictSession["Jump"] or ""
  else:
    strJump = ""
  if strJump != "":
    if strGroup != "":
      strJump = "Session:{}/{}".format(strGroup,strJump)
    else:
      strJump = "Session:{}".format(strJump)
  if "Port" in dictSession:
    if isInt(dictSession["Port"]):
      iPort = int(dictSession["Port"])
    else:
      iPort = 22
  else:
    iPort = 22
  if strGroup != "":
    strPath = "{}/{}".format(strGroup,strLabel)
  else:
    strPath = strLabel


  LogEntry("Working on {} - {}".format(strLabel,strAddress),4)
  try:
    objSession = crt.OpenSessionConfiguration(dictSession["Path"])
  except Exception as err:
    objSession = crt.OpenSessionConfiguration()

  try:
    objSession.SetOption("Hostname",strAddress)
    objSession.SetOption("Username",strUser)
    objSession.SetOption("Credential Title",strCredential)
    objSession.SetOption("Firewall Name",strJump)
    objSession.SetOption("[SSH2] Port",iPort)
    objSession.Save(strPath)
  except Exception as err:
     return err
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

  strScriptName = os.path.basename(os.path.abspath(__file__))
  iLoc = strScriptName.rfind(".")
  strLogFile = strLogDir + strScriptName[:iLoc] + ISO + ".log"
  objLogOut = open(strLogFile, "w", 1)
  LogEntry("Starting up",3)
  strInFile = crt.Dialog.FileOpenDialog(title="Please select the Input File")
  try:
    objInFile = open(strInFile,"r")
  except Exception as err:
     LogEntry("failed to open file: {}".format(err),1)
     objLogOut.close()
     crt.Dialog.MessageBox("Failed to open infile")
     sys.exit(0)
  objReader = csv.DictReader(objInFile)
  for dictTemp in objReader:
    strRet = createSession(dictTemp)
    LogEntry("create session returned:{}".format(strRet),4)
  LogEntry("Done",3)
  objInFile.close()
  objLogOut.close()
  crt.Dialog.MessageBox("Done")

main()