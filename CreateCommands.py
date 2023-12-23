# $language = "Python"
# $interface = "1.0"

'''
Script that reads in a csv of commands
Creates Commands in SecureCRT

Author Siggi Bjarnason Dec 2023
Copyright 2023 Siggi Bjarnason

'''
# Import libraries
import os
import csv

def GetConfigPath():
    objConfig = crt.OpenSessionConfiguration("Default")
    # Try and get at where the configuration folder is located. To achieve
    # this goal, we'll use one of SecureCRT's cross-platform path
    # directives that means "THE path this instance of SecureCRT
    # is using to load/save its configuration": ${VDS_CONFIG_PATH}.

    # First, let's use a session setting that we know will do the
    # translation between the cross-platform moniker ${VDS_CONFIG_PATH}
    # and the actual value... say, "Upload Directory V2"
    strOptionName = "Upload Directory V2"

    # Stash the original value, so we can restore it later...
    strOrigValue = objConfig.GetOption(strOptionName)

    # Now set the value to our moniker...
    objConfig.SetOption(strOptionName, "${VDS_CONFIG_PATH}")
    # Make the change, so that the above templated name will get written
    # to the config...
    objConfig.Save()

    # Now, load a fresh copy of the config, and pull the option... so
    # that SecureCRT will convert from the template path value to the
    # actual path value:
    objConfig = crt.OpenSessionConfiguration("Default")
    strConfigPath = objConfig.GetOption(strOptionName)

    # Now, let's restore the setting to its original value
    objConfig.SetOption(strOptionName, strOrigValue)
    objConfig.Save()

    # Now return the config path
    return strConfigPath

strConfigPath = GetConfigPath()
strConfigPath = strConfigPath.replace("\\", "/")
if strConfigPath[-1:] != "/":
  strConfigPath += "/"

strCommandPath = strConfigPath + "Commands/"
if not os.path.exists(strCommandPath):
  crt.Dialog.MessageBox("Commands folder {} doesn't exists, creating it".format(strCommandPath))
  os.makedirs(strCommandPath)

strCommandFile = strCommandPath + "__Commands__.ini"
objFileOut = open(strCommandFile,"w")
objFileOut.write('D:"Is Command List"=00000001\nZ:"Default"=00000014\n')

strInFile = crt.Dialog.FileOpenDialog(title="Please select the Input File")
try:
  objInFile = open(strInFile,"r")
except Exception as err:
    crt.Dialog.MessageBox("Failed to open infile")
    objFileOut.close()
    sys.exit(0)
objCSVIn = open("D:\OneDrive\Documents\CommandList.txt","r")
csvReader = csv.reader(objCSVIn, delimiter=";")
for lstRow in csvReader:
   objFileOut.write(" SEND,{},{},,,0,10,,\n".format(lstRow[0],lstRow[1]))

objFileOut.close()
objCSVIn.close()
crt.Quit()