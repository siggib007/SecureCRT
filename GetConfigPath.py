# $language = "Python"
# $interface = "1.0"

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
    crt.Dialog.MessageBox("OrigValue:\r\n{0}".format(strOrigValue))

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



strPath = GetConfigPath()
crt.Dialog.MessageBox("Config Path:\r\n{0}".format(strPath))
