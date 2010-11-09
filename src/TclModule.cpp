// $Id$
#pragma warning(disable: 4786)
#include "TclObject.h"
#include "TclModule.h"
#include "RegistryKey.h"

int
TclModule::registerFactoryByScript (const std::string &clsid)
{
    // Get registry key containing initialization data.
    std::string subkeyName("CLSID\\");
    subkeyName += clsid;
    subkeyName += "\\tcom";
    RegistryKey extensionKey(HKEY_CLASSES_ROOT, subkeyName);

    // Initialize Tcl interpreter.
    std::string tclDllPath;
    try {
        tclDllPath = extensionKey.value("TclDLL");
    }
    catch (std::runtime_error &)
    { }

    m_interp.initialize(tclDllPath);

    // Execute Tcl script which should register a class factory.
    std::string script = extensionKey.value("Script");
    int completionCode = m_interp.eval(script);
    if (completionCode != TCL_OK) {
        const char *errMsg = m_interp.resultString();
        MessageBox(NULL, errMsg, "tcom Server Error", MB_OK);
    }

    return completionCode;
}

void
TclModule::terminate ()
{
    revokeFactories();
    m_interp.terminate();
}
