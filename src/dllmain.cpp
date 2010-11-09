// $Id: dllmain.cpp,v 1.16 2002/07/14 18:42:57 cthuang Exp $
#pragma warning(disable: 4786)
#include "Uuid.h"
#include "HandleSupport.h"
#include "TclModule.h"
#include "TclInterp.h"
#include "tclRunTime.h"

// This class implements a COM module for DLL-based servers.

class DllModule: public TclModule
{
public:
    DllModule ()
    { }

    virtual void initializeCom(DWORD coinitFlags);
};

static DllModule module;

void
DllModule::initializeCom (DWORD /*coinitFlags*/)
{
    // Do nothing.  In-process servers should not call CoInitializeEx.
}


STDAPI
DllCanUnloadNow () 
{
    return (module.lockCount() == 0) ? S_OK : S_FALSE;
}

STDAPI
DllGetClassObject (REFCLSID clsid, REFIID iid, void **ppv)
{
    try {
        IClassFactory *pFactory = module.find(clsid);
        if (pFactory == 0) {
            // Use CLSID to find initialize script from registry.
            std::string clsidStr("{");
            Uuid uuid(clsid);
            clsidStr += uuid.toString();
            clsidStr += "}";

            int completionCode = module.registerFactoryByScript(clsidStr);
            if (completionCode != TCL_OK) {
                *ppv = 0; 
                return E_UNEXPECTED;
            }

            pFactory = module.find(clsid);
        }

        if (pFactory == 0) {
            *ppv = 0; 
            return CLASS_E_CLASSNOTAVAILABLE;
        }
        return pFactory->QueryInterface(iid, ppv);
    }
    catch (...) {
        *ppv = 0; 
        return CLASS_E_CLASSNOTAVAILABLE;
    }
}

BOOL WINAPI
DllMain (
    HINSTANCE hinstDLL, // handle to the DLL module
    DWORD reason,       // reason for calling function
    LPVOID reserved)    // reserved
{
    switch (reason) {
    case DLL_PROCESS_DETACH:
        module.terminate();
        break;
    }

    return TRUE;
}
