// $Id: ComModule.cpp 5 2005-02-16 14:57:24Z cthuang $
#pragma warning(disable: 4786)
#include "ComObjectFactory.h"
#include "ComModule.h"

// This is the default module for event sink objects.

class DefaultModule: public ComModule
{
public:
    DefaultModule ()
    { }

    ~DefaultModule ()
    { revokeFactories(); }
};


ComModule *ComModule::ms_pInstance;

Mutex ComModule::ms_singletonMutex;

ComModule &
ComModule::instance ()
{
    if (ms_pInstance == 0) {
        LOCK_MUTEX(ms_singletonMutex)
        static DefaultModule module;
    }
    return *ms_pInstance;
}

// This exit handler uninitializes COM.

static void
exitProc (ClientData)
{
    CoUninitialize();
}

void
ComModule::initializeCom (DWORD coinitFlags)
{
#ifdef _WIN32_DCOM
    CoInitializeEx(NULL, coinitFlags);
#else
    CoInitialize(NULL);
#endif

#ifdef TCL_THREADS
    Tcl_CreateThreadExitHandler(exitProc, 0);
#else
    Tcl_CreateExitHandler(exitProc, 0);
#endif
}

DWORD
ComModule::regclsFlags () const
{
    return REGCLS_MULTIPLEUSE;
}

void
ComModule::lock ()
{
    InterlockedIncrement(&m_lockCount);
}

long
ComModule::unlock ()
{
    InterlockedDecrement(&m_lockCount);
    return m_lockCount;
}

void
ComModule::registerFactory (REFCLSID clsid,
                            ComObjectFactory *pFactory)
{
    pFactory->registerFactory(clsid, regclsFlags());

    Uuid classId(clsid);
    m_clsidToFactoryMap.insert(ClsidToFactoryMap::value_type(
        classId, pFactory));
    pFactory->AddRef();
}

IClassFactory *
ComModule::find (REFCLSID clsid)
{
    Uuid classId(clsid);
    ClsidToFactoryMap::iterator p = m_clsidToFactoryMap.find(classId);
    if (p != m_clsidToFactoryMap.end()) {
        return p->second;
    }
    return 0;
}

void
ComModule::revokeFactories ()
{
    ClsidToFactoryMap::iterator p = m_clsidToFactoryMap.begin();
    for (; p != m_clsidToFactoryMap.end(); ++p) {
        p->second->Release();
    }

    m_clsidToFactoryMap.clear();
}
