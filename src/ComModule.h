// $Id: ComModule.h,v 1.13 2002/04/13 03:53:56 cthuang Exp $
#ifndef COMMODULE_H
#define COMMODULE_H

#include <map>
#include <comdef.h>
#include "tcomApi.h"
#include "Uuid.h"
#include "mutex.h"

class ComObjectFactory;

// This class manages the life cycle of a COM server.

class TCOM_API ComModule
{
    // used to track when the server can exit
    long m_lockCount;

    // This maps a CLSID to a class factory.
    typedef std::map<Uuid, ComObjectFactory *> ClsidToFactoryMap;
    ClsidToFactoryMap m_clsidToFactoryMap;

    // singleton instance
    static ComModule *ms_pInstance;

    // used to synchonize construction of singleton instance
    static Mutex ms_singletonMutex;

    // Do not allow others to create and copy instances of this class.
    ComModule(const ComModule &);       // not implemented
    void operator=(const ComModule &);  // not implemented

protected:
    ComModule ():
        m_lockCount(0)
    { ms_pInstance = this; }

    // Get class object registration flags.
    virtual DWORD regclsFlags() const;

public:
    // Get singleton instance.
    static ComModule &instance();

    // Initialize COM for the current thread.
    virtual void initializeCom(DWORD coinitFlags);

    // Get lock count.
    long lockCount () const
    { return m_lockCount; }

    // Increment lock count.
    virtual void lock();

    // Decrement lock count.
    virtual long unlock();

    // Register a class factory.
    void registerFactory(REFCLSID clsid, ComObjectFactory *pFactory);

    // Search for a class factory by CLSID.
    IClassFactory *find(REFCLSID clsid);

    // Revoke all class factories.
    void revokeFactories();
};

#endif
