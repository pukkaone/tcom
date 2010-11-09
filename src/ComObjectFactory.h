// $Id: ComObjectFactory.h,v 1.11 2002/04/13 03:53:56 cthuang Exp $
#ifndef COMOBJECTFACTORY_H
#define COMOBJECTFACTORY_H

#include "tcomApi.h"
#include "mutex.h"
#include "TclObject.h"
#include "TypeInfo.h"

// This is a factory of COM objects.

class TCOM_API ComObjectFactory: public IClassFactory
{
    // reference count of the factory
    long m_refCount;

    // interfaces to implement
    const Class::Interfaces &m_interfaces;

    // TODO: Directly accessing the Tcl interpreter means the object must run
    // in a single threaded apartment to comply with Tcl's threading rules.

    // Tcl interpreter used to execute Tcl commands
    Tcl_Interp *m_interp;

    // Tcl command executed to create a servant
    TclObject m_constructor;

    // Tcl command executed to destroy servant
    TclObject m_destructor;

    // handle of registered class object
    unsigned long m_classObjectHandle;

    // CLSID used to register active object
    CLSID m_clsid;

    // true if created objects should be registered in running object table
    bool m_registerActiveObject;

    // true if object factory was registered
    bool m_registeredFactory;

    // Execute Tcl script.  Returns Tcl completion code.
    int eval(TclObject script, TclObject *pResult=0);

    // Do not allow others to copy instances of this class.
    ComObjectFactory(const ComObjectFactory &); // not implemented
    void operator=(const ComObjectFactory &);   // not implemented

public:
    ComObjectFactory(
        const Class::Interfaces &interfaces,
        Tcl_Interp *interp,
        TclObject constructor,
        TclObject destructor,
        bool registerActiveObject);
    virtual ~ComObjectFactory();

    // Register factory.
    void registerFactory(REFCLSID clsid, DWORD regclsFlags);

    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();

    // IClassFactory methods
    STDMETHOD(CreateInstance)(IUnknown *pOuter, REFIID riid, void **ppvObj);
    STDMETHOD(LockServer)(BOOL fLock);
};

// This factory always returns the same instance.

class TCOM_API SingletonObjectFactory: public ComObjectFactory
{
    // singleton instance returned from factory
    IUnknown *m_pInstance;

    // used to synchronize construction of singleton instance
    Mutex m_mutex;

public:
    SingletonObjectFactory(
        const Class::Interfaces &interfaces,
        Tcl_Interp *interp,
        TclObject constructor,
        TclObject destructor,
        bool registerActiveObject);
    ~SingletonObjectFactory();

    // Override create function.
    STDMETHOD(CreateInstance)(IUnknown *pOuter, REFIID riid, void **ppvObj);
};

#endif 
