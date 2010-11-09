// $Id: Reference.h,v 1.42 2003/11/06 15:29:01 cthuang Exp $
#ifndef REFERENCE_H
#define REFERENCE_H

#include <vector>
#include "tcomApi.h"
#include "TclObject.h"
#include "TypeInfo.h"

class TypedArguments;

// Throw this exception when invoke returns DISP_E_EXCEPTION.

class DispatchException
{
    SCODE m_scode;
    _bstr_t m_description;

public:
    DispatchException (SCODE scode, const _bstr_t &description):
        m_scode(scode),
        m_description(description)
    { }

    SCODE scode () const
    { return m_scode; }

    const _bstr_t &description () const
    { return m_description; }
};

// This class holds an interface pointer and the interface description needed
// to invoke methods on it.

class TCOM_API Reference
{
    // This represents a connection from a connection point to an event sink.

    class TCOM_API Connection
    {
        // pointer to connection point
        IConnectionPoint *m_pConnectionPoint;

        // cookie returned from Advise
        DWORD m_adviseCookie;

    public:
        // Create an event sink object and connect it to the connection point.
        Connection(
            Tcl_Interp *interp,
            IUnknown *pSource,
            const Interface &eventInterfaceDesc,
            TclObject servant);

        // Disconnect from the connection point and release the pointer to the
        // connection point.
        ~Connection();
    };

    // collection of event connections
    typedef std::vector<Connection *> Connections;
    Connections m_connections;

    // interface pointer to the object
    IUnknown *m_pUnknown;

    // this pointer is non-null if the object implements IDispatch
    IDispatch *m_pDispatch;

    // interface description includes information about methods and properties
    const Interface *m_pInterface;

    // class description includes information about interfaces exposed
    Class *m_pClass;

    // CLSID of the class the COM object implements
    CLSID m_clsid;
    
    // true if we know the CLSID of the class the COM object implements
    bool m_haveClsid;

    // The constructor assumes the reference count on the interface pointer has
    // already been incremented.  This object will decrement the reference
    // count when it is destroyed.
    Reference(IUnknown *pUnknown, const Interface *pInterface);
    Reference(
        IUnknown *pUnknown, const Interface *pInterface, REFCLSID clsid);

    // Do not allow instances of this class to be copied.
    Reference(const Reference &rhs);
    Reference &operator=(const Reference &rhs);

    // Try to get interface description from IDispatch object.
    static const Interface *findInterfaceFromDispatch(IUnknown *pUnknown);

    // Try to get interface description from type library specified by CLSID.
    static const Interface *findInterfaceFromClsid(REFCLSID clsid);

    // Try to get interface description from type library specified by IID.
    static const Interface *findInterfaceFromIid(REFIID iid);

    // Get description of interface implemented by the object.
    static const Interface *findInterface(IUnknown *pUnknown, REFCLSID clsid);

public:
    // destructor
    ~Reference();

    // Perform a QueryInterface on the interface pointer and create a reference.
    static Reference *newReference(
        IUnknown *pUnknown, const Interface *pInterface=0);

    // Perform a QueryInterface on the interface pointer and create a reference.
    static Reference *queryInterface(IUnknown *pUnknown, REFIID iid);

    // Create an object using CoCreateInstance and construct a reference.
    static Reference *createInstance(
        REFCLSID clsid,
        const Interface *pInterface,
        DWORD clsCtx,
        const char *serverHost);

    // Create an object using CoCreateInstance and construct a reference.
    static Reference *createInstance(
        const char *progId, DWORD clsCtx, const char *serverHost);

    // Get an object using GetActiveObject and construct a reference.
    static Reference *getActiveObject(
        REFCLSID clsid, const Interface *pInterface);

    // Get an object using GetActiveObject and construct a reference.
    static Reference *getActiveObject(const char *progId);

    // Get an object using CoGetObject and construct a reference.
    static Reference *getObject(const wchar_t *displayName);

    // Get raw interface pointer.
    IUnknown *unknown () const
    { return m_pUnknown; }

    // If the object implements IDispatch, return an IDispatch pointer,
    // else return 0.
    IDispatch *dispatch();

    // Get interface description.
    const Interface *interfaceDesc () const
    { return m_pInterface; }

    // Get class description.
    const Class *classDesc();

    // Invoke a method or property using IDispatch.
    HRESULT invokeDispatch(
        MEMBERID memberid,
        WORD dispatchFlags,
        const TypedArguments &arguments,
        VARIANT *pResult);

    // Invoke a method or property.
    HRESULT invoke(
        MEMBERID memberid,
        WORD dispatchFlags,
        const TypedArguments &arguments,
        VARIANT *pResult);

    // Create an event sink object and connect it to the connection point.
    void advise(
        Tcl_Interp *interp,
        const Interface &eventInterfaceDesc,
        TclObject servant);

    // Disconnect all connected event sink objects.
    void unadvise();

    // Compare for COM identity.
    bool operator==(const Reference &rhs) const;
};

#endif 
