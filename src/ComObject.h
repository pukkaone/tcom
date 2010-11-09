// $Id: ComObject.h 13 2005-04-18 12:24:14Z cthuang $
#ifndef COMOBJECT_H
#define COMOBJECT_H

#include <stdarg.h>
#include "tcomApi.h"
#include "HashTable.h"
#include "TclObject.h"
#include "TypeInfo.h"
#include "SupportErrorInfo.h"

class TCOM_API InterfaceAdapter;

// This class represents a COM object.
// The COM object methods are implemented by executing a Tcl command. 

class TCOM_API ComObject
{
    // Implement method invocation through virtual function table call.
    friend void __cdecl invokeComObjectFunction(
        volatile HRESULT hresult,
        volatile DWORD pArgEnd,
        DWORD ebp,
        DWORD funcIndex,
        DWORD retAddr,
        InterfaceAdapter *pThis,
        ...);

    // count of references to this object
    long m_refCount;

    // description of default interface
    const Interface &m_defaultInterface;

    // TODO: Directly accessing the Tcl interpreter means the object must run
    // in a single threaded apartment to comply with Tcl's threading rules.

    // interpreter used to execute Tcl command
    Tcl_Interp *m_interp;

    // Tcl command executed to implement methods
    TclObject m_servant;

    // Tcl command executed when this COM object is destroyed
    TclObject m_destructor;

    // collection of interfaces this object can implement
    typedef HashTable<IID, Interface *> SupportedInterfaceMap;
    SupportedInterfaceMap m_supportedInterfaceMap;

    // collection of implemented interface adapters
    typedef HashTable<IID, void *> IidToAdapterMap;
    IidToAdapterMap m_iidToAdapterMap;

    // implements default interface
    void *m_pDefaultAdapter;

    // implements ISupportErrorInfo
    SupportErrorInfo m_supportErrorInfo;

    // implements IDispatch
    void *m_pDispatch;

    // token returned from RegisterActiveObject
    unsigned long m_activeObjectHandle;

    // true if object registered in running object table
    bool m_registeredActiveObject;

    // true if object is an event sink
    bool m_isSink;

    // Do not allow others to create or copy instances of this class.
    ComObject(
        const Class::Interfaces &interfaces,
        Tcl_Interp *interp,
        TclObject servant,
        TclObject destructor,
        bool isSink);
    ComObject(const ComObject &);       // not implemented
    void operator=(const ComObject &);  // not implemented

    // Create an adapter which implements the specified interface.
    void *implementInterface(const Interface &interfaceDesc);

    // Convert IDispatch argument to Tcl value.
    TclObject getArgument(VARIANT *pArg, const Parameter &param);

    // Convert the native value that the va_list points to into a Tcl value.
    // Returns a va_list pointing to the next argument.
    va_list getArgument(va_list pArg, const Parameter &param, TclObject &dest);

public:
    static ComObject *newInstance(
        const Interface &defaultInterface,
        Tcl_Interp *interp,
        TclObject servant,
        TclObject destructor,
        bool isSink = false);
    static ComObject *newInstance(
        const Class::Interfaces &interfaces,
        Tcl_Interp *interp,
        TclObject servant,
        TclObject destructor);
    ~ComObject();

    // Register object in running object table.
    void registerActiveObject(REFCLSID clsid);

    // Return true if the interface is implemented.
    bool implemented (REFIID iid) const
    { return m_iidToAdapterMap.find(iid) != 0; }
    
    // Get IUnknown pointer to default interface.
    IUnknown *unknown () const
    { return reinterpret_cast<IUnknown *>(m_pDefaultAdapter); }

    // Execute Tcl script.  Returns Tcl completion code.
    int eval(TclObject script, TclObject *pResult=0);

    // Get Tcl variable.  Returns Tcl completion code.
    int getVariable(TclObject name, TclObject &value) const;

    // Set Tcl variable.  Returns Tcl completion code.
    int setVariable(TclObject name, TclObject value);

    // If the first element of the Tcl errorCode variable is "COM", convert
    // second element to an HRESULT.  Return E_UNEXPECTED if errorCode does
    // not contain a recognizable value.
    HRESULT hresultFromErrorCode() const;

    // IUnknown implementation
    HRESULT queryInterface(REFIID riid, void **ppvObj);
    ULONG addRef();
    ULONG release();

    // IDispatch implementation
    HRESULT invoke(
        const Method &method,
        bool isProperty,
        REFIID iid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pDispParams,
        VARIANT *pReturnValue,
        EXCEPINFO *pExcepInfo,
        UINT *pArgErr);
};

#endif 
