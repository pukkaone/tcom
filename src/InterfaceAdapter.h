// $Id$
#ifndef INTERFACEADAPTER_H
#define INTERFACEADAPTER_H

#include <map>
#include <set>
#include "tcomApi.h"
#include "DispatchImpl.h"

class TCOM_API ComObject;

// This class implements an interface for COM clients to invoke functions
// through a virtual function table.  It delegates the operations to the
// ComObject class.

class TCOM_API InterfaceAdapter
{
    // We rely on the knowledge that the C++ compiler implements objects having
    // virtual functions by storing a pointer to a virtual function table
    // at the beginning of the object.  We simulate such an object by defining
    // a class with a virtual function table pointer as the first data member.
    const void *m_pVtbl;

    // delegate operations to this object
    DispatchImpl m_dispatchImpl;

    // virtual function index to method description map
    typedef std::map<short, const Method *> VtblIndexToMethodMap;
    VtblIndexToMethodMap m_vtblIndexToMethodMap;

    // virtual function table for custom (IUnknown derived) interfaces
    static const void *customVtbl[];

    // virtual function table for dual (IDispatch derived) interfaces
    static const void *dualVtbl[];

     // not implemented
    InterfaceAdapter(const InterfaceAdapter &);
    void operator=(const InterfaceAdapter &);

public:
    InterfaceAdapter(
        ComObject &object,
        const Interface &interfaceDesc,
        bool forceDispatch=false);

    // Get delegate object.
    ComObject &object () const
    { return m_dispatchImpl.object(); }

    // Get COM method description.
    const Method *findComMethod(int funcIndex);

    // IUnknown implementation
    static STDMETHODIMP QueryInterface(
        InterfaceAdapter *pThis, REFIID iid, void **ppvObj);
    static STDMETHODIMP_(ULONG) AddRef(InterfaceAdapter *pThis);
    static STDMETHODIMP_(ULONG) Release(InterfaceAdapter *pThis);

    // IDispatch implementation
    static STDMETHODIMP GetTypeInfoCount(
        InterfaceAdapter *pThis, UINT *pctinfo);
    static STDMETHODIMP GetTypeInfo(
        InterfaceAdapter *pThis, UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
    static STDMETHODIMP GetIDsOfNames(
        InterfaceAdapter *pThis,
        REFIID iid,
        OLECHAR **rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgdispid);
    static STDMETHODIMP Invoke(
        InterfaceAdapter *pThis,
        DISPID dispidMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS *pdispparams,
        VARIANT *pvarResult,
        EXCEPINFO *pexcepinfo,
        UINT *puArgErr);
};

#endif 
