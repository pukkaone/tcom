// $Id: InterfaceAdapter.h,v 1.3 2002/02/27 01:58:45 cthuang Exp $
#ifndef INTERFACEADAPTER_H
#define INTERFACEADAPTER_H

#include <map>
#include <set>
#include "tcomApi.h"
#include "TypeInfo.h"

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
    ComObject &m_object;

    // description of the interface to implement
    const Interface &m_interface;

    // virtual function index to method description map
    typedef std::map<short, const Method *> VtblIndexToMethodMap;
    VtblIndexToMethodMap m_vtblIndexToMethodMap;

    // dispatch member ID to method description map
    typedef std::map<DISPID, const Method *> DispIdToMethodMap;
    DispIdToMethodMap m_dispIdToMethodMap;

    // dispatch member ID's which are actually properties
    typedef std::set<DISPID> DispIdSet;
    DispIdSet m_propertyDispIds;

    // virtual function table for IUnknown derived interfaces
    static const void *unknownVtbl[];

    // virtual function table for IDispatch derived interfaces
    static const void *dispatchVtbl[];

    InterfaceAdapter(const InterfaceAdapter &); // not implemented
    void operator=(const InterfaceAdapter &);   // not implemented

public:
    InterfaceAdapter(
        ComObject &object,
        const Interface &interfaceDesc,
        bool forceDispatch=false);

    // Get delegate object.
    ComObject &object () const
    { return m_object; }

    // Get COM method description.
    const Method *findComMethod(int funcIndex);

    // Get dispatch method description.
    const Method *findDispatchMethod(DISPID dispid);

    // Return true if the dispatch member ID identifies a property.
    bool isProperty (DISPID dispid) const
    { return m_propertyDispIds.count(dispid) != 0; }

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
