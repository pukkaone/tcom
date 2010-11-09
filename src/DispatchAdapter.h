// $Id$
#ifndef DISPATCHADAPTER_H
#define DISPATCHADAPTER_H

#include "tcomApi.h"
#include "DispatchImpl.h"

// This class implements an IDispatch interface and delegates the operations to
// the ComObject class.

class TCOM_API DispatchAdapter: public IDispatch
{
    // provides IDispatch implementation
    DispatchImpl m_dispatchImpl;

    // not implemented
    DispatchAdapter(const DispatchAdapter &);
    void operator=(const DispatchAdapter &);

public:
    DispatchAdapter (
        ComObject &object,
        const Interface &interfaceDesc):
            m_dispatchImpl(object, interfaceDesc)
    { }

    // IUnknown functions
    STDMETHODIMP QueryInterface(REFIID iid, void **ppvObj);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IDispatch functions
    STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
    STDMETHODIMP GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
    STDMETHODIMP GetIDsOfNames(
        REFIID iid,
        OLECHAR **rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID *rgdispid);
    STDMETHODIMP Invoke(
        DISPID dispidMember,
        REFIID iid,
        LCID lcid,
        WORD flags,
        DISPPARAMS *pParams,
        VARIANT *pResult,
        EXCEPINFO *pExcepInfo,
        UINT *pArgErr);
};

#endif 
