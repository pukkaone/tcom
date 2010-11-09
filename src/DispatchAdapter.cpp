// $Id$
#pragma warning(disable: 4786)
#include "DispatchAdapter.h"
#include <stdexcept>
#include "ComObject.h"
#include "Reference.h"
#include "Extension.h"

// Implement IUnknown methods

STDMETHODIMP
DispatchAdapter::QueryInterface (REFIID iid, void **ppvObj)
{
   return m_dispatchImpl.object().queryInterface(iid, ppvObj);
}

STDMETHODIMP_(ULONG)
DispatchAdapter::AddRef ()
{
    return m_dispatchImpl.object().addRef();
}

STDMETHODIMP_(ULONG)
DispatchAdapter::Release ()
{
    return m_dispatchImpl.object().release();
}

// Implement IDispatch methods

STDMETHODIMP
DispatchAdapter::GetTypeInfoCount (UINT *pCount)
{
    *pCount = 1;
    return S_OK;
}

STDMETHODIMP
DispatchAdapter::GetTypeInfo (UINT index, LCID, ITypeInfo **ppTypeInfo)
{
    if (index != 0) {
        *ppTypeInfo = 0;
        return DISP_E_BADINDEX;
    }

    ITypeInfo *pTypeInfo = m_dispatchImpl.typeInfo();
    pTypeInfo->AddRef();
    *ppTypeInfo = pTypeInfo;
    return S_OK;
}

STDMETHODIMP
DispatchAdapter::GetIDsOfNames (
    REFIID,
    OLECHAR **rgszNames,
    UINT cNames,
    LCID,
    DISPID *rgDispId)
{
    ITypeInfo *pTypeInfo = m_dispatchImpl.typeInfo();
    return pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP
DispatchAdapter::Invoke (
    DISPID dispid,
    REFIID iid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pDispParams,
    VARIANT *pReturnValue,
    EXCEPINFO *pExcepInfo,
    UINT *pArgErr)
{
    return m_dispatchImpl.invoke(
        dispid,
        iid,
        lcid,
        wFlags,
        pDispParams,
        pReturnValue,
        pExcepInfo,
        pArgErr);
}
