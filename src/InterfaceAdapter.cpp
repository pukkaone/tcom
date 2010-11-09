// $Id: InterfaceAdapter.cpp 16 2005-04-19 14:47:52Z cthuang $
#ifdef TCOM_VTBL_SERVER

#pragma warning(disable: 4786)
#include "InterfaceAdapter.h"
#include "ComObject.h"

InterfaceAdapter::InterfaceAdapter (
    ComObject &object,
    const Interface &interfaceDesc,
    bool forceDispatch):
        m_dispatchImpl(object, interfaceDesc)
{
    // Initialize virtual function index to method description map.
    const Interface::Methods &methods = interfaceDesc.methods();
    for (Interface::Methods::const_iterator p = methods.begin();
     p != methods.end(); ++p) {
        m_vtblIndexToMethodMap.insert(VtblIndexToMethodMap::value_type(
            p->vtblIndex(), &(*p)));
    }

    if (interfaceDesc.dispatchable() || forceDispatch) {
        m_pVtbl = dualVtbl;
    } else {
        m_pVtbl = customVtbl;
    }
}

const Method *
InterfaceAdapter::findComMethod (int funcIndex)
{
    VtblIndexToMethodMap::const_iterator p =
        m_vtblIndexToMethodMap.find(funcIndex);
    if (p == m_vtblIndexToMethodMap.end()) {
        return 0;
    }
    return p->second;
}

// Implement IUnknown methods

STDMETHODIMP
InterfaceAdapter::QueryInterface (
    InterfaceAdapter *pThis, REFIID iid, void **ppvObj)
{
   return pThis->m_dispatchImpl.object().queryInterface(iid, ppvObj);
}

STDMETHODIMP_(ULONG)
InterfaceAdapter::AddRef (InterfaceAdapter *pThis)
{
    return pThis->m_dispatchImpl.object().addRef();
}

STDMETHODIMP_(ULONG)
InterfaceAdapter::Release (InterfaceAdapter *pThis)
{
    return pThis->m_dispatchImpl.object().release();
}

// Implement IDispatch methods

STDMETHODIMP
InterfaceAdapter::GetTypeInfoCount (InterfaceAdapter *, UINT *pCount)
{
    *pCount = 1;
    return S_OK;
}

STDMETHODIMP
InterfaceAdapter::GetTypeInfo (
    InterfaceAdapter *pThis, UINT index, LCID, ITypeInfo **ppTypeInfo)
{
    if (index != 0) {
        *ppTypeInfo = 0;
        return DISP_E_BADINDEX;
    }

    ITypeInfo *pTypeInfo = pThis->m_dispatchImpl.typeInfo();
    pTypeInfo->AddRef();
    *ppTypeInfo = pTypeInfo;
    return S_OK;
}

STDMETHODIMP
InterfaceAdapter::GetIDsOfNames (
    InterfaceAdapter *pThis,
    REFIID,
    OLECHAR **rgszNames,
    UINT cNames,
    LCID,
    DISPID *rgDispId)
{
    ITypeInfo *pTypeInfo = pThis->m_dispatchImpl.typeInfo();
    return pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP
InterfaceAdapter::Invoke (
    InterfaceAdapter *pThis,
    DISPID dispid,
    REFIID iid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pDispParams,
    VARIANT *pVarResult,
    EXCEPINFO *pExcepInfo,
    UINT *pArgErr)
{
    return pThis->m_dispatchImpl.invoke(
        dispid,
        iid,
        lcid,
        wFlags,
        pDispParams,
        pVarResult,
        pExcepInfo,
        pArgErr);
}

#endif
