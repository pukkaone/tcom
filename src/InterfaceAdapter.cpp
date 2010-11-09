// $Id: InterfaceAdapter.cpp,v 1.3 2002/02/27 01:58:45 cthuang Exp $
#pragma warning(disable: 4786)
#include "ComObject.h"
#include "InterfaceAdapter.h"

InterfaceAdapter::InterfaceAdapter (
    ComObject &object,
    const Interface &interfaceDesc,
    bool forceDispatch):
        m_object(object),
        m_interface(interfaceDesc)
{
    // Initialize virtual function index to method description map.
    const Interface::Methods &methods = m_interface.methods();
    for (Interface::Methods::const_iterator p = methods.begin();
     p != methods.end(); ++p) {
        m_vtblIndexToMethodMap.insert(VtblIndexToMethodMap::value_type(
            p->vtblIndex(), &(*p)));
    }

    if (m_interface.dispatchable() || forceDispatch) {
        m_pVtbl = dispatchVtbl;

        // Initialize dispatch member ID to method description map.
        const Interface::Methods &methods = m_interface.methods();
        for (Interface::Methods::const_iterator pMethod = methods.begin();
         pMethod != methods.end(); ++pMethod) {
            m_dispIdToMethodMap.insert(DispIdToMethodMap::value_type(
                pMethod->memberid(), &(*pMethod)));
        }

        // Initialize set of property dispatch member ID's.
        const Interface::Properties &properties = m_interface.properties();
        for (Interface::Properties::const_iterator pProp = properties.begin();
         pProp != properties.end(); ++pProp) {
            m_propertyDispIds.insert(pProp->memberid());
        }

    } else {
        m_pVtbl = unknownVtbl;
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

const Method *
InterfaceAdapter::findDispatchMethod (DISPID dispid)
{
    DispIdToMethodMap::const_iterator p = m_dispIdToMethodMap.find(dispid);
    if (p == m_dispIdToMethodMap.end()) {
        return 0;
    }
    return p->second;
}

// Implement IUnknown methods

STDMETHODIMP
InterfaceAdapter::QueryInterface (
    InterfaceAdapter *pThis, REFIID iid, void **ppvObj)
{
   return pThis->m_object.queryInterface(iid, ppvObj);
}

STDMETHODIMP_(ULONG)
InterfaceAdapter::AddRef (InterfaceAdapter *pThis)
{
    return pThis->m_object.addRef();
}

STDMETHODIMP_(ULONG)
InterfaceAdapter::Release (InterfaceAdapter *pThis)
{
    return pThis->m_object.release();
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

    ITypeInfo *pTypeInfo = pThis->m_interface.typeInfo();
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
    ITypeInfo *pTypeInfo = pThis->m_interface.typeInfo();
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
    return pThis->m_object.invoke(
        pThis,
        dispid,
        iid,
        lcid,
        wFlags,
        pDispParams,
        pVarResult,
        pExcepInfo,
        pArgErr);
}
