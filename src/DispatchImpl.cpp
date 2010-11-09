// $Id: DispatchImpl.cpp 14 2005-04-18 14:14:12Z cthuang $
#pragma warning(disable: 4786)
#include "DispatchImpl.h"
#include <stdexcept>
#include "ComObject.h"

DispatchImpl::DispatchImpl (
    ComObject &object,
    const Interface &interfaceDesc):
        m_object(object),
        m_interface(interfaceDesc)
{
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
}

const Method *
DispatchImpl::findDispatchMethod (DISPID dispid)
{
    DispIdToMethodMap::const_iterator p = m_dispIdToMethodMap.find(dispid);
    if (p == m_dispIdToMethodMap.end()) {
        return 0;
    }
    return p->second;
}

HRESULT
DispatchImpl::invoke (
    DISPID dispid,
    REFIID iid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pDispParams,
    VARIANT *pReturnValue,
    EXCEPINFO *pExcepInfo,
    UINT *pArgErr)
{
    // Get the method description for method being invoked.
    const Method *pMethod = findDispatchMethod(dispid);
    if (pMethod == 0) {
        return DISP_E_MEMBERNOTFOUND;
    }

    return m_object.invoke(
        *pMethod,
        isProperty(dispid),
        iid,
        lcid,
        wFlags,
        pDispParams,
        pReturnValue,
        pExcepInfo,
        pArgErr);
}
