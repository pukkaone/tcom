// $Id: DispatchImpl.h 14 2005-04-18 14:14:12Z cthuang $
#ifndef DISPATCHIMPL_H
#define DISPATCHIMPL_H

#include <map>
#include <set>
#include "tcomApi.h"
#include "TypeInfo.h"

class TCOM_API ComObject;

// This class implements an IDispatch interface and delegates the operations to
// the ComObject class.

class TCOM_API DispatchImpl
{
    // delegate operations to this object
    ComObject &m_object;

    // description of the interface to implement
    const Interface &m_interface;

    // dispatch member ID to method description map
    typedef std::map<DISPID, const Method *> DispIdToMethodMap;
    DispIdToMethodMap m_dispIdToMethodMap;

    // dispatch member ID's which are actually properties
    typedef std::set<DISPID> DispIdSet;
    DispIdSet m_propertyDispIds;

    // not implemented
    DispatchImpl(const DispatchImpl &);
    void operator=(const DispatchImpl &);

    // Get dispatch method description.
    const Method *findDispatchMethod(DISPID dispid);

    // Return true if the dispatch member ID identifies a property.
    bool isProperty (DISPID dispid) const
    { return m_propertyDispIds.count(dispid) != 0; }

public:
    DispatchImpl(
        ComObject &object,
        const Interface &interfaceDesc);

    // Get object for delegating operations.
    ComObject &object () const
    { return m_object; }

    // Get ITypeInfo for the interface that is implemented.
    ITypeInfo *typeInfo () const
    { return m_interface.typeInfo(); }

    // Implement IDispatch::Invoke function.
    HRESULT invoke(
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
