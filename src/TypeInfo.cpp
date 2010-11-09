// $Id$
#pragma warning(disable: 4786)
#include <sstream>
#include <map>
#include "Uuid.h"
#include "TypeInfo.h"

std::string
getTypeInfoName (ITypeInfo *pTypeInfo, MEMBERID id)
{
    BSTR nameStr = 0;
    HRESULT hResult = pTypeInfo->GetDocumentation(
        id, &nameStr, NULL, NULL, NULL);
    if (FAILED(hResult) || nameStr == 0) {
        return std::string();
    }
    _bstr_t wrapper(nameStr, false);
    return std::string(wrapper);
}


struct VarTypeStringAssoc {
    VARTYPE vt;
    const char *name;
};

static VarTypeStringAssoc varTypeStringAssocs[] = {
    { VT_EMPTY, "EMPTY" },
    { VT_NULL, "NULL" },
    { VT_I2, "I2" },
    { VT_I4, "I4" },
    { VT_R4, "R4" },
    { VT_R8, "R8" },
    { VT_CY, "CY" },
    { VT_DATE, "DATE" },
    { VT_BSTR, "BSTR" },
    { VT_DISPATCH, "DISPATCH" },
    { VT_ERROR, "SCODE" },
    { VT_BOOL, "BOOL" },
    { VT_VARIANT, "VARIANT" },
    { VT_UNKNOWN, "UNKNOWN" },
    { VT_DECIMAL, "DECIMAL" },
    { VT_RECORD, "RECORD" },
    { VT_I1, "I1" },
    { VT_UI1, "UI1" },
    { VT_UI2, "UI2" },
    { VT_UI4, "UI4" },
    { VT_INT, "INT" },
    { VT_UINT, "UINT" },
    { VT_VOID, "VOID" },
    { VT_LPSTR, "LPSTR" },
    { VT_LPWSTR, "LPWSTR" },
    { VT_ARRAY, "ARRAY" },
};

// This class maps from a VARTYPE to a string representation.

class VarTypeToStringMap: public std::map<VARTYPE, std::string>
{
public:
    VarTypeToStringMap();
};

VarTypeToStringMap::VarTypeToStringMap ()
{
    const int n = sizeof(varTypeStringAssocs) / sizeof(VarTypeStringAssoc);
    for (int i = 0; i < n; ++i) {
        const VarTypeStringAssoc &assoc = varTypeStringAssocs[i];
        insert(value_type(assoc.vt, assoc.name));
    }
}

static VarTypeToStringMap varTypeToStringMap;

// This class maps from a string representation to a VARTYPE.

class StringToVarTypeMap: public std::map<std::string, VARTYPE>
{
public:
    StringToVarTypeMap();
};

StringToVarTypeMap::StringToVarTypeMap ()
{
    const int n = sizeof(varTypeStringAssocs) / sizeof(VarTypeStringAssoc);
    for (int i = 0; i < n; ++i) {
        const VarTypeStringAssoc &assoc = varTypeStringAssocs[i];
        insert(value_type(assoc.name, assoc.vt));
    }
}

static StringToVarTypeMap stringToVarTypeMap;

Type Type::ms_variant("VARIANT");

Type::Type (const ITypeInfoPtr &pTypeInfo, TYPEDESC &typeDesc):
    m_pointerCount(0),
    m_pElementType(0)
{
    UuidCreateNil(&m_iid);
    readTypeDesc(pTypeInfo, typeDesc);
}

Type::Type (const std::string &str):
    m_pointerCount(0),
    m_pElementType(0)
{
    UuidCreateNil(&m_iid);

    std::istringstream in(str);
    std::string token;
    while (in >> token) {
        if (token[0] == '*') {
            ++m_pointerCount;
        } else {
            StringToVarTypeMap::iterator p = stringToVarTypeMap.find(token);
            if (p == stringToVarTypeMap.end()) {
                m_vt = VT_USERDEFINED;
                m_name = token;
            } else {
                m_vt = p->second;
            }
        }
    }
}

Type::Type (const Type &rhs):
    m_name(rhs.m_name),
    m_vt(rhs.m_vt),
    m_iid(rhs.m_iid),
    m_pointerCount(rhs.m_pointerCount),
    m_pElementType(rhs.m_pElementType ? new Type(*rhs.m_pElementType) : 0)
{ }

Type::~Type ()
{
    delete m_pElementType;
}

Type &
Type::operator= (const Type &rhs)
{
    m_name = rhs.m_name;
    m_vt = rhs.m_vt;
    m_iid = rhs.m_iid;
    m_pointerCount = rhs.m_pointerCount;

    delete m_pElementType;
    m_pElementType = rhs.m_pElementType ? new Type(*rhs.m_pElementType) : 0;

    return *this;
}

std::string
Type::toString () const
{
    switch (m_vt) {
    case VT_USERDEFINED:
        return m_name;

    case VT_SAFEARRAY:
        {
            std::ostringstream out;
            out << "SAFEARRAY(" << elementType().toString() << ")";
            return out.str();
        }

    default:
        {
            VarTypeToStringMap::iterator p = varTypeToStringMap.find(m_vt);
            if (p != varTypeToStringMap.end()) {
                return p->second;
            }

            std::ostringstream out;
            out << "vartype0x" << std::hex << m_vt;
            return out.str();
        }
    }
}

void
Type::readTypeDesc (const ITypeInfoPtr &pTypeInfo, TYPEDESC &typeDesc)
{
    HRESULT hr;

    switch (typeDesc.vt) {
    case VT_USERDEFINED:
        {
	    // It's an alias.  Expand the alias.
            ITypeInfoPtr pRefTypeInfo;
	    hr = pTypeInfo->GetRefTypeInfo(typeDesc.hreftype, &pRefTypeInfo);
            if (SUCCEEDED(hr)) {
	        if (m_name.empty()) {
                    m_name = getTypeInfoName(pRefTypeInfo);
	        }

                TypeAttr typeAttr(pRefTypeInfo);
                if (typeAttr->typekind == TKIND_ALIAS) {
		    // Type expanded to another alias!
		    readTypeDesc(pRefTypeInfo, typeAttr->tdescAlias);
                } else if (typeAttr->typekind == TKIND_ENUM) {
        	    m_vt = VT_I4;
                } else {
        	    m_vt = typeDesc.vt;
                    m_iid = typeAttr->guid;
                }
	    }
        }
	break;

    case VT_PTR:
	// It's a pointer.  Dereference and try to interpret with one less
	// level of indirection.
	++m_pointerCount;
        readTypeDesc(pTypeInfo, *typeDesc.lptdesc);
	break;

    case VT_SAFEARRAY:
	// It's a SAFEARRAY.  Get the element type.
        m_pElementType = new Type(pTypeInfo, *typeDesc.lptdesc);
	m_vt = typeDesc.vt;
	break;

    default:
	m_vt = typeDesc.vt;
    }
}

Parameter::Parameter (const ITypeInfoPtr &pTypeInfo,
                      ELEMDESC &elemDesc,
                      const char *name):
    m_flags(elemDesc.paramdesc.wParamFlags),
    m_type(pTypeInfo, elemDesc.tdesc),
    m_name(name)
{
    if (m_flags == 0) {
        // No parameter passing flags are set.  Assume it's an in parameter.
        m_flags = PARAMFLAG_FIN;
    }
}

Parameter::Parameter (const std::string &modes,
                      const std::string &type,
                      const std::string &name):
    m_flags(0),
    m_type(type),
    m_name(name)
{
    std::istringstream in(modes);
    std::string token;
    while (in >> token) {
        if (token == "in") {
            m_flags |= PARAMFLAG_FIN;
        } else if (token == "out") {
            m_flags |= PARAMFLAG_FOUT;
        }
    }
}


Method::Method (const ITypeInfoPtr &pTypeInfo, DISPID memberid, ELEMDESC &elemDesc):
    m_memberid(memberid),
    m_type(pTypeInfo, elemDesc.tdesc)
{
    // Get name.
    BSTR nameBstr;
    unsigned numNames;
    HRESULT hr = pTypeInfo->GetNames(
        memberid,
        &nameBstr,
        1,
        &numNames);
    if (FAILED(hr)) {
        // TODO: I should throw an exception here.
        return;
    }

    // Initialize name.
    _bstr_t name(nameBstr, false);
    m_name = std::string(name);
}

Method::Method (const ITypeInfoPtr &pTypeInfo, FuncDesc &funcDesc):
    m_memberid(funcDesc->memid),
    m_type(pTypeInfo, funcDesc->elemdescFunc.tdesc),
    m_invokeKind(funcDesc->invkind),
    m_vtblIndex(funcDesc->oVft / 4),
    m_vararg(funcDesc->cParamsOpt == -1)
{
    // Get method and parameter names.
    BSTR *pNames = new BSTR[funcDesc->cParams + 1];
    unsigned numNames;
    HRESULT hr = pTypeInfo->GetNames(
        funcDesc->memid,
        pNames,
        funcDesc->cParams + 1,
        &numNames);
    if (FAILED(hr)) {
        // TODO: I should throw an exception here.
        delete[] pNames;
        return;
    }

    // Initialize method name.
    _bstr_t methodName(pNames[0], false);
    m_name = std::string(methodName);

    // Get parameter types.
    for (unsigned i = 1; i < numNames; ++i) {
        _bstr_t paramName(pNames[i], false);
        Parameter parameter(
            pTypeInfo,
            funcDesc->lprgelemdescParam[i - 1],
            paramName);
        addParameter(parameter);
    }

    // Some TKIND_INTERFACE methods specify a parameter is a return value
    // with the PARAMFLAG_FRETVAL flag.  Check if the last parameter is
    // actually a return value.
    if (m_type.vartype() == VT_HRESULT) {
        m_type = Type("VOID");

        if (m_parameters.size() > 0) {
            Parameter last = m_parameters.back();
            if ((last.flags() & (PARAMFLAG_FOUT|PARAMFLAG_FRETVAL))
             == (PARAMFLAG_FOUT|PARAMFLAG_FRETVAL)) {
                m_type = last.type();
                m_parameters.pop_back();
            }
        }
    }

    delete[] pNames;
}

Method::Method (MEMBERID memberid,
                const std::string &type,
                const std::string &name,
                INVOKEKIND invokeKind):
    m_memberid(memberid),
    m_type(type),
    m_name(name),
    m_invokeKind(invokeKind)
{ }


Property::Property (const ITypeInfoPtr &pTypeInfo, FuncDesc &funcDesc):
    Method(pTypeInfo, funcDesc),
    m_readOnly(true)
{
    initialize();
}

Property::Property (const ITypeInfoPtr &pTypeInfo, VarDesc &varDesc):
    Method(pTypeInfo, varDesc->memid, varDesc->elemdescVar),
    m_readOnly((varDesc->wVarFlags & VARFLAG_FREADONLY) != 0)
{
    initialize();
}

Property::Property (MEMBERID memberid,
                    const std::string &modes,
                    const std::string &type,
                    const std::string &name):
    Method(memberid, type, name, INVOKE_PROPERTYGET),
    m_readOnly(true)
{
    // Initialize readable/writable flags.
    std::istringstream in(modes);
    std::string token;
    while (in >> token) {
        if (token == "in") {
            m_readOnly = false;
        }
    }
}

void
Property::initialize ()
{
    switch (invokeKind()) {
    case INVOKE_PROPERTYPUT:
        m_putDispatchFlag = DISPATCH_PROPERTYPUT;
        break;
    case INVOKE_PROPERTYPUTREF:
        m_putDispatchFlag = DISPATCH_PROPERTYPUTREF;
        break;
    }
}

void
Property::merge (const Property &property)
{
    // Set dispatch flag used for property put.
    switch (property.invokeKind()) {
    case INVOKE_PROPERTYPUT:
        m_putDispatchFlag = DISPATCH_PROPERTYPUT;
        break;
    case INVOKE_PROPERTYPUTREF:
        m_putDispatchFlag = DISPATCH_PROPERTYPUTREF;
        break;
    }

    // Set read only flag.
    if (property.invokeKind() != INVOKE_PROPERTYGET) {
        m_readOnly = false;
    }

    // Set property value type.
    if (property.type().vartype() != VT_VOID) {
        type(property.type());
    }
}


std::string
Interface::iidString () const
{
    Uuid uuid(m_iid);
    return uuid.toString();
}

void
Interface::addMethod (const Method &method)
{
    // Do we already have information on this method?
    for (Methods::iterator p = m_methods.begin(); p != m_methods.end(); ++p) {
        if (p->memberid() == method.memberid()
         && p->vtblIndex() == method.vtblIndex()) {
            return;
        }
    }

    // Add method description.
    m_methods.push_back(method);
}

void
Interface::addProperty (const Property &property)
{
    // Do we already have information on this property?
    for (Properties::iterator p = m_properties.begin();
     p != m_properties.end(); ++p) {
        if (p->memberid() == property.memberid()) {
            p->merge(property);
            return;
        }
    }

    // Add property description.
    m_properties.push_back(property);
}

void
Interface::readFunctions (const ITypeInfoPtr &pTypeInfo, TypeAttr &typeAttr)
{
    HRESULT hr;

    // Don't expose the IUnknown and IDispatch functions to the Tcl script
    // because of potential dangers.
    if (IsEqualIID(typeAttr->guid, IID_IUnknown)
     || IsEqualIID(typeAttr->guid, IID_IDispatch)) {
        return;
    }

    // Get properties and methods from inherited interfaces.
    for (int i = 0; i < typeAttr->cImplTypes; ++i) {
	HREFTYPE hRefType;
	hr = pTypeInfo->GetRefTypeOfImplType(i, &hRefType);
        if (FAILED(hr)) {
            break;
        }

        ITypeInfoPtr pSuperTypeInfo;
	hr = pTypeInfo->GetRefTypeInfo(hRefType, &pSuperTypeInfo);
        if (FAILED(hr)) {
            break;
        }
        TypeAttr superTypeAttr(pSuperTypeInfo);
        readFunctions(pSuperTypeInfo, superTypeAttr);
    }

    bool dual = (typeAttr->wTypeFlags & TYPEFLAG_FDUAL) != 0;

    // Get properties and methods of this interface.
    for (int j = 0; j < typeAttr->cFuncs; ++j) {
        FuncDesc funcDesc(pTypeInfo, j);

        if (dual && funcDesc->funckind == FUNC_DISPATCH
         && funcDesc->oVft < 28) {
            // Don't expose the IUnknown and IDispatch functions to the Tcl
            // script because of potential dangers.
            continue;
        }

        Method method(pTypeInfo, funcDesc);
        addMethod(method);

	if (funcDesc->invkind != INVOKE_FUNC) {
	    // This is a property get/set function.
            Property property(pTypeInfo, funcDesc);
            addProperty(property);
	}
    }

    // Some objects expose their properties as variable members.
    for (int k = 0; k < typeAttr->cVars; ++k) {
        VarDesc varDesc(pTypeInfo, k);

	if (varDesc->varkind == VAR_DISPATCH) {
            Property property(pTypeInfo, varDesc);
            addProperty(property);
	}
    }

    // Add missing parameter description to property put functions.
    for (Methods::iterator pMethod = m_methods.begin();
     pMethod != m_methods.end(); ++pMethod) {
        if (pMethod->invokeKind() == INVOKE_PROPERTYPUT
         || pMethod->invokeKind() == INVOKE_PROPERTYPUTREF) {
            const Property *pProperty = findProperty(pMethod->name().c_str());
            if (pProperty != 0) {
                Parameter parameter(
                    PARAMFLAG_FIN,
                    pProperty->type(),
                    "propertyValue");
                pMethod->addParameter(parameter);
            }
        }
    }
}

void
Interface::readTypeInfo (const ITypeInfoPtr &pTypeInfo)
{
    m_name = getTypeInfoName(pTypeInfo);

    TypeAttr typeAttr(pTypeInfo);
    m_iid = typeAttr->guid;
    m_dispatchOnly = (typeAttr->typekind == TKIND_DISPATCH) &&
        ((typeAttr->wTypeFlags & TYPEFLAG_FDUAL) == 0);
    m_dispatchable = (typeAttr->wTypeFlags & TYPEFLAG_FDISPATCHABLE) != 0;

    m_methods.reserve(typeAttr->cFuncs);
    m_properties.reserve(typeAttr->cFuncs);
    readFunctions(pTypeInfo, typeAttr);
}

const Method *
Interface::findMethod (const char *name) const
{
    for (Methods::const_iterator p = m_methods.begin();
     p != m_methods.end(); ++p) {
        if (p->name() == name) {
            return &(*p);
        }
    }
    return 0;
}

const Property *
Interface::findProperty (const char *name) const
{
    for (Properties::const_iterator p = m_properties.begin();
     p != m_properties.end(); ++p) {
        if (p->name() == name) {
            return &(*p);
        }
    }
    return 0;
}


Singleton<InterfaceManager> InterfaceManager::ms_singleton;

InterfaceManager &
InterfaceManager::instance ()
{
    return ms_singleton.instance();
}

InterfaceManager::InterfaceManager ()
{
}

InterfaceManager::~InterfaceManager ()
{
    // Delete cached interface descriptions.
    m_hashTable.forEach(Delete());
}

const Interface *
InterfaceManager::newInterface (REFIID iid, const ITypeInfoPtr &pTypeInfo)
{
    LOCK_MUTEX(m_mutex)

    Interface *pInterface = m_hashTable.find(iid);
    if (pInterface == 0) {
        pInterface = new Interface(pTypeInfo);
        m_hashTable.insert(iid, pInterface);
    }
    return pInterface;
}

Interface *
InterfaceManager::newInterface (REFIID iid, const char *name)
{
    LOCK_MUTEX(m_mutex)

    Interface *pInterface = m_hashTable.find(iid);
    if (pInterface == 0) {
        pInterface = new Interface(iid, name);
        m_hashTable.insert(iid, pInterface);
    }
    return pInterface;
}

const Interface *
InterfaceManager::find (REFIID iid) const
{
    LOCK_MUTEX(m_mutex)

    return m_hashTable.find(iid);
}


FuncDesc::FuncDesc (const ITypeInfoPtr &pTypeInfo, unsigned index):
    m_pTypeInfo(pTypeInfo)
{
    HRESULT hr = m_pTypeInfo->GetFuncDesc(index, &m_pFuncDesc);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
}

VarDesc::VarDesc (const ITypeInfoPtr &pTypeInfo, unsigned index):
    m_pTypeInfo(pTypeInfo)
{
    HRESULT hr = m_pTypeInfo->GetVarDesc(index, &m_pVarDesc);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
}
