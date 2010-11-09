// $Id$
#pragma warning(disable: 4786)
#include <sstream>
#include "RegistryKey.h"
#include "TypeInfo.h"
#include "TypeLib.h"

Class::Class (const ITypeInfoPtr &pTypeInfo):
    m_name(getTypeInfoName(pTypeInfo)),
    m_pDefaultInterface(0),
    m_pSourceInterface(0)
{
    HRESULT hr;

    TypeAttr typeAttr(pTypeInfo);
    m_clsid = typeAttr->guid;

    // Get interfaces this class implements.
    unsigned interfaceCount = static_cast<unsigned>(typeAttr->cImplTypes);
    for (unsigned i = 0; i < interfaceCount; ++i) {
	HREFTYPE hRefType;
	hr = pTypeInfo->GetRefTypeOfImplType(i, &hRefType);
        if (FAILED(hr)) {
            break;
        }

        ITypeInfoPtr pInterfaceTypeInfo;
	hr = pTypeInfo->GetRefTypeInfo(hRefType, &pInterfaceTypeInfo);
        if (FAILED(hr)) {
            break;
        }
        TypeAttr interfaceTypeAttr(pInterfaceTypeInfo);

        const Interface *pInterface = InterfaceManager::instance().newInterface(
            interfaceTypeAttr->guid, pInterfaceTypeInfo);

        int flags;
        hr = pTypeInfo->GetImplTypeFlags(i, &flags);
        if (FAILED(hr)) {
            break;
        }
        flags &= (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE);

        if (flags == IMPLTYPEFLAG_FDEFAULT) {
            m_pDefaultInterface = pInterface;
        } else if (flags == (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)) {
            m_pSourceInterface = pInterface;
        }

        if ((flags & IMPLTYPEFLAG_FSOURCE) == 0) {
            if (m_pDefaultInterface == 0) {
                m_pDefaultInterface = pInterface;
            }
            m_interfaces.push_back(pInterface);
        }
    }
}

Class::Class (
    const char *name,
    REFCLSID clsid,
    const Interface *pDefaultInterface,
    const Interface *pSourceInterface):
        m_name(name),
        m_clsid(clsid),
        m_pDefaultInterface(pDefaultInterface),
        m_pSourceInterface(pSourceInterface)
{
    m_interfaces.push_back(pDefaultInterface);
}


Enum::Enum (const ITypeInfoPtr &pTypeInfo, TypeAttr &attr):
    m_name(getTypeInfoName(pTypeInfo))
{
   HRESULT hr;

   for (int iEnum = 0; iEnum < attr->cVars; ++iEnum) {
       // Get enumerator description.
       VARDESC *pVarDesc;
       hr = pTypeInfo->GetVarDesc(iEnum, &pVarDesc);
       if (FAILED(hr)) {
           break;
       }

       // Get enumerator name.
       BSTR bstrName;
       UINT namesReturned;
       hr = pTypeInfo->GetNames(pVarDesc->memid, &bstrName, 1, &namesReturned);
       if (SUCCEEDED(hr)) {
           // Remember enumerator name and value.
	   _bstr_t enumNameBstr(bstrName);
	   std::string enumName(enumNameBstr);

	   _variant_t enumValueVar(pVarDesc->lpvarValue);
	   _bstr_t enumValueBstr(enumValueVar);
	   std::string enumValue(enumValueBstr);

           insert(value_type(enumName, enumValue));
       }

       pTypeInfo->ReleaseVarDesc(pVarDesc);
   }
}


TypeLib *
TypeLib::load (const char *name, bool registerTypeLib)
{
    _bstr_t nameStr(name);
    REGKIND regKind = registerTypeLib ? REGKIND_REGISTER : REGKIND_NONE;
    ITypeLibPtr pTypeLib;
    HRESULT hr = LoadTypeLibEx(nameStr, regKind, &pTypeLib);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
    return new TypeLib(pTypeLib);
}

TypeLib *
TypeLib::loadByLibid (const std::string &libidStr, const std::string &version)
{
    // Remove braces so UuidFromString does not complain.
    std::string cleanLibid;
    if (libidStr[0] == '{') {
        cleanLibid = libidStr.substr(1, libidStr.size() - 2);
    } else {
        cleanLibid = libidStr;
    }

    IID libid;
    if (UuidFromString(
     reinterpret_cast<unsigned char *>(const_cast<char *>(cleanLibid.c_str())),
     &libid) != RPC_S_OK) {
        return 0;
    }

    std::string::size_type i = version.find('.');
    std::istringstream majorIn(version.substr(0, i));
    unsigned short majorVersion;
    majorIn >> majorVersion;

    unsigned short minorVersion = 0;
    if (i != std::string::npos) {
        std::istringstream minorIn(version.substr(i + 1));
        minorIn >> minorVersion;
    }

    ITypeLibPtr pTypeLib;
    HRESULT hr = LoadRegTypeLib(
        libid, majorVersion, minorVersion, LOCALE_NEUTRAL, &pTypeLib);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
    return new TypeLib(pTypeLib);
}

TypeLib *
TypeLib::loadByClsid (REFCLSID clsid)
{
    std::string libidStr, version;
    try {
        // Get the LIBID of the type library for the class.
        std::string clsidSubkeyName("CLSID\\{");
        Uuid uuid(clsid);
        clsidSubkeyName.append(uuid.toString());
        clsidSubkeyName.append("}");

        std::string typeLibSubkeyName =
            clsidSubkeyName + "\\TypeLib";
        RegistryKey typeLibKey(HKEY_CLASSES_ROOT, typeLibSubkeyName);
        libidStr = typeLibKey.value();

        std::string versionSubkeyName =
            clsidSubkeyName + "\\Version";
        RegistryKey versionKey(HKEY_CLASSES_ROOT, versionSubkeyName);
        version = versionKey.value();
    }
    catch (std::runtime_error &) {
        return 0;
    }
    return loadByLibid(libidStr, version);
}

TypeLib *
TypeLib::loadByIid (REFIID iid)
{
    std::string libidStr, version;
    try {
        // Get the LIBID of the type library for the interface.
        std::string typeLibSubkeyName("Interface\\{");
        Uuid uuid(iid);
        typeLibSubkeyName.append(uuid.toString());
        typeLibSubkeyName.append("}\\TypeLib");

        RegistryKey typeLibKey(HKEY_CLASSES_ROOT, typeLibSubkeyName);
        libidStr = typeLibKey.value();
        version = typeLibKey.value("Version");
    }
    catch (std::runtime_error &) {
        return 0;
    }
    return loadByLibid(libidStr, version);
}

void
TypeLib::unregister (const char *name)
{
    HRESULT hr;

    ITypeLibPtr pTypeLib;
    _bstr_t nameStr(name);
    hr = LoadTypeLibEx(nameStr, REGKIND_NONE, &pTypeLib);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    TypeLibAttr pLibAttr(pTypeLib);

    hr = UnRegisterTypeLib(
        pLibAttr->guid,
        pLibAttr->wMajorVerNum,
        pLibAttr->wMinorVerNum,
        LANG_NEUTRAL, 
        SYS_WIN32);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
}

std::string
TypeLib::libidString () const
{
    TypeLibAttr pLibAttr(m_pTypeLib);

    Uuid uuid(pLibAttr->guid);
    return uuid.toString();
}

std::string
TypeLib::version () const
{
    TypeLibAttr pLibAttr(m_pTypeLib);

    std::ostringstream out;
    out << pLibAttr->wMajorVerNum << '.' << pLibAttr->wMinorVerNum;
    return out.str();
}

std::string
TypeLib::name () const
{
    BSTR nameStr;
    HRESULT hr = m_pTypeLib->GetDocumentation(
        MEMBERID_NIL, &nameStr, NULL, NULL, NULL);
    if (FAILED(hr)) {
        return std::string();
    }
    _bstr_t wrapper(nameStr, false);
    return std::string(wrapper);
}

std::string
TypeLib::documentation () const
{
    BSTR docStr;
    HRESULT hr = m_pTypeLib->GetDocumentation(
        MEMBERID_NIL, NULL, &docStr, NULL, NULL);
    if (FAILED(hr)) {
        return std::string();
    }
    _bstr_t wrapper(docStr, false);
    return std::string(wrapper);
}

const Interface *
TypeLib::findInterface (const char *name) const
{
    for (Interfaces::const_iterator p = m_interfaces.begin();
     p != m_interfaces.end(); ++p) {
        if ((*p)->name() == name) {
            return *p;
        }
    }
    return 0;
}

const Class *
TypeLib::findClass (const char *name) const
{
    for (Classes::const_iterator p = m_classes.begin();
     p != m_classes.end(); ++p) {
        if (p->name() == name) {
            return &(*p);
        }
    }
    return 0;
}

const Class *
TypeLib::findClass (REFCLSID clsid) const
{
    for (Classes::const_iterator p = m_classes.begin();
     p != m_classes.end(); ++p) {
        if (IsEqualCLSID(p->clsid(), clsid)) {
            return &(*p);
        }
    }
    return 0;
}

const Enum *
TypeLib::findEnum (const char *name) const
{
    for (Enums::const_iterator p = m_enums.begin(); p != m_enums.end(); ++p) {
        if (p->name() == name) {
            return &(*p);
        }
    }
    return 0;
}

void
TypeLib::readTypeLib ()
{
    HRESULT hResult;

    unsigned count = m_pTypeLib->GetTypeInfoCount();
    for (unsigned index = 0; index < count; ++index) {
        ITypeInfoPtr pTypeInfo;
        hResult = m_pTypeLib->GetTypeInfo(index, &pTypeInfo);
        if (FAILED(hResult)) {
            continue;
        }
        TypeAttr typeAttr(pTypeInfo);

        switch (typeAttr->typekind) {
        case TKIND_DISPATCH:
        case TKIND_INTERFACE:
            // Read interface description.
            {
                const Interface *pInterface =
                    InterfaceManager::instance().newInterface(
                        typeAttr->guid, pTypeInfo);
                m_interfaces.push_back(pInterface);
            }
            break;

        case TKIND_COCLASS:
            // Read class description.
            {
                Class aClass(pTypeInfo);
                m_classes.push_back(aClass);
            }
            break;

        case TKIND_ENUM:
            // Read the enumeration values.
            {
                Enum anEnum(pTypeInfo, typeAttr);
                m_enums.push_back(anEnum);
            }
            break;
        }
    }
}
