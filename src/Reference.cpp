// $Id$
#pragma warning(disable: 4786)
#include <string.h>
#include "ComObject.h"
#include "TypeLib.h"
#include "Uuid.h"
#include "Arguments.h"
#include "Reference.h"

Reference::Connection::Connection (Tcl_Interp *interp,
                                   IUnknown *pSource,
                                   const Interface &eventInterfaceDesc,
                                   TclObject servant)
{
    HRESULT hr;

    // Get connection point container.
    IConnectionPointContainerPtr pContainer;
    hr = pSource->QueryInterface(
        IID_IConnectionPointContainer,
        reinterpret_cast<void **>(&pContainer));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Find connection point.
    hr = pContainer->FindConnectionPoint(
        eventInterfaceDesc.iid(), &m_pConnectionPoint);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Create event sink.
    ComObject *pComObject = ComObject::newInstance(
        eventInterfaceDesc,
        interp,
        servant,
        "",
        true);

    // Connect to connection point.
    hr = m_pConnectionPoint->Advise(pComObject->unknown(), &m_adviseCookie);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
}

Reference::Connection::~Connection ()
{
    m_pConnectionPoint->Unadvise(m_adviseCookie);
    m_pConnectionPoint->Release();
}

Reference::Reference (IUnknown *pUnknown, const Interface *pInterface):
    m_pUnknown(pUnknown),
    m_pDispatch(0),
    m_pInterface(pInterface),
    m_pClass(0),
    m_haveClsid(false)
{ }

Reference::Reference (
        IUnknown *pUnknown, const Interface *pInterface, REFCLSID clsid):
    m_pUnknown(pUnknown),
    m_pDispatch(0),
    m_pInterface(pInterface),
    m_pClass(0),
    m_clsid(clsid),
    m_haveClsid(true)
{ }

Reference::~Reference()
{
    unadvise();
    if (m_pDispatch != 0) {
        m_pDispatch->Release();
    }
    m_pUnknown->Release();
    delete m_pClass;
}

IDispatch *
Reference::dispatch ()
{
    if (m_pDispatch == 0) {
        HRESULT hr = m_pUnknown->QueryInterface(
            IID_IDispatch, reinterpret_cast<void **>(&m_pDispatch));
        if (FAILED(hr)) {
            m_pDispatch = 0;
        }
    }
    return m_pDispatch;
}

const Class *
Reference::classDesc ()
{
    if (!m_haveClsid) {
        return 0;
    }

    if (m_pClass == 0) {
        TypeLib *pTypeLib = TypeLib::loadByClsid(m_clsid);
        if (pTypeLib != 0) {
            const Class *pClass = pTypeLib->findClass(m_clsid);
            if (pClass != 0) {
                m_pClass = new Class(*pClass);
            }
        }
        delete pTypeLib;
    }

    return m_pClass;
}

void
Reference::advise (Tcl_Interp *interp,
                   const Interface &eventInterfaceDesc,
                   TclObject servant)
{
    m_connections.push_back(new Connection(
        interp, m_pUnknown, eventInterfaceDesc, servant));
}

void
Reference::unadvise ()
{
    for (Connections::iterator p = m_connections.begin();
     p != m_connections.end(); ++p) {
        delete *p;
    }
    m_connections.clear();
}

bool
Reference::operator== (const Reference &rhs) const
{
    HRESULT hr;

    IUnknown *pUnknown1;
    hr = m_pUnknown->QueryInterface(
        IID_IUnknown, reinterpret_cast<void **>(&pUnknown1));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    IUnknown *pUnknown2;
    rhs.m_pUnknown->QueryInterface(
        IID_IUnknown, reinterpret_cast<void **>(&pUnknown2));
    if (FAILED(hr)) {
        pUnknown1->Release();
        _com_issue_error(hr);
    }

    bool result = pUnknown1 == pUnknown2;

    pUnknown1->Release();
    pUnknown2->Release();

    return result;
}

static void
throwDispatchException (EXCEPINFO &excepInfo)
{
    // Clean up exception information strings.
    _bstr_t source(excepInfo.bstrSource, false);
    _bstr_t description(excepInfo.bstrDescription, false);
    _bstr_t helpFile(excepInfo.bstrHelpFile, false);

    HRESULT hr = excepInfo.scode;
    if (hr == 0) {
        hr = _com_error::WCodeToHRESULT(excepInfo.wCode);
    }
    throw DispatchException(hr, description);
}

HRESULT
Reference::invokeDispatch (
    MEMBERID memberid,
    WORD dispatchFlags,
    const TypedArguments &arguments,
    VARIANT *pResult)
{
    IDispatch *pDispatch = dispatch();
    if (pDispatch == 0) {
        return E_NOINTERFACE;
    }

    // Remove missing optional arguments from the end of the argument list.
    // This permits calling servers which modify their action depending on
    // the actual number of arguments.
    DISPPARAMS *pParams = arguments.dispParams();

    // Count the number of missing arguments.
    unsigned cMissingArgs = 0;
    VARIANT *pArg = pParams->rgvarg + pParams->cNamedArgs;
    for (unsigned i = pParams->cNamedArgs; i < pParams->cArgs; ++i) {
        if (V_VT(pArg) != VT_ERROR || V_ERROR(pArg) != DISP_E_PARAMNOTFOUND) {
            break;
        }

        ++cMissingArgs;
        ++pArg;
    }

    // Move the named arguments up next to the remaining unnamed arguments and
    // adjust the DISPPARAMS struct.
    if (cMissingArgs > 0) {
        for (unsigned i = 0; i < pParams->cNamedArgs; ++i) {
            pParams->rgvarg[i + cMissingArgs] = pParams->rgvarg[i];
        }
        pParams->cArgs -= cMissingArgs;
        pParams->rgvarg += cMissingArgs;
    }

    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(excepInfo));
    unsigned argErr;

    // Invoke through IDispatch interface.
    HRESULT hr = pDispatch->Invoke(
        memberid,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        dispatchFlags,
        pParams,
        pResult,
        &excepInfo,
        &argErr);

    if (hr == DISP_E_EXCEPTION) {
        throwDispatchException(excepInfo);
    } else if (hr == DISP_E_TYPEMISMATCH || hr == DISP_E_PARAMNOTFOUND) {
        throw InvokeException(hr, pParams->cArgs - argErr);
    }

    return hr;
}

HRESULT
Reference::invoke (MEMBERID memberid,
                   WORD dispatchFlags,
                   const TypedArguments &arguments,
                   VARIANT *pResult)
{
    if (m_pInterface != 0 && !m_pInterface->dispatchOnly()) {
        EXCEPINFO excepInfo;
        memset(&excepInfo, 0, sizeof(excepInfo));
        unsigned argErr;

        // Invoke through virtual function table.
        ITypeInfo *pTypeInfo = interfaceDesc()->typeInfo();
        HRESULT hr = pTypeInfo->Invoke(
            m_pUnknown,
            memberid,
            dispatchFlags,
            arguments.dispParams(),
            pResult,
            &excepInfo,
            &argErr);

        if (SUCCEEDED(hr)) {
            return hr;
        }

        if (hr == DISP_E_EXCEPTION) {
            throwDispatchException(excepInfo);
        } else if (hr == DISP_E_TYPEMISMATCH || hr == DISP_E_PARAMNOTFOUND) {
            throw InvokeException(hr, arguments.dispParams()->cArgs - argErr);
        }
    }

    return invokeDispatch(memberid, dispatchFlags, arguments, pResult);
}

// IID of .NET Framework _Object interface
struct __declspec(uuid("65074F7F-63C0-304E-AF0A-D51741CB4A8D")) DotNetObject;

const Interface *
Reference::findInterfaceFromDispatch (IUnknown *pUnknown)
{
    HRESULT hr;

    // See if the object implements IDispatch.
    IDispatchPtr pDispatch;
    hr = pUnknown->QueryInterface(
        IID_IDispatch, reinterpret_cast<void **>(&pDispatch));
    if (FAILED(hr)) {
        return 0;
    }

    // Ask the IDispatch interface for type information.
    unsigned count;
    hr = pDispatch->GetTypeInfoCount(&count);
    if (hr == E_NOTIMPL) {
        return 0;
    }
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
    if (count == 0) {
        return 0;
    }

    ITypeInfoPtr pTypeInfo;
    hr = pDispatch->GetTypeInfo(
        0, 
        LOCALE_USER_DEFAULT,
        &pTypeInfo);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Get the interface description.
    TypeAttr typeAttr(pTypeInfo);

    if (IsEqualIID(typeAttr->guid, __uuidof(DotNetObject))) {
        // The .NET Framework implements IDispatch::GetTypeInfo for classes
        // declared with the attribute ClassInterface(ClassInterfaceType.None)
        // by returning a description of the _Object interface.
        return 0;
    }

    const Interface *pInterface =
        InterfaceManager::instance().newInterface(typeAttr->guid, pTypeInfo);

    if (pInterface->methods().empty() && pInterface->properties().empty()) {
        // No invokable methods or properties where found in the interface
        // description.
        return 0;
    }
    return pInterface;
}

const Interface *
Reference::findInterfaceFromClsid (REFCLSID clsid)
{
    const Interface *pInterface = 0;

    TypeLib *pTypeLib = TypeLib::loadByClsid(clsid);
    if (pTypeLib != 0) {
        const Class *pClass = pTypeLib->findClass(clsid);
        if (pClass != 0) {
            pInterface = pClass->defaultInterface();
        }
    }
    delete pTypeLib;

    return pInterface;
}

const Interface *
Reference::findInterfaceFromIid (REFIID iid)
{
    const Interface *pInterface = 0;

    TypeLib *pTypeLib = TypeLib::loadByIid(iid);
    if (pTypeLib != 0) {
        pInterface = InterfaceManager::instance().find(iid);
    }
    delete pTypeLib;

    return pInterface;
}

const Interface *
Reference::findInterface (IUnknown *pUnknown, REFCLSID clsid)
{
    const Interface *pInterface = 0;

    if (pUnknown != 0) {
        pInterface = findInterfaceFromDispatch(pUnknown);
    }

    if (pInterface == 0) {
        pInterface = findInterfaceFromClsid(clsid);
    }

    return pInterface;
}

Reference *
Reference::createInstance (
    REFCLSID clsid,
    const Interface *pInterface,
    DWORD clsCtx,
    const char *serverHost)
{
    // If we know it's a custom interface, then query for an interface pointer
    // to that interface, otherwise query for an IUnknown interface.
    const IID &iid = (pInterface == 0) ? IID_IUnknown : pInterface->iid();

    HRESULT hr;
    IUnknown *pUnknown;

    // Create an instance of the specified class.
#ifdef _WIN32_DCOM
    if (serverHost == 0
     || serverHost[0] == '\0'
     || strcmp(serverHost, "localhost") == 0
     || strcmp(serverHost, "127.0.0.1") == 0) {
        // When creating an instance on the local machine, call
        // CoCreateInstance instead of CoCreateInstanceEx with a null pointer
        // to COSERVERINFO.  This works around occasional failures in the RPC
        // DLL on Windows NT 4.0, even when connecting to a server on the local
        // machine.
        hr = CoCreateInstance(
            clsid,
            NULL,
            clsCtx,
            iid,
            reinterpret_cast<void **>(&pUnknown));
    } else {
        COSERVERINFO serverInfo;
        memset(&serverInfo, 0, sizeof(serverInfo));
	_bstr_t serverHostBstr(serverHost);
        serverInfo.pwszName = serverHostBstr;

        MULTI_QI qi;
        qi.pIID = &iid;
        qi.pItf = NULL;
        qi.hr = 0;

        hr = CoCreateInstanceEx(
            clsid,
            NULL,
            clsCtx,
            &serverInfo,
            1,
            &qi);
        if (SUCCEEDED(hr)) {
            pUnknown = static_cast<IUnknown *>(qi.pItf);
        }
    }
#else
    hr = CoCreateInstance(
        clsid,
        NULL,
        clsCtx,
        iid,
        reinterpret_cast<void **>(&pUnknown));
#endif
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    if (pInterface == 0) {
        pInterface = findInterface(pUnknown, clsid);
        if (pInterface != 0) {
            // Get a pointer to the derived interface.
            IUnknown *pNew;
            hr = pUnknown->QueryInterface(
                pInterface->iid(),
                reinterpret_cast<void **>(&pNew));
            if (SUCCEEDED(hr)) {
                pUnknown->Release();
                pUnknown = pNew;
            }
        }
    }

    return new Reference(pUnknown, pInterface, clsid);
}

Reference *
Reference::createInstance (
    const char *progId,
    DWORD clsCtx,
    const char *serverHost)
{
    // Convert the Prog ID to a CLSID.
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(_bstr_t(progId), &clsid);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    return createInstance(clsid, 0, clsCtx, serverHost);
}

Reference *
Reference::getActiveObject (REFCLSID clsid, const Interface *pInterface)
{
    HRESULT hr;

    // Retrieve the instance of the object.
    IUnknownPtr pActive;
    hr = GetActiveObject(clsid, NULL, &pActive);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // If we know it's a custom interface, then query for an interface pointer
    // to that interface, otherwise query for an IUnknown interface.
    IUnknown *pUnknown;
    hr = pActive->QueryInterface(
        (pInterface == 0) ? IID_IUnknown : pInterface->iid(),
        reinterpret_cast<void **>(&pUnknown));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    if (pInterface == 0) {
        pInterface = findInterface(pUnknown, clsid);
        if (pInterface != 0) {
            // Get a pointer to the derived interface.
            IUnknown *pNew;
            hr = pUnknown->QueryInterface(
                pInterface->iid(),
                reinterpret_cast<void **>(&pNew));
            if (SUCCEEDED(hr)) {
                pUnknown->Release();
                pUnknown = pNew;
            }
        }
    }

    return new Reference(pUnknown, pInterface, clsid);
}

Reference *
Reference::getActiveObject (const char *progId)
{
    // Convert the Prog ID to a CLSID.
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(_bstr_t(progId), &clsid);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    return getActiveObject(clsid, 0);
}

Reference *
Reference::getObject (const wchar_t *displayName)
{
    IUnknown *pUnknown;
    HRESULT hr = CoGetObject(
        displayName,
        NULL,
        IID_IUnknown,
        reinterpret_cast<void **>(&pUnknown));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    const Interface *pInterface = findInterfaceFromDispatch(pUnknown);
    return new Reference(pUnknown, pInterface);
}

Reference *
Reference::queryInterface (IUnknown *pOrig, REFIID iid)
{
    if (pOrig == 0) {
        _com_issue_error(E_POINTER);
    }

    IUnknown *pUnknown;
    HRESULT hr = pOrig->QueryInterface(
        iid,
        reinterpret_cast<void **>(&pUnknown));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    const Interface *pInterface = 0;
    if (!IsEqualIID(iid, IID_IDispatch)) {
        pInterface = findInterfaceFromIid(iid);
        if (pInterface == 0) {
            pInterface = findInterfaceFromDispatch(pUnknown);
        }
    }

    return new Reference(pUnknown, pInterface);
}

Reference *
Reference::newReference (IUnknown *pOrig, const Interface *pInterface)
{
    if (pOrig == 0) {
        _com_issue_error(E_POINTER);
    }

    if (pInterface == 0) {
        pInterface = findInterfaceFromDispatch(pOrig);
    }

    // If we know it's a custom interface, then query for an interface pointer
    // to that interface, otherwise query for an IUnknown interface.
    const IID &iid = (pInterface == 0) ? IID_IUnknown : pInterface->iid();

    IUnknown *pUnknown;
    HRESULT hr = pOrig->QueryInterface(
        iid,
        reinterpret_cast<void **>(&pUnknown));
    if (FAILED(hr)) {
        pUnknown = pOrig;
        pUnknown->AddRef();
    }

    return new Reference(pUnknown, pInterface);
}
