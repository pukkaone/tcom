// $Id: ComObject.cpp,v 1.41 2003/04/04 23:55:04 cthuang Exp $
#pragma warning(disable: 4786)
#include "ComObject.h"
#include <stdexcept>
#include "ComModule.h"
#include "InterfaceAdapter.h"
#include "Reference.h"
#include "Extension.h"

// prefix prepended to operation name for property get
static const char getPrefix[] = "_get_";

// prefix prepended to operation name for property put
static const char setPrefix[] = "_set_";

ComObject::ComObject (const Class::Interfaces &interfaces,
                      Tcl_Interp *interp,
                      TclObject servant,
                      TclObject destructor,
                      bool isSink):
    m_refCount(0),
    m_defaultInterface(*(interfaces.front())),
    m_interp(interp),
    m_servant(servant),
    m_destructor(destructor),
    m_supportErrorInfo(*this),
    m_pDispatch(0),
    m_registeredActiveObject(false),
    m_isSink(isSink)
{
//    Tcl_Preserve(reinterpret_cast<ClientData>(m_interp));
    ComModule::instance().lock();

    for (Class::Interfaces::const_iterator p = interfaces.begin();
     p != interfaces.end(); ++p) {
        const Interface *pInterface = *p;
        m_supportedInterfaceMap.insert(
            pInterface->iid(), const_cast<Interface *>(pInterface));
    }

    m_pDefaultAdapter = implementInterface(m_defaultInterface);
}

ComObject::~ComObject ()
{
    if (m_registeredActiveObject) {
        // TODO: This call may return an error but I don't want to throw an
        // exception from a destructor.
        RevokeActiveObject(m_activeObjectHandle, 0);
    }

    m_iidToAdapterMap.forEach(Delete());
    delete m_pDispatch;

    // Execute destructor Tcl command if defined.
    int length;
    Tcl_GetStringFromObj(m_destructor, &length);
    if (length > 0) {
        TclObject script(m_destructor);
        script.lappend(m_servant);
        eval(script);
    }

    ComModule::instance().unlock();
//    Tcl_Release(reinterpret_cast<ClientData>(m_interp));
}

void
ComObject::registerActiveObject (REFCLSID clsid)
{
    HRESULT hr = RegisterActiveObject(
        unknown(), clsid, ACTIVEOBJECT_WEAK, &m_activeObjectHandle);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
    m_registeredActiveObject = true;
}

InterfaceAdapter *
ComObject::implementInterface (const Interface &interfaceDesc)
{
    InterfaceAdapter *pAdapter = new InterfaceAdapter(*this, interfaceDesc);
    m_iidToAdapterMap.insert(interfaceDesc.iid(), pAdapter);
    return pAdapter;
}

ComObject *
ComObject::newInstance (
    const Interface &defaultInterface,
    Tcl_Interp *interp,
    TclObject servant,
    TclObject destructor,
    bool isSink)
{
    Class::Interfaces interfaces;
    interfaces.push_back(&defaultInterface);

    return new ComObject(
        interfaces,
        interp,
        servant,
        destructor,
        isSink);
}

ComObject *
ComObject::newInstance (
    const Class::Interfaces &interfaces,
    Tcl_Interp *interp,
    TclObject servant,
    TclObject destructor)
{
    return new ComObject(
        interfaces,
        interp,
        servant,
        destructor,
        false);
}

int
ComObject::eval (TclObject script, TclObject *pResult)
{
    int completionCode =
#if TCL_MINOR_VERSION >= 1
        Tcl_EvalObjEx(m_interp, script, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);
#else
        Tcl_GlobalEvalObj(m_interp, script);
#endif

    if (pResult != 0) {
        *pResult = Tcl_GetObjResult(m_interp);
    }
    return completionCode;
}

int
ComObject::getVariable (TclObject name, TclObject &value) const
{
    Tcl_Obj *pValue = Tcl_ObjGetVar2(m_interp, name, 0, TCL_LEAVE_ERR_MSG);
    if (pValue == 0) {
        return TCL_ERROR;
    }
    value = pValue;
    return TCL_OK;
}

int
ComObject::setVariable (TclObject name, TclObject value)
{
    Tcl_Obj *pValue =
        Tcl_ObjSetVar2(m_interp, name, 0, value, TCL_LEAVE_ERR_MSG);
    return (pValue == 0) ? TCL_ERROR : TCL_OK;
}

HRESULT
ComObject::hresultFromErrorCode () const
{
#if TCL_MINOR_VERSION >= 1
    Tcl_Obj *pErrorCode =
        Tcl_GetVar2Ex(m_interp, "::errorCode", 0, TCL_LEAVE_ERR_MSG);
#else
    TclObject errorCodeVarName("::errorCode");
    Tcl_Obj *pErrorCode =
        Tcl_ObjGetVar2(m_interp, errorCodeVarName, 0, TCL_LEAVE_ERR_MSG);
#endif

    if (pErrorCode == 0) {
        return E_UNEXPECTED;
    }

    Tcl_Obj *pErrorClass;
    if (Tcl_ListObjIndex(m_interp, pErrorCode, 0, &pErrorClass) != TCL_OK) {
        return E_UNEXPECTED;
    }
    if (strcmp(Tcl_GetStringFromObj(pErrorClass, 0), "COM") != 0) {
        return E_UNEXPECTED;
    }

    Tcl_Obj *pHresult;
    if (Tcl_ListObjIndex(m_interp, pErrorCode, 1, &pHresult) != TCL_OK) {
        return E_UNEXPECTED;
    }

    HRESULT hr;
    if (Tcl_GetLongFromObj(m_interp, pHresult, &hr) != TCL_OK) {
        return E_UNEXPECTED;
    }
    return hr;
}

// Implement IUnknown methods

HRESULT
ComObject::queryInterface (REFIID iid, void **ppvObj)
{
    if (IsEqualIID(iid, IID_IUnknown)) {
        *ppvObj = m_pDefaultAdapter;
        addRef();
        return S_OK;
    }

    if (IsEqualIID(iid, IID_IDispatch)) {
        // Expose the operations of the default interface through IDispatch.
        if (m_pDispatch == 0) {
            m_pDispatch = new InterfaceAdapter(*this, m_defaultInterface, true);
        }
        *ppvObj = m_pDispatch;
        addRef();
        return S_OK;
    }

    if (IsEqualIID(iid, IID_ISupportErrorInfo)) {
        *ppvObj = &m_supportErrorInfo;
        addRef();
        return S_OK;
    }

    InterfaceAdapter *pAdapter = m_iidToAdapterMap.find(iid);
    if (pAdapter == 0) {
        const Interface *pInterface = m_supportedInterfaceMap.find(iid);
        if (pInterface != 0) {
            pAdapter = implementInterface(*pInterface);
        }
    }

    if (pAdapter != 0) {
        *ppvObj = pAdapter;
        addRef();
        return S_OK;
    }

    *ppvObj = 0;
    return E_NOINTERFACE;
}

ULONG
ComObject::addRef ()
{
    InterlockedIncrement(&m_refCount);
    return m_refCount;
}

ULONG
ComObject::release ()
{
    InterlockedDecrement(&m_refCount);
    if (m_refCount == 0) {
        delete this;
        return 0;
    }
    return m_refCount;
}

// Generate a name for a Tcl variable used to hold an argument out value.

static TclObject
getOutVariableName (const Parameter &param)
{
    return TclObject(PACKAGE_NAMESPACE "arg_" + param.name());
}

// Convert IDispatch argument to Tcl value.

TclObject
ComObject::getArgument (VARIANT *pArg, const Parameter &param)
{
    if (vtMissing == pArg) {
        return Extension::newNaObj();

    } else if (param.flags() & PARAMFLAG_FOUT) {
        // Get name of Tcl variable to hold out value.
        TclObject varName = getOutVariableName(param);

        if (param.flags() & PARAMFLAG_FIN) {
            // For in/out parameters, set the Tcl variable to the input value.
            TclObject value(pArg, param.type(), m_interp);
            setVariable(varName, value);
        }
        return varName;

    } else {
        return TclObject(pArg, param.type(), m_interp);
    }
}

// Fill exception information structure.

static void
fillExcepInfo (EXCEPINFO *pExcepInfo,
               HRESULT hresult,
               const char *source,
               const char *description)
{
    if (pExcepInfo != 0) {
        memset(pExcepInfo, 0, sizeof(EXCEPINFO));
        pExcepInfo->scode = hresult;

        _bstr_t bstrSource(source);
        pExcepInfo->bstrSource = SysAllocString(bstrSource);

        if (description != 0) {
            _bstr_t bstrDescription(description);
            pExcepInfo->bstrDescription = SysAllocString(bstrDescription);
        }
    }
}

static void
putOutVariant (Tcl_Interp *interp,
               VARIANT *pDest,
               TclObject &tclObject,
               const Type &type)
{
    switch (type.vartype()) {
    case VT_BOOL:
        *V_BOOLREF(pDest) = tclObject.getBool() ? VARIANT_TRUE : VARIANT_FALSE;
        break;

    case VT_R4:
        *V_R4REF(pDest) = static_cast<float>(tclObject.getDouble());
        break;

    case VT_R8:
        *V_R8REF(pDest) = tclObject.getDouble();
        break;

    case VT_DISPATCH:
    case VT_UNKNOWN:
    case VT_USERDEFINED:
        {
            IUnknown *pUnknown;

            Tcl_Obj *pObj = tclObject;
            if (pObj->typePtr == &Extension::unknownPointerType) {
                pUnknown =
                    static_cast<IUnknown *>(pObj->internalRep.otherValuePtr);
            } else {
                Reference *pRef = Extension::referenceHandles.find(
                    interp, tclObject);
                pUnknown = (pRef == 0) ? 0 : pRef->unknown();
            }

            *V_UNKNOWNREF(pDest) = pUnknown;

            // The COM rules say we must increment the reference count of
            // interface pointers returned from methods.
            if (pUnknown != 0) {
                pUnknown->AddRef();
            }
        }
        break;

    case VT_BSTR:
        *V_BSTRREF(pDest) = tclObject.getBSTR();
        break;

    case VT_VARIANT:
        {
            // Must increment reference count of interface pointers returned
            // from methods.
            tclObject.toVariant(
                V_VARIANTREF(pDest), Type::variant(), interp, true);
        }
        break;

    default:
        *V_I4REF(pDest) = tclObject.getLong();
    }
}

HRESULT
ComObject::invoke (InterfaceAdapter *pAdapter,
                   DISPID dispid,
                   REFIID /*riid*/,
                   LCID /*lcid*/,
                   WORD wFlags,
                   DISPPARAMS *pDispParams,
                   VARIANT *pReturnValue,
                   EXCEPINFO *pExcepInfo,
                   UINT *pArgErr)
{
    // Get the method description for method being invoked.
    const Method *pMethod = pAdapter->findDispatchMethod(dispid);
    if (pMethod == 0) {
        return DISP_E_MEMBERNOTFOUND;
    }

    HRESULT hresult;

    try {
        // Construct Tcl script to invoke operation on the servant.
        TclObject script(m_servant);

        // Get the method or property to invoke on the servant.
        std::string operation;
        if ((wFlags & DISPATCH_PROPERTYGET) != 0
         && pAdapter->isProperty(dispid)) {
            operation = getPrefix + pMethod->name();

        } else if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
            operation = setPrefix + pMethod->name();

        } else if (wFlags & DISPATCH_METHOD) {
            operation = pMethod->name();

        } else {
            return DISP_E_MEMBERNOTFOUND;
        }

        script.lappend(
            Tcl_NewStringObj(const_cast<char *>(operation.c_str()), -1));

        // Set the argument error pointer in case we need to use it.
        UINT argErr;
        if (pArgErr == 0) {
            pArgErr = &argErr;
        }

        // Convert arguments to Tcl values.
        // TODO: Should handle named arguments differently than positional
        // arguments.
        const Method::Parameters &parameters = pMethod->parameters();

        int argIndex = pDispParams->cArgs - 1;
        Method::Parameters::const_iterator pParam;
        for (pParam = parameters.begin(); pParam != parameters.end();
         ++pParam, --argIndex) {
            // Append argument value.
            VARIANT *pArg = &(pDispParams->rgvarg[argIndex]);
            try {
                script.lappend(getArgument(pArg, *pParam));
            }
            catch (_com_error &) {
                *pArgErr = argIndex;
                throw;
            }
        }
        
        if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
            VARIANT *pArg = &(pDispParams->rgvarg[argIndex]);
            try {
                TclObject value(pArg, pMethod->type(), m_interp);
                script.lappend(value);
            }
            catch (_com_error &) {
                *pArgErr = argIndex;
                throw;
            }
        }

        // Execute the Tcl script.
        TclObject result;
        int completionCode = eval(script, &result);
        if (completionCode == TCL_OK) {
            hresult = S_OK;
        } else {
            if (m_isSink) {
                Tcl_BackgroundError(m_interp);
            }

            hresult = hresultFromErrorCode();
            if (FAILED(hresult)) {
                fillExcepInfo(
                    pExcepInfo,
                    hresult,
                    m_servant.c_str(),
                    result.c_str());
                hresult = DISP_E_EXCEPTION;
            }
        }

        // Copy values to out arguments.
        argIndex = pDispParams->cArgs - 1;
        for (pParam = parameters.begin(); pParam != parameters.end();
         ++pParam, --argIndex) {
            if (pParam->flags() & PARAMFLAG_FOUT) {
                // Get name of Tcl variable that holds out value.
                TclObject varName = getOutVariableName(*pParam);

                // Copy variable value to out argument.
                TclObject value;
                if (getVariable(varName, value) == TCL_OK) {
                    putOutVariant(
                        m_interp,
                        &pDispParams->rgvarg[argIndex],
                        value,
                        pParam->type());
                }
            }
        }

        // Convert return value.
        if (pReturnValue != 0 && pMethod->type().vartype() != VT_VOID) {
            // Must increment reference count of interface pointers returned
            // from methods.
            result.toVariant(pReturnValue, pMethod->type(), m_interp, true);
        }
    }
    catch (_com_error &e) {
        fillExcepInfo(pExcepInfo, e.Error(), m_servant.c_str(), 0);
        hresult = DISP_E_EXCEPTION;
    }
    return hresult;
}

// Convert the native value that the va_list points to into a Tcl object.
// Returns a va_list pointing to the next argument.

static va_list
convertNativeToTclObject (va_list pArg,
                          Tcl_Interp *interp,
                          TclObject &tclObject,
                          const Type &type,
                          bool byRef=false)
{
    switch (type.vartype()) {
    case VT_BOOL:
        tclObject = Tcl_NewBooleanObj(
            byRef ? *va_arg(pArg, VARIANT_BOOL *) : va_arg(pArg, VARIANT_BOOL));
        break;

    case VT_DATE:
    case VT_R4:
    case VT_R8:
        tclObject = Tcl_NewDoubleObj(
            byRef ? *va_arg(pArg, double *) : va_arg(pArg, double));
        break;

    case VT_USERDEFINED:
        if (type.name() == "GUID") {
            UUID *pUuid = va_arg(pArg, UUID *);
            Uuid uuid(*pUuid);
            tclObject = Tcl_NewStringObj(
                const_cast<char *>(uuid.toString().c_str()), -1);
            break;
        }
        // Fall through

    case VT_DISPATCH:
    case VT_UNKNOWN:
        {
            IUnknown *pUnknown = va_arg(pArg, IUnknown *);
            if (pUnknown == 0) {
                tclObject = Tcl_NewObj();
            } else {
                const Interface *pInterface =
                    InterfaceManager::instance().find(type.iid());
                tclObject = Extension::referenceHandles.newObj(
                    interp, Reference::newReference(pUnknown, pInterface));
            }
        }
        break;

    case VT_NULL:
        tclObject = Tcl_NewObj();
        break;

    case VT_LPWSTR:
    case VT_BSTR:
        {
#if TCL_MINOR_VERSION >= 2
            // Uses Unicode function introduced in Tcl 8.2.
            Tcl_UniChar *pUnicode = byRef ?
                *va_arg(pArg, Tcl_UniChar **) : va_arg(pArg, Tcl_UniChar *);
            if (pUnicode != 0) {
                tclObject = Tcl_NewUnicodeObj(pUnicode, -1);
            } else {
                tclObject = Tcl_NewObj();
            }
#else
            wchar_t *pUnicode = byRef ?
                *va_arg(pArg, wchar_t **) : va_arg(pArg, wchar_t *);
            _bstr_t str(pUnicode);
            tclObject = Tcl_NewStringObj(str, -1);
#endif
        }
        break;

    case VT_VARIANT:
        tclObject = TclObject(
            byRef ? va_arg(pArg, VARIANT *) : &va_arg(pArg, VARIANT),
            type,
            interp);
        break;

    default:
        tclObject = Tcl_NewLongObj(
            byRef ? *va_arg(pArg, int *) : va_arg(pArg, int));
    }

    return pArg;
}

// Convert the native value that the va_list points to into a Tcl value.
// Returns a va_list pointing to the next argument.

va_list
ComObject::getArgument (
    va_list pArg, const Parameter &param, TclObject &dest)
{
    if (param.flags() & PARAMFLAG_FOUT) {
        // Get name of Tcl variable to hold out value.
        TclObject varName = getOutVariableName(param);

        if (param.flags() & PARAMFLAG_FIN) {
            // For in/out parameters, set the Tcl variable to the input value.
            TclObject value;
            pArg = convertNativeToTclObject(
                pArg, m_interp, value, param.type(), true);
            setVariable(varName, value);
        } else {
            // Advance to next argument.
            va_arg(pArg, void *);
        }
        dest = varName;
        return pArg;

    } else {
        return convertNativeToTclObject(
            pArg, m_interp, dest, param.type());
    }
}

// Convert Tcl value to native value and store it at the address the va_list
// points to.
// Returns a va_list pointing to the next argument.

static va_list
putArgument (va_list pArg,
             Tcl_Interp *interp,
             TclObject tclObject,
             const Type &type)
{
    void *pDest = va_arg(pArg, void *);
    if (pDest == 0) {
        return pArg;
    }

    switch (type.vartype()) {
    case VT_BOOL:
        *static_cast<VARIANT_BOOL *>(pDest) =
            tclObject.getBool() ? VARIANT_TRUE : VARIANT_FALSE;
        break;

    case VT_R4:
        *static_cast<float *>(pDest) =
            static_cast<float>(tclObject.getDouble());
        break;

    case VT_R8:
        *static_cast<double *>(pDest) = tclObject.getDouble();
        break;

    case VT_USERDEFINED:
        if (type.name() == "GUID") {
            char *uuidStr = const_cast<char *>(tclObject.c_str());
            UUID uuid;
            UuidFromString(reinterpret_cast<unsigned char *>(uuidStr), &uuid);
            *static_cast<UUID *>(pDest) = uuid;
            break;
        }
        // Fall through

    case VT_DISPATCH:
    case VT_UNKNOWN:
        {
            IUnknown *pUnknown;

            Tcl_Obj *pObj = tclObject;
            if (pObj->typePtr == &Extension::unknownPointerType) {
                pUnknown =
                    static_cast<IUnknown *>(pObj->internalRep.otherValuePtr);
            } else {
                Reference *pRef = Extension::referenceHandles.find(
                    interp, tclObject);
                pUnknown = (pRef == 0) ? 0 : pRef->unknown();
            }

            *static_cast<IUnknown **>(pDest) = pUnknown;

            // The COM rules say we must increment the reference count of
            // interface pointers returned from methods.
            if (pUnknown != 0) {
                pUnknown->AddRef();
            }
        }
        break;

    case VT_BSTR:
        *static_cast<BSTR *>(pDest) = tclObject.getBSTR();
        break;

    case VT_VARIANT:
        {
            // Must increment reference count of interface pointers returned
            // from methods.
            tclObject.toVariant(
                static_cast<VARIANT *>(pDest),
                Type::variant(),
                interp,
                true);
        }
        break;

    default:
        *static_cast<int *>(pDest) = tclObject.getLong();
    }

    return pArg;
}

// Advance the va_list to the next argument.
// Returns a va_list pointing to the next argument.

static va_list
nextArgument (va_list pArg, const Type &type)
{
    switch (type.vartype()) {
    case VT_R4:
    case VT_R8:
    case VT_DATE:
        va_arg(pArg, double);
        break;

    case VT_DISPATCH:
    case VT_UNKNOWN:
    case VT_USERDEFINED:
        va_arg(pArg, IUnknown *);
        break;

    case VT_BSTR:
        va_arg(pArg, BSTR);
        break;

    case VT_VARIANT:
        va_arg(pArg, VARIANT);
        break;

    default:
        va_arg(pArg, int);
    }

    return pArg;
}

// Set error info.

static void
setErrorInfo (const char *source, const char *description)
{
    HRESULT hr;

    ICreateErrorInfoPtr pCreateErrorInfo;
    hr = CreateErrorInfo(&pCreateErrorInfo);
    if (FAILED(hr)) {
        return;
    }

    _bstr_t sourceBstr(source);
    pCreateErrorInfo->SetSource(sourceBstr);

    _bstr_t descriptionBstr(description);
    pCreateErrorInfo->SetDescription(descriptionBstr);

    IErrorInfoPtr pErrorInfo;
    hr = pCreateErrorInfo->QueryInterface(
        IID_IErrorInfo, reinterpret_cast<void **>(&pErrorInfo));
    if (SUCCEEDED(hr)) {
        SetErrorInfo(0, pErrorInfo);
    }
}

// Note that this function is called in an odd way to avoid copying the
// arguments onto the stack (for efficiency and simplicity in the calling
// code).  This is why the call is explicitly declared __cdecl.

void __cdecl
invokeComObjectFunction (volatile HRESULT hresult,
                         volatile DWORD pArgEnd,
                         DWORD /*ebp*/,
                         DWORD funcIndex,
                         DWORD /*retAddr*/,
                         InterfaceAdapter *pAdapter,
                         ...)
{
    // Get the method description for method being invoked.
    const Method *pMethod = pAdapter->findComMethod(funcIndex);
    if (pMethod == 0) {
        // If we don't have a method description, we don't know how many bytes
        // the arguments take on the stack.
        throw std::runtime_error("unknown virtual function index");
    }

    ComObject &object = pAdapter->object();

    // Construct Tcl script to invoke operation on the servant.
    TclObject script(object.m_servant);

    std::string operation;
    switch (pMethod->invokeKind()) {
    case INVOKE_PROPERTYGET:
        operation = getPrefix + pMethod->name();
        break;
    case INVOKE_PROPERTYPUT:
    case INVOKE_PROPERTYPUTREF:
        operation = setPrefix + pMethod->name();
        break;
    default:
        operation = pMethod->name();
    }
    script.lappend(
        Tcl_NewStringObj(const_cast<char *>(operation.c_str()), -1));

    // Convert arguments to Tcl values.
    va_list pArg;
    va_start(pArg, pAdapter);
    const Method::Parameters &parameters = pMethod->parameters();
    Method::Parameters::const_iterator pParam = parameters.begin();
    for (; pParam != parameters.end(); ++pParam) {
        // Append argument value.
        TclObject argument;
        pArg = object.getArgument(pArg, *pParam, argument);
        script.lappend(argument);
    }

    // Set end of arguments pointer.
    if (pMethod->type().vartype() != VT_VOID) {
        va_arg(pArg, void *);
    }
    pArgEnd = reinterpret_cast<DWORD>(pArg);
    va_end(pArg);

    // Execute the Tcl script.
    TclObject result;
    int completionCode = object.eval(script, &result);
    if (completionCode == TCL_OK) {
        hresult = S_OK;
    } else {
        hresult = object.hresultFromErrorCode();
        if (FAILED(hresult)) {
            setErrorInfo(object.m_servant.c_str(), result.c_str());
        }
    }

    // Copy values to out arguments.
    va_start(pArg, pAdapter);
    pParam = parameters.begin();
    for (; pParam != parameters.end(); ++pParam) {
        if (pParam->flags() & PARAMFLAG_FOUT) {
            // Get name of Tcl variable that holds out value.
            TclObject varName = getOutVariableName(*pParam);

            // Copy variable value to out argument.
            TclObject value;
            if (object.getVariable(varName, value) == TCL_OK) {
                pArg = putArgument(
                    pArg, object.m_interp, value, pParam->type());
                continue;
            }
        }

        pArg = nextArgument(pArg, pParam->type());
    }

    // Convert return value.
    if (pMethod->type().vartype() != VT_VOID) {
        putArgument(pArg, object.m_interp, result, pMethod->type());
    }

    va_end(pArg);
}
