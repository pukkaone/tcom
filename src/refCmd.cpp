// $Id: refCmd.cpp 16 2005-04-19 14:47:52Z cthuang $
#pragma warning(disable: 4786)
#include "Extension.h"
#include <sstream>
#include "Reference.h"
#include "TypeInfo.h"
#include "TclObject.h"
#include "Arguments.h"

static int referenceObjCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
HandleSupport<Reference> Extension::referenceHandles(referenceObjCmd);

static const char unknownErrorDescription[] = "Unknown error";

// Check if the object implements ISupportErrorInfo.  If it does, get the
// error information.  Return true if successful.

static bool
getErrorInfo (Reference *pReference, IErrorInfo **ppErrorInfo)
{
    const Interface *pInterface = pReference->interfaceDesc();

    // The .NET Framework uses GUID_NULL to identify the interface which
    // raised the error.
    const IID &iid = (pInterface == 0) ? GUID_NULL : pInterface->iid();

    ISupportErrorInfoPtr pSupportErrorInfo;
    HRESULT hr = pReference->unknown()->QueryInterface(
        IID_ISupportErrorInfo, reinterpret_cast<void **>(&pSupportErrorInfo));
    if (FAILED(hr)) {
        return false;
    }

    if (pSupportErrorInfo->InterfaceSupportsErrorInfo(iid) != S_OK) {
        return false;
    }

    return GetErrorInfo(0, ppErrorInfo) == S_OK;
}

// Get description text for an HRESULT.

static Tcl_Obj *
formatMessage (HRESULT hresult)
{
#if TCL_MINOR_VERSION >= 2
    // Uses Unicode functions introduced in Tcl 8.2.
    wchar_t *pMessage;
    DWORD nLen = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
        NULL,
        hresult,
        0,
        reinterpret_cast<LPWSTR>(&pMessage),
        0,
        NULL);
#else
    char *pMessage;
    DWORD nLen = FormatMessageA(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
        NULL,
        hresult,
        0,
        reinterpret_cast<LPSTR>(&pMessage),
        0,
        NULL);
#endif

    Tcl_Obj *pDescription;
    if (nLen > 0) {
        if (nLen > 1 && pMessage[nLen - 1] == '\n') {
            --nLen;
            if (nLen > 1 && pMessage[nLen - 1] == '\r') {
                --nLen;
            }
        }
        pMessage[nLen] = '\0';

        
#if TCL_MINOR_VERSION >= 2
        // Uses Unicode functions introduced in Tcl 8.2.
        pDescription = Tcl_NewUnicodeObj(pMessage, nLen);
#else
        pDescription = Tcl_NewStringObj(pMessage, nLen);
#endif
    } else {
        pDescription = Tcl_NewStringObj(unknownErrorDescription, -1);
    }
    LocalFree(pMessage);

    return pDescription;
}

// Set the Tcl errorCode variable and the Tcl interpreter result.
// Returns TCL_ERROR.

static int
setErrorCodeAndResult (
    Tcl_Interp *interp,
    HRESULT hresult,
    Tcl_Obj *pDescription,
    const char *file,
    int line)
{
    TclObject errorCode(Tcl_NewListObj(0, 0));
    errorCode.lappend(Tcl_NewStringObj("COM", -1));

    TclObject result(Tcl_NewListObj(0, 0));

    // Append HRESULT value in hexadecimal string format.
    std::ostringstream hrOut;
    hrOut << "0x" << std::hex << hresult;
    TclObject hrObj(hrOut.str());
    errorCode.lappend(hrObj);
    result.lappend(hrObj);

    // Append description.
    errorCode.lappend(pDescription);
    result.lappend(pDescription);

#ifndef NDEBUG
    // Append file and line number.
    std::ostringstream fileLine;
    fileLine << file << ' ' << line;
    TclObject fileLineObj(fileLine.str());
    result.lappend(fileLineObj);
#endif

    Tcl_SetObjErrorCode(interp, errorCode);
    Tcl_SetObjResult(interp, result);
    return TCL_ERROR;
}

static int
setErrorCodeAndResult (
    Tcl_Interp *interp,
    HRESULT hresult,
    const _bstr_t &description,
    const char *file,
    int line)
{
    TclObject descriptionObj;
    int length;
    Tcl_GetStringFromObj(descriptionObj, &length);
    if (length == 0) {
        descriptionObj = Tcl_NewStringObj(unknownErrorDescription, -1);
    }
    return setErrorCodeAndResult(interp, hresult, descriptionObj, file, line);
}

int
Extension::setComErrorResult (
    Tcl_Interp *interp, _com_error &e, const char *file, int line)
{
    return setErrorCodeAndResult(
        interp, e.Error(), formatMessage(e.Error()), file, line);
}

// Invoke a method or property.

static int
invoke (Tcl_Interp *interp,
        int objc,               // number of arguments
        Tcl_Obj *CONST objv[],  // arguments
        Reference *pReference,
        const Method *pMethod,
        bool namedArgOpt,
        TypedArguments &arguments,
        WORD dispatchFlags)
{
    // Set up return value.
    NativeValue returnValue;
    VARIANT *pReturnValue = (pMethod->type().vartype() == VT_VOID)
        ? 0 : &returnValue;

    // Invoke it.
    HRESULT hr;
    if (namedArgOpt) {
        hr = pReference->invokeDispatch(
            pMethod->memberid(),
            dispatchFlags,
            arguments,
            pReturnValue);
    } else {
        hr = pReference->invoke(
            pMethod->memberid(),
            dispatchFlags,
            arguments,
            pReturnValue);
    }
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Store values returned from out parameters.
    arguments.storeOutValues(interp, objc, objv, pMethod->parameters());

    // Convert return value.
    if (pReturnValue != 0) {
        TclObject value(pReturnValue, pMethod->type(), interp);
        Tcl_SetObjResult(interp, value);
    }
    return TCL_OK;
}

// Set Tcl result to a wrong number of arguments error message.

static void
wrongNumArgs (
    Tcl_Interp *interp,		            // current interpreter
    Tcl_Obj *CONST objv[],                  // method name
    const Method::Parameters &parameters)   // expected parameters
{
    if (parameters.size() > 0) {
        std::ostringstream paramNames;
        bool first = true;
        for (Property::Parameters::const_iterator p =
         parameters.begin(); p != parameters.end(); ++p) {
            if (first) {
                first = false;
            } else {
                paramNames << ' ';
            }
            paramNames << p->name();
        }
        Tcl_WrongNumArgs(
            interp, 1, objv, const_cast<char *>(paramNames.str().c_str()));
    } else {
        Tcl_WrongNumArgs(interp, 1, objv, 0);
    }

}

// Get or put an object property.

static int
invokeProperty (
    Tcl_Interp *interp,		// Current interpreter
    int objc,
    Tcl_Obj *CONST objv[],	// property name and arguments
    Reference *pReference,
    const Property *pProperty)
{
    WORD dispatchFlags;
    const Property::Parameters &parameters = pProperty->parameters();

    if (objc > parameters.size() + 2) {
	wrongNumArgs(interp, objv, parameters);
        return TCL_ERROR;

    } else if (objc == parameters.size() + 2) {
        // Put property.
        dispatchFlags = pProperty->putDispatchFlag();

    } else {
        // Get property.
        dispatchFlags = DISPATCH_PROPERTYGET;
    }

    PositionalArguments arguments;
    int result = arguments.initialize(
        interp, objc - 1, objv + 1, *pProperty, dispatchFlags);
    if (result != TCL_OK) {
        return result;
    }

    return invoke(
        interp,
        objc - 1,
        objv + 1,
        pReference,
        pProperty,
        false,
        arguments,
        dispatchFlags);
}

// Invoke a method without any type information using IDispatch.
// Return a Tcl completion code.

static int
invokeWithoutInterfaceDesc (
    Tcl_Interp *interp,
    Reference *pReference,
    int objc,
    Tcl_Obj *CONST objv[],      // method name and arguments
    WORD dispatchFlags)
{
    HRESULT hr;

    IDispatch *pDispatch = pReference->dispatch();
    if (pDispatch == 0) {
        Tcl_AppendResult(interp, "object does not implement IDispatch", NULL);
        return TCL_ERROR;
    }

    // Ask for named method or property.
    const char *name = Tcl_GetStringFromObj(objv[0], 0);
    _bstr_t bstrName(name);
    OLECHAR *names[1];
    names[0] = bstrName;

    DISPID dispatchID;
    hr = pDispatch->GetIDsOfNames(
        IID_NULL, names, 1, LOCALE_USER_DEFAULT, &dispatchID);
    if (FAILED(hr)) {
        Tcl_AppendResult(
            interp,
            "object does not implement method or property ",
            name,
            NULL);
        return TCL_ERROR;
    }

    UntypedArguments arguments;
    int result = arguments.initialize(
        interp, objc - 1, objv + 1, dispatchFlags);
    if (result != TCL_OK) {
        return result;
    }

    // Set up return value.
    NativeValue varReturnValue;
    VARIANT *pReturnValue =
        (dispatchFlags & DISPATCH_PROPERTYPUT) ? 0 : &varReturnValue;

    // Invoke method.
    EXCEPINFO excepInfo;
    memset(&excepInfo, 0, sizeof(excepInfo));

    unsigned argErr;
    hr = pDispatch->Invoke(
        dispatchID,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        dispatchFlags,
        arguments.dispParams(),
        pReturnValue,
        &excepInfo,
        &argErr);
    if (hr == DISP_E_EXCEPTION) {
        // Clean up exception information strings.
        _bstr_t source(excepInfo.bstrSource, false);
        _bstr_t description(excepInfo.bstrDescription, false);
        _bstr_t helpFile(excepInfo.bstrHelpFile, false);

        throw DispatchException(excepInfo.scode, description);
    }

    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    if (pReturnValue != 0) {
        TclObject returnValue(pReturnValue, Type::variant(), interp);
        Tcl_SetObjResult(interp, returnValue);
    }
    return TCL_OK;
}

// This Tcl command invokes a method or property on an interface pointer.

static int
referenceObjCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    bool namedArgOpt = false;
    WORD dispatchFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;

    int i = 1;
    for (; i < objc; ++i) {
        static char *options[] = {
	    "-call", "-get", "-method", "-namedarg", "-set", NULL
        };
        enum OptionEnum {
            OPTION_CALL, OPTION_GET, OPTION_METHOD, OPTION_NAMEDARG, OPTION_SET
        };

        int index;
        if (Tcl_GetIndexFromObj(NULL, objv[i], options, "option", 0, &index)
         != TCL_OK) {
            break;
        }

        switch (index) {
        case OPTION_CALL:
        case OPTION_METHOD:
            dispatchFlags = DISPATCH_METHOD;
            break;
        case OPTION_GET:
            dispatchFlags = DISPATCH_PROPERTYGET;
            break;
        case OPTION_NAMEDARG:
            namedArgOpt = true;
            break;
        case OPTION_SET:
            dispatchFlags = DISPATCH_PROPERTYPUT;
            break;
        }
    }

    if (objc - i < 1) {
	Tcl_AppendResult(
            interp, "usage: handle ?options? method ?arg ...?", NULL);
	return TCL_ERROR;
    }

    Reference *pReference = reinterpret_cast<Reference *>(clientData);

    int result;
    try {
        // Get interface description.
        const Interface *pInterface = pReference->interfaceDesc();
        if (pInterface == 0) {
            return invokeWithoutInterfaceDesc(
                interp, pReference, objc - i, objv + i, dispatchFlags);
        }

        const Method *pMethod;
        const Property *pProperty;
        const char *name = Tcl_GetStringFromObj(objv[i], 0);

        if ((pProperty = pInterface->findProperty(name)) != 0) {
            // It's a property.
            result = invokeProperty(
                interp,
                objc - i,
                objv + i,
                pReference,
                pProperty);

        } else if ((pMethod = pInterface->findMethod(name)) != 0) {
            // It's a method.
            ++i;
            NamedArguments namedArguments;
            PositionalArguments positionalArguments;
            TypedArguments *pArguments;

            if (namedArgOpt) {
                pArguments = &namedArguments;
            } else {
                // Return an error if too many arguments were given.
                const Method::Parameters &parameters = pMethod->parameters();
                if (!pMethod->vararg() && objc - i > parameters.size()) {
                    wrongNumArgs(interp, objv + i - 1, parameters);
                    return TCL_ERROR;
                }
                pArguments = &positionalArguments;
            }
            result = pArguments->initialize(
                interp, objc - i, objv + i, *pMethod, DISPATCH_METHOD);
            if (result != TCL_OK) {
                return result;
            }

            result = invoke(
                interp,
                objc - i,
                objv + i,
                pReference,
                pMethod,
                namedArgOpt,
                *pArguments,
                DISPATCH_METHOD);

        } else {
            Tcl_AppendResult(
                interp,
                "interface ",
                pInterface->name().c_str(),
                " does not have method or property ",
                name,
                NULL);
            result = TCL_ERROR;
        }
    }
    catch (_com_error &e) {
        IErrorInfoPtr pErrorInfo;
        if (getErrorInfo(pReference, &pErrorInfo)) {
            BSTR descBstr;
            pErrorInfo->GetDescription(&descBstr);
            _bstr_t description(descBstr, false);

            result = setErrorCodeAndResult(
                interp, e.Error(), description, __FILE__, __LINE__);
        } else {
            result = Extension::setComErrorResult(
                interp, e, __FILE__, __LINE__);
        }
    }
    catch (DispatchException &e) {
        result = setErrorCodeAndResult(
            interp, e.scode(), e.description(), __FILE__, __LINE__);
    }
    catch (InvokeException &e) {
        std::ostringstream argOut;
        argOut << "Argument " << e.argIndex() << ": ";
        TclObject descriptionObj(argOut.str());

        TclObject messageObj(formatMessage(e.hresult()));
        Tcl_AppendObjToObj(descriptionObj, messageObj);

        result = setErrorCodeAndResult(
            interp, e.hresult(), descriptionObj, __FILE__, __LINE__);
    }

    return result;
}

// This command gets an interface pointer to an object identified by a moniker.

static int
getObjectCmd (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 3) {
	Tcl_WrongNumArgs(interp, 2, objv, "monikerName");
	return TCL_ERROR;
    }

#if TCL_MINOR_VERSION >= 2
    const wchar_t *monikerName = Tcl_GetUnicode(objv[2]);
#else
    _bstr_t monikerName(Tcl_GetStringFromObj(objv[2], 0));
#endif

    try {
        Reference *pReference = Reference::getObject(monikerName);
        Tcl_SetObjResult(
            interp, Extension::referenceHandles.newObj(interp, pReference));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}

// This command returns the reference count of an interface pointer.

static int
countCmd(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 3) {
        Tcl_WrongNumArgs(interp, 2, objv, "handle");
        return TCL_ERROR;
    }

    Reference *pReference = Extension::referenceHandles.find(interp, objv[2]);
    if (pReference == 0) {
        char *arg = Tcl_GetStringFromObj(objv[2], 0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle ", arg, NULL);
        return TCL_ERROR;
    }

    IUnknown *pUnknown = pReference->unknown();
    pUnknown->AddRef();
    long count = pUnknown->Release();

    Tcl_SetObjResult(interp, Tcl_NewLongObj(count));
    return TCL_OK;
}

// This command compares two interface pointers for COM identity.

static int
equalCmd(
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 4) {
        Tcl_WrongNumArgs(interp, 2, objv, "handle1 handle2");
        return TCL_ERROR;
    }

    Reference *pReference1 = Extension::referenceHandles.find(interp, objv[2]);
    if (pReference1 == 0) {
        char *arg = Tcl_GetStringFromObj(objv[2], 0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle1 ", arg, NULL);
        return TCL_ERROR;
    }

    Reference *pReference2 = Extension::referenceHandles.find(interp, objv[3]);
    if (pReference2 == 0) {
        char *arg = Tcl_GetStringFromObj(objv[3], 0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle2 ", arg, NULL);
        return TCL_ERROR;
    }

    Tcl_SetObjResult(
        interp, Tcl_NewBooleanObj(*pReference1 == *pReference2));
    return TCL_OK;
}

// This command queries an interface pointer for an IDispatch interface.

static int
queryDispatchCmd (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 3) {
	Tcl_WrongNumArgs(interp, 2, objv, "handle");
	return TCL_ERROR;
    }

    Reference *pReference = Extension::referenceHandles.find(interp, objv[2]);
    if (pReference == 0) {
        char *arg = Tcl_GetStringFromObj(objv[2], (int *)0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle ", arg, NULL);
        return TCL_ERROR;
    }

    try {
        Reference *pNewRef = Reference::queryInterface(
            pReference->unknown(), IID_IDispatch);
        Tcl_SetObjResult(
            interp,
            Extension::referenceHandles.newObj(interp, pNewRef));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}

// This command queries an interface pointer for a given interface.

static int
queryInterfaceCmd (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 4) {
	Tcl_WrongNumArgs(interp, 2, objv, "handle IID");
	return TCL_ERROR;
    }

    Reference *pReference = Extension::referenceHandles.find(interp, objv[2]);
    if (pReference == 0) {
        char *arg = Tcl_GetStringFromObj(objv[2], (int *)0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle ", arg, NULL);
        return TCL_ERROR;
    }

    char *iidStr = Tcl_GetStringFromObj(objv[3], (int *)0);
    IID iid;
    if (UuidFromString(reinterpret_cast<unsigned char *>(iidStr), &iid)
     != RPC_S_OK) {
	Tcl_AppendResult(
            interp,
            "cannot convert to IID: ",
            iidStr,
            NULL);
        return TCL_ERROR;
    }

    try {
        Reference *pNewRef = Reference::queryInterface(
            pReference->unknown(), iid);
        Tcl_SetObjResult(
            interp,
            Extension::referenceHandles.newObj(interp, pNewRef));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}

// This Tcl command gets a reference to an object.

int
Extension::refCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 3) {
	Tcl_WrongNumArgs(interp, 1, objv, "subcommand argument ...");
	return TCL_ERROR;
    }

    Extension *pExtension =
        static_cast<Extension *>(clientData);
    pExtension->initializeCom();

    static char *options[] = {
        "count",
	"createobject",
        "equal",
        "getactiveobject",
        "getobject",
        "querydispatch",
        "queryinterface",
        NULL
    };
    enum SubCommandEnum {
        COUNT, CREATEOBJECT, EQUAL, GETACTIVEOBJECT, GETOBJECT, QUERYDISPATCH,
        QUERYINTERFACE
    };

    int subCommand;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "subcommand", 0,
     &subCommand) != TCL_OK) {
	return TCL_ERROR;
    }

    switch (subCommand) {
    case COUNT:
        return countCmd(interp, objc, objv);
    case EQUAL:
        return equalCmd(interp, objc, objv);
    case GETOBJECT:
        return getObjectCmd(interp, objc, objv);
    case QUERYDISPATCH:
        return queryDispatchCmd(interp, objc, objv);
    case QUERYINTERFACE:
        return queryInterfaceCmd(interp, objc, objv);
    }

    bool clsIdOpt = false;
    DWORD clsCtx = CLSCTX_SERVER;

    int i = 2;
    for (; i < objc; ++i) {
        static char *options[] = {
	    "-clsid", "-inproc", "-local", "-remote", NULL
        };
        enum OptionEnum {
            OPTION_CLSID, OPTION_INPROC, OPTION_LOCAL, OPTION_REMOTE
        };

        int index;
        if (Tcl_GetIndexFromObj(NULL, objv[i], options, "option", 0, &index)
         != TCL_OK) {
            break;
        }

        switch (index) {
        case OPTION_CLSID:
            clsIdOpt = true;
            break;
        case OPTION_INPROC:
            clsCtx = CLSCTX_INPROC_SERVER;
            break;
        case OPTION_LOCAL:
            clsCtx = CLSCTX_LOCAL_SERVER;
            break;
        case OPTION_REMOTE:
            clsCtx = CLSCTX_REMOTE_SERVER;
            break;
        }
    }

    if (i >= objc) {
	Tcl_WrongNumArgs(
            interp,
            2,
            objv,
            "?-clsid? ?-inproc? ?-local? ?-remote? progID ?hostName?");
	return TCL_ERROR;
    }

    char *progId = Tcl_GetStringFromObj(objv[i], 0);

    char *hostName = (i + 1 < objc) ? Tcl_GetStringFromObj(objv[i + 1], 0) : 0;
    if (clsCtx == CLSCTX_REMOTE_SERVER && hostName == 0) {
        Tcl_AppendResult(
            interp, "hostname required with -remote option", NULL);
        return TCL_ERROR;
    }

    try {
        Reference *pReference;
        
        if (clsIdOpt) {
            CLSID clsid;
            if (UuidFromString(
             reinterpret_cast<unsigned char *>(progId), &clsid) != RPC_S_OK) {
	        Tcl_AppendResult(
                    interp,
                    "cannot convert to CLSID: ",
                    progId,
                    NULL);
                return TCL_ERROR;
            }
            pReference = (subCommand == GETACTIVEOBJECT)
                ? Reference::getActiveObject(clsid, 0)
                : Reference::createInstance(clsid, 0, clsCtx, hostName);

        } else {
            pReference = (subCommand == GETACTIVEOBJECT)
                ? Reference::getActiveObject(progId)
                : Reference::createInstance(progId, clsCtx, hostName);
        }

        Tcl_SetObjResult(
            interp,
            referenceHandles.newObj(interp, pReference));
    }
    catch (_com_error &e) {
        return setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}
