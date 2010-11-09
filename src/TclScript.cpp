// $Id: TclScript.cpp,v 1.12 2003/04/02 22:46:51 cthuang Exp $
#include "ActiveScriptError.h"
#include "Reference.h"
#include "TypeInfo.h"
#include "Extension.h"
#include "tclRunTime.h"

#define NAMESPACE "::TclScriptEngine::"
#define ENGINE_PACKAGE_NAME "TclScript"
#define ENGINE_PACKAGE_VERSION "1.0"

static int
getnameditemCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 3 || objc > 4) {
	Tcl_WrongNumArgs(
            interp, 1, objv, "scriptSiteHandle itemName ?subItemName?");
	return TCL_ERROR;
    }

    Reference *pRef = Extension::referenceHandles.find(interp, objv[1]);
    if (pRef == 0) {
        const char *arg = Tcl_GetStringFromObj(objv[1], 0);
        Tcl_AppendResult(interp, "invalid handle ", arg, NULL);
	return TCL_ERROR;
    }

    try {
        HRESULT hr;

        IActiveScriptSitePtr pScriptSite;
        hr = pRef->unknown()->QueryInterface(
            IID_IActiveScriptSite, reinterpret_cast<void **>(&pScriptSite));
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

        TclObject itemName(objv[2]);
        IUnknownPtr pUnknown;
        hr = pScriptSite->GetItemInfo(
            itemName.getUnicode(), SCRIPTINFO_IUNKNOWN, &pUnknown, 0);
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

        TclObject subItemName;
        if (objc == 4) {
            subItemName = objv[3];
        }

        int subItemNameLength;
        Tcl_GetStringFromObj(subItemName, &subItemNameLength);

        Reference *pNewRef;
        if (subItemNameLength == 0) {
            pNewRef = Reference::newReference(pUnknown);
        } else {
            IDispatchPtr pDispatch;
            hr = pUnknown->QueryInterface(
                IID_IDispatch, reinterpret_cast<void **>(&pDispatch));
            if (FAILED(hr)) {
                _com_issue_error(hr);
            }

            // Get the DISPID of the name.
            wchar_t *wideSubItemName = const_cast<wchar_t *>(
                subItemName.getUnicode());

            DISPID dispid;
            hr = pDispatch->GetIDsOfNames(
                IID_NULL,
                &wideSubItemName,
                1,
                LOCALE_USER_DEFAULT,
                &dispid);
            if (FAILED(hr)) {
                // If we didn't find the name, then return an empty Tcl result.
                Tcl_ResetResult(interp);
                return TCL_OK;
            }

            // Try to get the property.
            EXCEPINFO excepInfo;
            memset(&excepInfo, 0, sizeof(excepInfo));
            UINT argErr = 0;
            
            _variant_t returnValue;
            
            DISPPARAMS dispParams;
            dispParams.rgvarg            = NULL;
            dispParams.rgdispidNamedArgs = NULL;
            dispParams.cArgs             = 0;
            dispParams.cNamedArgs        = 0;
            
            hr = pDispatch->Invoke(
                dispid,
                IID_NULL,
                LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYGET,
                &dispParams,
                &returnValue,
                &excepInfo,
                &argErr);
            if (FAILED(hr)) {
                _com_issue_error(hr);
            }
            
            if (V_VT(&returnValue) != VT_DISPATCH) {
                Tcl_AppendResult(interp, "sub item is not an IDispatch", NULL);
                return TCL_ERROR;
            }

            pNewRef = Reference::newReference(V_DISPATCH(&returnValue));
        }

        Tcl_SetObjResult(
            interp, Extension::referenceHandles.newObj(interp, pNewRef));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }

    return TCL_OK;
}

static int
activescripterrorCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 7) {
	Tcl_WrongNumArgs(
            interp,
            1,
            objv,
            "hresult source description lineNum charPos sourceLineText");
	return TCL_ERROR;
    }

    TclObject hresult(objv[1]);
    TclObject source(objv[2]);
    TclObject description(objv[3]);
    TclObject lineNumber(objv[4]);
    TclObject characterPosition(objv[5]);
    TclObject sourceLineText(objv[6]);

    try {
        ActiveScriptError *pActiveScriptError = new ActiveScriptError(
            hresult.getLong(),
            source.c_str(),
            description.c_str(),
            lineNumber.getLong(),
            characterPosition.getLong(),
            sourceLineText.c_str());

        Tcl_Obj *pResult = Tcl_NewObj();
        Tcl_InvalidateStringRep(pResult);
        pResult->typePtr = &Extension::unknownPointerType;
        pResult->internalRep.otherValuePtr = pActiveScriptError;

        Tcl_SetObjResult(interp, pResult);
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }

    return TCL_OK;
}

extern "C" DLLEXPORT int
Tclscript_Init (Tcl_Interp *interp)
{
#ifdef USE_TCL_STUBS
    // Stubs were introduced in Tcl 8.1.
    if (Tcl_InitStubs(interp, "8.1", 0) == NULL) {
        return TCL_ERROR;
    }
#endif

    Tcl_CreateObjCommand(
        interp, NAMESPACE "getnameditem", getnameditemCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, NAMESPACE "activescripterror", activescripterrorCmd, 0, 0);

    return Tcl_PkgProvide(interp, ENGINE_PACKAGE_NAME, ENGINE_PACKAGE_VERSION);
}

extern "C" DLLEXPORT int
Tclscript_SafeInit (Tcl_Interp *interp)
{
#ifdef USE_TCL_STUBS
    // Stubs were introduced in Tcl 8.1.
    if (Tcl_InitStubs(interp, "8.1", 0) == NULL) {
        return TCL_ERROR;
    }
#endif

    return Tcl_PkgProvide(interp, ENGINE_PACKAGE_NAME, ENGINE_PACKAGE_VERSION);
}
