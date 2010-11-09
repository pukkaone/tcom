// $Id: foreachCmd.cpp,v 1.10 2002/05/31 04:03:06 cthuang Exp $
#include "Extension.h"
#include <sstream>
#include "Reference.h"
#include "Arguments.h"

// This Tcl command iterates through the elements in a COM collection.

int
Extension::foreachCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *const objv[])
{
    if (objc != 4) {
	Tcl_WrongNumArgs(
            interp, 1, objv, "varList collectionHandle script");
	return TCL_ERROR;
    }

    Tcl_Obj *pVarList = objv[1];
    Tcl_Obj *pBody = objv[3];

    Reference *pCollection = referenceHandles.find(interp, objv[2]);
    if (pCollection == 0) {
        const char *arg = Tcl_GetStringFromObj(objv[2], 0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle ", arg, NULL);
        return TCL_ERROR;
    }

    // Collections should implement a _NewEnum method which returns an object
    // that enumerates the elements.
    HRESULT hr;
    PositionalArguments arguments;
    _variant_t varResult;

    hr = pCollection->invoke(
        DISPID_NEWENUM,
        DISPATCH_METHOD | DISPATCH_PROPERTYGET,
        arguments,
        &varResult);
    if (FAILED(hr) || V_VT(&varResult) != VT_UNKNOWN) {
	Tcl_AppendResult(interp, "object is not a collection", NULL);
	return TCL_ERROR;
    }
    IUnknownPtr pUnk(V_UNKNOWN(&varResult));
    
    // Get a specific kind of enumeration.
    IEnumVARIANTPtr pEnumVARIANT;
    IEnumUnknownPtr pEnumUnknown;
    enum EnumKind { ENUM_VARIANT, ENUM_UNKNOWN };
    EnumKind enumKind;

    hr = pUnk->QueryInterface(
        IID_IEnumVARIANT, reinterpret_cast<void **>(&pEnumVARIANT));
    if (SUCCEEDED(hr)) {
        enumKind = ENUM_VARIANT;
    } else {
        hr = pUnk->QueryInterface(
            IID_IEnumUnknown, reinterpret_cast<void **>(&pEnumUnknown));
        if (SUCCEEDED(hr)) {
	    enumKind = ENUM_UNKNOWN;
        }
    }
    
    if (FAILED(hr)) {
        Tcl_AppendResult(interp,
            "Unknown enumerator type: not IEnumVARIANT or IEnumUnknown", NULL);
        return TCL_ERROR;
    }

    int completionCode;

    int varc;				// number of loop variables
    completionCode = Tcl_ListObjLength(interp, pVarList, &varc);
    if (completionCode != TCL_OK) {
        return TCL_ERROR;
    }
    if (varc < 1) {
	Tcl_AppendResult(interp, "foreach varlist is empty", NULL);
	return TCL_ERROR;
    }

    while (true) {
        // If the variable list has been converted to another kind of Tcl
        // object, convert it back to a list and refetch the pointer to its
        // element array.
        Tcl_Obj **varv;
        completionCode =
            Tcl_ListObjGetElements(interp, pVarList, &varc, &varv);
        if (completionCode != TCL_OK) {
            return TCL_ERROR;
        }

        // Assign values to all loop variables.
        int v = 0;
	for (;  v < varc;  ++v) {
            TclObject value;
	    ULONG count;

	    switch (enumKind) {
	    case ENUM_VARIANT:
		{
		    _variant_t elementVar;
		    hr = pEnumVARIANT->Next(1, &elementVar, &count);
		    if (hr == S_OK && count > 0) {
			value = TclObject(&elementVar, Type::variant(), interp);
		    }
		}
		break;

	    case ENUM_UNKNOWN:
		{
		    IUnknown *pElement;
		    hr = pEnumUnknown->Next(1, &pElement, &count);
		    if (hr == S_OK && count > 0) {
			value = referenceHandles.newObj(
			    interp, Reference::newReference(pElement));
		    }
		}
		break;
	    }

	    if (FAILED(hr)) {
		_com_issue_error(hr);
	    }
	    if (hr != S_OK || count == 0) {
		break;
	    }

	    Tcl_Obj *varValuePtr = Tcl_ObjSetVar2(
                interp, varv[v], NULL, value, TCL_LEAVE_ERR_MSG);
	    if (varValuePtr == NULL) {
		return TCL_ERROR;
	    }
	}

        if (v == 0) {
            completionCode = TCL_OK;
            break;
        }

        if (v < varc) {
            TclObject empty;

            for (; v < varc;  ++v) {
	        Tcl_Obj *varValuePtr = Tcl_ObjSetVar2(
                    interp, varv[v], NULL, empty, TCL_LEAVE_ERR_MSG);
	        if (varValuePtr == NULL) {
		    return TCL_ERROR;
	        }
            }
        }

        // Execute the script body.
        completionCode =
#if TCL_MINOR_VERSION >= 1
            Tcl_EvalObjEx(interp, pBody, 0);
#else
            Tcl_EvalObj(interp, pBody);
#endif

        if (completionCode == TCL_CONTINUE) {
            // do nothing
        } else if (completionCode == TCL_BREAK) {
            completionCode = TCL_OK;
            break;
        } else if (completionCode == TCL_ERROR) {
	    std::ostringstream oss;
            oss << "\n    (\"foreach\" body line %d)" << interp->errorLine;
            Tcl_AddObjErrorInfo(
                interp, const_cast<char *>(oss.str().c_str()), -1);
            break;
        } else if (completionCode != TCL_OK) {
            break;
        }
    }

    if (completionCode == TCL_OK) {
	Tcl_ResetResult(interp);
    }
    return completionCode;
}
