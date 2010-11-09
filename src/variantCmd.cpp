// $Id: variantCmd.cpp,v 1.1 2003/05/29 03:33:08 cthuang Exp $
#include "Extension.h"
#include <string.h>

static void
variantFreeInternalRep (Tcl_Obj *pObj)
{
    delete static_cast<_variant_t *>(pObj->internalRep.otherValuePtr);
}

static void
variantDuplicateInternalRep (Tcl_Obj *pSrc, Tcl_Obj *pDup)
{
    pDup->typePtr = &Extension::variantType;
    pDup->internalRep.otherValuePtr = new _variant_t(
        static_cast<_variant_t *>(pSrc->internalRep.otherValuePtr));
}

static void
variantUpdateString (Tcl_Obj *pObj)
{
    try {
        _bstr_t bstr(
            static_cast<_variant_t *>(pObj->internalRep.otherValuePtr));
        const char *stringRep = bstr;
        pObj->length = strlen(stringRep);
        pObj->bytes = Tcl_Alloc(pObj->length + 1);
        strcpy(pObj->bytes, stringRep);
    }
    catch (_com_error &) {
        pObj->length = 0;
        pObj->bytes = Tcl_Alloc(1);
        pObj->bytes[0] = '\0';
    }
}

static int
variantSetFromAny (Tcl_Interp *interp, Tcl_Obj *pObj)
{
    const char *stringRep = Tcl_GetStringFromObj(pObj, 0);

    Tcl_ObjType *pOldType = pObj->typePtr;
    if (pOldType != NULL && pOldType->freeIntRepProc != NULL) {
	pOldType->freeIntRepProc(pObj);
    }

    pObj->typePtr = &Extension::variantType;
    pObj->internalRep.otherValuePtr = new _variant_t(stringRep);
    return TCL_OK;
}

Tcl_ObjType Extension::variantType = {
    PACKAGE_NAMESPACE "VARIANT",
    variantFreeInternalRep,
    variantDuplicateInternalRep,
    variantUpdateString,
    variantSetFromAny
};

// Create a Tcl value representing a VARIANT.

Tcl_Obj *
Extension::newVariantObj (_variant_t *pVar)
{
    Tcl_Obj *pObj = Tcl_NewObj();
    Tcl_InvalidateStringRep(pObj);
    pObj->typePtr = &variantType;
    pObj->internalRep.otherValuePtr = pVar;
    return pObj;
}

// This Tcl command returns a Tcl value representing a VARIANT.

int
Extension::variantCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 2 || objc > 3) {
	Tcl_WrongNumArgs(interp, 1, objv, "type ?value?");
	return TCL_ERROR;
    }

    static char *types[] = {
        "empty",
        "null",
        "i2",
        "i4",
        "r4",
        "r8",
        "cy",
        "date",
        "bstr",
        "dispatch",
        "error",
        "bool",
        "variant",
        "unknown",
        "decimal",
        "record",
        "i1",
        "ui1",
        "ui2",
        "ui4",
        "i8",   // VT_I8 and VT_UI8 actually are not valid VARIANT types.
        "ui8",
        "int",
        "uint",
        NULL
    };

    int vt;
    if (Tcl_GetIndexFromObj(NULL, objv[1], types, "type", 0, &vt) != TCL_OK) {
        if (Tcl_GetIntFromObj(interp, objv[1], &vt) != TCL_OK) {
            return TCL_ERROR;
        }
    }

    _variant_t *pVar = new _variant_t;
    switch (vt) {
    case VT_DISPATCH:
        V_VT(pVar) = vt;
        V_DISPATCH(pVar) = 0;
        break;

    case VT_UNKNOWN:
        V_VT(pVar) = vt;
        V_UNKNOWN(pVar) = 0;
        break;
    }

    try {
        if (objc == 3) {
            *pVar = Tcl_GetStringFromObj(objv[2], 0);
        }

        if (vt == VT_DATE) {
            pVar->ChangeType(VT_R8);
        }

        pVar->ChangeType(vt);
    }
    catch (_com_error &e) {
        return setComErrorResult(interp, e, __FILE__, __LINE__);
    }

    Tcl_SetObjResult(interp, newVariantObj(pVar));
    return TCL_OK;
}
