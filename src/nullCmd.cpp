// $Id$
#include "Extension.h"
#include <string.h>

// The string representation is the same for all objects of this type.

static char nullStringRep[] = PACKAGE_NAMESPACE "NULL";

static void
nullUpdateString (Tcl_Obj *pObj)
{
    pObj->length = sizeof(nullStringRep) - 1;
    pObj->bytes = Tcl_Alloc(pObj->length + 1);
    strcpy(pObj->bytes, nullStringRep);
}

// Do not allow conversion from other types.

static int
nullSetFromAny (Tcl_Interp *interp, Tcl_Obj *)
{
    if (interp != NULL) {
        Tcl_AppendResult(
            interp, "cannot convert to ", Extension::nullType.name, NULL);
    }
    return TCL_ERROR;
}

Tcl_ObjType Extension::nullType = {
    nullStringRep,
    NULL,
    NULL,
    nullUpdateString,
    nullSetFromAny
};

// Create a Tcl value representing a null value in SQL operations.

Tcl_Obj *
Extension::newNullObj ()
{
    Tcl_Obj *pObj = Tcl_NewObj();
    Tcl_InvalidateStringRep(pObj);
    pObj->typePtr = &nullType;
    return pObj;
}

// This Tcl command returns a Tcl value representing a null value in SQL
// operations.

int
Extension::nullCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 1) {
	Tcl_WrongNumArgs(interp, 1, objv, NULL);
	return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, newNullObj());
    return TCL_OK;
}
