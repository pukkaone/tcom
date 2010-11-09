// $Id: naCmd.cpp,v 1.7 2003/03/07 00:17:30 cthuang Exp $
#include "Extension.h"
#include <string.h>

// The string representation is the same for all objects of this type.

static char naStringRep[] = PACKAGE_NAMESPACE "NA";

static void
naUpdateString (Tcl_Obj *pObj)
{
    pObj->length = sizeof(naStringRep) - 1;
    pObj->bytes = Tcl_Alloc(pObj->length + 1);
    strcpy(pObj->bytes, naStringRep);
}

// Do not allow conversion from other types.

static int
naSetFromAny (Tcl_Interp *interp, Tcl_Obj *)
{
    if (interp != NULL) {
        Tcl_AppendResult(
            interp, "cannot convert to ", Extension::naType.name, NULL);
    }
    return TCL_ERROR;
}

Tcl_ObjType Extension::naType = {
    naStringRep,
    NULL,
    NULL,
    naUpdateString,
    naSetFromAny
};

// Create a Tcl value representing a missing optional argument.

Tcl_Obj *
Extension::newNaObj ()
{
    Tcl_Obj *pObj = Tcl_NewObj();
    Tcl_InvalidateStringRep(pObj);
    pObj->typePtr = &naType;
    return pObj;
}

// This Tcl command returns a Tcl value representing a missing optional
// argument.

int
Extension::naCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc == 1) {
        // Return a missing argument token.
        Tcl_SetObjResult(interp, newNaObj());
        return TCL_OK;
    }

    static char *options[] = {
	"ismissing", NULL
    };
    enum SubCommandEnum {
        ISMISSING
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "subcommand", 0,
     &index) != TCL_OK) {
        return TCL_ERROR;
    }

    switch (index) {
    case ISMISSING:
        // Return true if the object is a missing argument token.
        {
            if (objc != 3) {
	        Tcl_WrongNumArgs(interp, 2, objv, "object");
	        return TCL_ERROR;
            }

            Tcl_SetObjResult(
                interp,
                Tcl_NewBooleanObj(objv[2]->typePtr == &naType));
        }
	return TCL_OK;
    }
    return TCL_ERROR;
}
