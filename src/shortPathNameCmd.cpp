// $Id: shortPathNameCmd.cpp 5 2005-02-16 14:57:24Z cthuang $
#include "Extension.h"
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

// This Tcl command returns the short path form of a input path.

int
Extension::shortPathNameCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "inputPathName");
	return TCL_ERROR;
    }

    char shortPath[MAX_PATH];
    GetShortPathName(
        Tcl_GetStringFromObj(objv[1], 0), shortPath, sizeof(shortPath));
    Tcl_SetObjResult(interp, Tcl_NewStringObj(shortPath, -1));
    return TCL_OK;
}
