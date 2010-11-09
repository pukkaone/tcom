// $Id: configureCmd.cpp 5 2005-02-16 14:57:24Z cthuang $
#pragma warning(disable: 4786)
#include "Extension.h"

// This Tcl command sets and retrieves configuration options.

int
Extension::configureCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 2) {
	Tcl_WrongNumArgs(
            interp, 1, objv, "?optionName? ?value? ?optionName value? ...");
	return TCL_ERROR;
    }

    Extension *pExtension =
        static_cast<Extension *>(clientData);

    static char *options[] = {
	"-concurrency", NULL
    };
    enum OptionEnum {
        CONCURRENCY
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "option", 0,
     &index) != TCL_OK) {
        return TCL_ERROR;
    }

    switch (index) {
    case CONCURRENCY:
        if (objc == 2) {
            // Get concurrency model.
            char *result;
            switch (pExtension->concurrencyModel()) {
            case COINIT_APARTMENTTHREADED:
                result = "apartmentthreaded";
                break;
#ifdef _WIN32_DCOM
            case COINIT_MULTITHREADED:
#else
            case 0:
#endif
                result = "multithreaded";
                break;
            default:
                result = "unknown";
            }
            Tcl_AppendResult(interp, result, NULL);

        } else if (objc == 3) {
            // Set concurrency model.
            static char *options[] = {
	        "apartmentthreaded", "multithreaded", NULL
            };
            enum OptionEnum {
                APARTMENTTHREADED, MULTITHREADED
            };

            int index;
            if (Tcl_GetIndexFromObj(interp, objv[2], options, "concurrency", 0,
             &index) != TCL_OK) {
                return TCL_ERROR;
            }

            DWORD flags;
            switch (index) {
            case APARTMENTTHREADED:
                flags = COINIT_APARTMENTTHREADED;
                break;
            case MULTITHREADED:
#ifdef _WIN32_DCOM
                flags = COINIT_MULTITHREADED;
#else
                flags = 0;
#endif
                break;
            }
            pExtension->concurrencyModel(flags);

        } else {
	    Tcl_WrongNumArgs(
                interp, 2, objv, "apartmentthreaded|multithreaded");
	    return TCL_ERROR;
        }
	return TCL_OK;
    }
    return TCL_ERROR;
}
