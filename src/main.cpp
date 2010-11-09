// $Id: main.cpp 5 2005-02-16 14:57:24Z cthuang $
#pragma warning(disable: 4786)
#include "ComModule.h"
#include "Extension.h"
#include "TclObject.h"
#include "version.h"
#include "tclRunTime.h"

/*
 *	This procedure performs application-specific initialization.
 *	Most applications, especially those that incorporate additional
 *	packages, will have their own version of this procedure.
 *
 * Results:
 *	Returns a standard Tcl completion code, and leaves an error
 *	message in interp->result if an error occurs.
 *
 * Side effects:
 *	Depends on the startup script.
 */
extern "C" DLLEXPORT int
Tcom_Init (Tcl_Interp *interp)
{
#ifdef USE_TCL_STUBS
    // Stubs were introduced in Tcl 8.1.
    if (Tcl_InitStubs(interp, "8.1", 0) == NULL) {
        return TCL_ERROR;
    }
#endif

    // Get pointers to Tcl's built-in internal representation types.
    TclTypes::initialize();

    Extension *pExtension = new Extension(interp);
    pExtension->concurrencyModel(COINIT_APARTMENTTHREADED);

    // Initialize handle support.
    CmdNameType::instance();
    new HandleNameToRepMap(interp);

    return Tcl_PkgProvide(interp, PACKAGE_NAME, PACKAGE_VERSION);
}

/*
 *	This procedure initializes commands for a safe interpreter.
 *	You would leave out of this procedure any commands you deemed unsafe.
 *
 * Results:
 *	A standard Tcl result.
 *
 * Side effects:
 *	None.
 */
extern "C" DLLEXPORT int
Tcom_SafeInit (
    Tcl_Interp *interp)
{
#ifdef USE_TCL_STUBS
    // Stubs were introduced in Tcl 8.1.
    if (Tcl_InitStubs(interp, "8.1", 0) == NULL) {
        return TCL_ERROR;
    }
#endif

    return Tcl_PkgProvide(interp, PACKAGE_NAME, PACKAGE_VERSION);
}
