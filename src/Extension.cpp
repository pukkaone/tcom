// $Id: Extension.cpp,v 1.3 2003/04/02 22:46:51 cthuang Exp $
#pragma warning(disable: 4786)
#include "Extension.h"
#include "ComModule.h"

Extension::Extension (Tcl_Interp *interp):
    m_interp(interp),
    m_comInitialized(false)
{
    // Register new internal representation types.
    Tcl_RegisterObjType(&naType);
    Tcl_RegisterObjType(&nullType);
    Tcl_RegisterObjType(&unknownPointerType);

    // Create additional commands.
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "bind", bindCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "class", classCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "configure", configureCmd, this, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "foreach", foreachCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "import", importCmd, this, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "info", infoCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "interface", interfaceCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "method", methodCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "na", naCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "null", nullCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "object", objectCmd, this, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "outputdebug", outputdebugCmd, this, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "property", propertyCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "ref", refCmd, this, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "shortPathName", shortPathNameCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "typelib", typelibCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "typeof", typeofCmd, 0, 0);
    Tcl_CreateObjCommand(
        interp, PACKAGE_NAMESPACE "variant", variantCmd, 0, 0);

    Tcl_CallWhenDeleted(interp, interpDeleteProc, this);
    Tcl_CreateExitHandler(exitProc, this);
}

void
Extension::interpDeleteProc (ClientData clientData, Tcl_Interp *)
{
    Tcl_DeleteExitHandler(exitProc, clientData);
    delete static_cast<Extension *>(clientData);
}

void
Extension::exitProc (ClientData clientData)
{
    Extension *pExtension =
        static_cast<Extension *>(clientData);
    Tcl_DontCallWhenDeleted(pExtension->m_interp, interpDeleteProc, clientData);
    delete pExtension;
}

void
Extension::initializeCom ()
{
    if (!m_comInitialized) {
        ComModule::instance().initializeCom(m_coinitFlags);
        m_comInitialized = true;
    }
}

// This Tcl command returns the name of the argument's internal
// representation type.

int
Extension::typeofCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "value");
	return TCL_ERROR;
    }

    Tcl_ObjType *pType = objv[1]->typePtr;
    char *name = (pType == 0) ? "NULL" : pType->name;
    Tcl_SetResult(interp, name, TCL_STATIC);
    return TCL_OK;
}

// This Tcl command outputs a string to the debug stream.

int
Extension::outputdebugCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "message");
	return TCL_ERROR;
    }

    Tcl_Obj *pMessage = objv[1];
    Tcl_Obj *pWithNewLine =
        Tcl_IsShared(pMessage) ? Tcl_DuplicateObj(pMessage) : pMessage;

    Tcl_AppendToObj(pWithNewLine, "\n", 1);
    OutputDebugString(Tcl_GetStringFromObj(pWithNewLine, 0));

    if (Tcl_IsShared(pMessage)) {
        Tcl_DecrRefCount(pWithNewLine);
    }
    return TCL_OK;
}
