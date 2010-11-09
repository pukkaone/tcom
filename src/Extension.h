// $Id$
#ifndef EXTENSION_H
#define EXTENSION_H

#include <comdef.h>
#include <tcl.h>
#include "tcomApi.h"
#include "HandleSupport.h"

// package name
#define PACKAGE_NAME "tcom"

// namespace where the package defines new commands
#define PACKAGE_NAMESPACE "::tcom::"

class Class;
class Interface;
class InterfaceHolder;
class Reference;
class TypeLib;

// This class implements the commands and state of an extension loaded into a
// Tcl interpreter.

class TCOM_API Extension
{
    // interpreter associated with this object
    Tcl_Interp *m_interp;

    // flags used to initialize COM
    DWORD m_coinitFlags;

    // true if COM was initialized
    bool m_comInitialized;

    static void interpDeleteProc(ClientData clientData, Tcl_Interp *interp);
    static void exitProc(ClientData clientData);

    static int bindCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int classCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int configureCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int foreachCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int importCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int infoCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int interfaceCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int methodCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int naCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int nullCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int objectCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int outputdebugCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int propertyCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int refCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int shortPathNameCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int typelibCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int typeofCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
    static int variantCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);

    // not implemented
    Extension(const Extension &);
    void operator=(const Extension &);

    ~Extension ()
    { }

public:
    Extension(Tcl_Interp *interp);

    // Set the concurrency model to be used by the current thread.
    void concurrencyModel (DWORD flags)
    { m_coinitFlags = flags; }

    // Get the concurrency model to be used by the current thread.
    DWORD concurrencyModel () const
    { return m_coinitFlags; }

    // Initialize COM if not already initialized.
    void initializeCom();

    // handle support objects
    static HandleSupport<InterfaceHolder> interfaceHolderHandles;
    static HandleSupport<Reference> referenceHandles;
    static HandleSupport<TypeLib> typeLibHandles;

    // new Tcl internal representation types
    static Tcl_ObjType naType;
    static Tcl_ObjType nullType;
    static Tcl_ObjType unknownPointerType;
    static Tcl_ObjType variantType;

    // Create a Tcl value representing a missing optional argument.
    static Tcl_Obj *newNaObj();

    // Create a Tcl value representing a SQL-style null.
    static Tcl_Obj *newNullObj();

    // Create a Tcl value representing a VARIANT.
    static Tcl_Obj *newVariantObj(_variant_t *pVar);

    // Set the Tcl result to a description of the COM error and return TCL_ERROR.
    static int setComErrorResult(
        Tcl_Interp *interp, _com_error &e, const char *file, int line);

    // Find class description by name.
    static const Class *findClassByCmdName(Tcl_Interp *interp, Tcl_Obj *pName);

    // Find interface description by name.
    static const Interface *findInterfaceByCmdName(
        Tcl_Interp *interp, Tcl_Obj *pName);
};

#endif
