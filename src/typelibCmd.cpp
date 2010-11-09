// $Id: typelibCmd.cpp 5 2005-02-16 14:57:24Z cthuang $
#pragma warning(disable: 4786)
#include "Extension.h"
#include "TypeLib.h"

static int typeLibObjCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
HandleSupport<TypeLib> Extension::typeLibHandles(typeLibObjCmd);

// Implement type library object command.

static int
typeLibObjCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "method ?arg ...?");
	return TCL_ERROR;
    }

    static char *options[] = {
	"class", "documentation", "enum", "interface", "libid", "name", "version", NULL
    };
    enum MethodEnum {
        CLASS, DOCUMENTATION, ENUM, INTERFACE, LIBID, NAME, VERSION
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "method", 0, &index)
     != TCL_OK) {
	return TCL_ERROR;
    }

    const TypeLib *pTypeLib = reinterpret_cast<const TypeLib *>(clientData);

    switch (index) {
    case CLASS:
        if (objc == 2) {
            // Get list of class names.
            const TypeLib::Classes &classes = pTypeLib->classes();
            for (TypeLib::Classes::const_iterator p = classes.begin();
            p != classes.end(); ++p) {
                Tcl_AppendElement(
                    interp,
                    const_cast<char *>(p->name().c_str()));
            }

        } else if (objc == 3) {
            // Get class description.
            char *className = Tcl_GetStringFromObj(objv[2], 0);

            const Class *pClass = pTypeLib->findClass(className);
            if (pClass == 0) {
                Tcl_AppendResult(
                    interp,
                    "class not found: ",
                    className,
                    NULL);
	        return TCL_ERROR;
            }

            // Append CLSID.
            Tcl_AppendElement(
                interp,
                const_cast<char *>(pClass->clsidString().c_str()));

            // Append name of default interface.
            Tcl_AppendElement(
                interp,
                const_cast<char *>(
                    pClass->defaultInterface()->name().c_str()));

            // Append name of source interface.
            if (pClass->sourceInterface() != 0) {
                Tcl_AppendElement(
                    interp,
                    const_cast<char *>(
                        pClass->sourceInterface()->name().c_str()));
            }

        } else {

	    Tcl_WrongNumArgs(interp, 2, objv, "?className?");
	    return TCL_ERROR;
        }
	return TCL_OK;

    case DOCUMENTATION:
        // Get type library documentation.
        Tcl_AppendResult(interp, pTypeLib->documentation().c_str(), NULL);
	return TCL_OK;

    case ENUM:
        if (objc == 2) {
            // Return list of enumerations.
            const TypeLib::Enums &enums = pTypeLib->enums();
            for (TypeLib::Enums::const_iterator p = enums.begin();
            p != enums.end(); ++p) {
                Tcl_AppendElement(
                    interp,
                    const_cast<char *>(p->name().c_str()));
            }

        } else {
            // Get the named enumeration.
            char *name = Tcl_GetStringFromObj(objv[2], 0);
            const Enum *pEnum = pTypeLib->findEnum(name);
            if (pEnum == 0) {
                Tcl_AppendResult(interp, "unknown enumeration ", name, NULL);
                return TCL_ERROR;
            }

            if (objc == 3) {
                // Return list of enumerator name/value pairs.
                for (Enum::const_iterator p = pEnum->begin(); p != pEnum->end();
                ++p) {
                    Tcl_AppendElement(interp,
                        const_cast<char *>(p->first.c_str()));
                    Tcl_AppendElement(interp,
                        const_cast<char *>(p->second.c_str()));
                }

            } else if (objc == 4) {
                // Return value of named enumerator.
                char *name = Tcl_GetStringFromObj(objv[3], 0);

                Enum::const_iterator p = pEnum->find(name);
                if (p == pEnum->end()) {
                    Tcl_AppendResult(interp, "unknown enumerator ", name, NULL);
                    return TCL_ERROR;
                }

                Tcl_AppendElement(interp,
                    const_cast<char *>(p->second.c_str()));

            } else {
	        Tcl_WrongNumArgs(
                    interp,
                    2,
                    objv,
                    "?enumerationName? ?enumeratorName?");
                return TCL_ERROR;
            }
        }
        return TCL_OK;

    case INTERFACE:
        if (objc == 2) {
            // Get list of interface names.
            const TypeLib::Interfaces &interfaces = pTypeLib->interfaces();
            for (TypeLib::Interfaces::const_iterator p = interfaces.begin();
            p != interfaces.end(); ++p) {
                Tcl_AppendElement(
                    interp,
                    const_cast<char *>((*p)->name().c_str()));
            }

        } else if (objc == 3) {
            // Get interface description.
            char *name = Tcl_GetStringFromObj(objv[2], 0);

            const Interface *pInterface = pTypeLib->findInterface(name);
            if (pInterface == 0) {
                Tcl_AppendResult(
                    interp,
                    "interface not found: ",
                    name,
                    NULL);
	        return TCL_ERROR;
            }

            InterfaceHolder *pHolder = new InterfaceHolder(pInterface);
            Tcl_Obj *pHandle =
                Extension::interfaceHolderHandles.newObj(interp, pHolder);
            Tcl_SetObjResult(interp, pHandle);

        } else {
            Tcl_WrongNumArgs(interp, 2, objv, "?interfaceName?");
	    return TCL_ERROR;
        }
	return TCL_OK;

    case LIBID:
        Tcl_AppendResult(interp, pTypeLib->libidString().c_str(), NULL);
	return TCL_OK;

    case NAME:
        Tcl_AppendResult(interp, pTypeLib->name().c_str(), NULL);
	return TCL_OK;

    case VERSION:
        Tcl_AppendResult(interp, pTypeLib->version().c_str(), NULL);
	return TCL_OK;
    }

    return TCL_ERROR;
}

// This Tcl command loads a type library.

int
Extension::typelibCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 3) {
	Tcl_WrongNumArgs(interp, 1, objv, "option typeLibrary");
	return TCL_ERROR;
    }

    static char *options[] = {
	"load", "register", "unregister", NULL
    };
    enum SubCommandEnum {
        LOAD, REGISTER, UNREGISTER
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "subcommand", 0, &index)
     != TCL_OK) {
	return TCL_ERROR;
    }

    char *typeLibName = Tcl_GetStringFromObj(objv[2], 0);

    try {
        TypeLib *pTypeLib;

        switch (index) {
        case LOAD:
            pTypeLib = TypeLib::load(typeLibName);
            Tcl_SetObjResult(
                interp,
                typeLibHandles.newObj(interp, pTypeLib));
            break;

        case REGISTER:
            pTypeLib = TypeLib::load(typeLibName, true);
            delete pTypeLib;
            break;

        case UNREGISTER:
            TypeLib::unregister(typeLibName);
            break;
        }
    }
    catch (_com_error &e) {
        return setComErrorResult(interp, e, __FILE__, __LINE__);
    }

    return TCL_OK;
}
