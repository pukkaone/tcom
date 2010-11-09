// $Id$
#pragma warning(disable: 4786)
#include "Extension.h"
#include <sstream>
#include "ComObject.h"
#include "ComObjectFactory.h"
#include "ComModule.h"

// Set the string representation to match the internal representation.

static void
unknownPointerUpdateString (Tcl_Obj *pObj)
{
    std::ostringstream oss;
    oss << "0x" << std::hex << pObj->internalRep.otherValuePtr;
    std::string stringRep(oss.str());

    pObj->length = stringRep.size();
    pObj->bytes = Tcl_Alloc(pObj->length + 1);
    stringRep.copy(pObj->bytes, pObj->length);
    pObj->bytes[pObj->length] = '\0';
}

// Set internal representation from string representation.

static int
unknownPointerSetFromAny (Tcl_Interp *interp, Tcl_Obj *)
{
    // Do not allow conversion from other types.
    if (interp != NULL) {
        Tcl_AppendResult(
            interp,
            "cannot convert to ",
            Extension::unknownPointerType.name,
            NULL);
    }
    return TCL_ERROR;
}

Tcl_ObjType Extension::unknownPointerType = {
    PACKAGE_NAMESPACE "UnknownPointer",
    NULL,
    NULL,
    unknownPointerUpdateString,
    unknownPointerSetFromAny
};

// This Tcl command registers a factory that creates COM objects.

static int
objectRegisterFactoryCmd (
    Tcl_Interp *interp,		/* Current interpreter. */
    int objc,			/* Number of arguments. */
    Tcl_Obj *CONST objv[])	/* The argument objects. */
{
    bool registerActiveOpt = false;
    bool singletonOpt = false;

    int i = 2;
    for (; i < objc; ++i) {
        static char *options[] = {
	    "-registeractive", "-singleton", NULL
        };
        enum OptionEnum {
            OPTION_REGISTERACTIVE,
            OPTION_SINGLETON
        };

        int index;
        if (Tcl_GetIndexFromObj(NULL, objv[i], options, "option", 0, &index)
         != TCL_OK) {
            break;
        }

        switch (index) {
        case OPTION_REGISTERACTIVE:
            registerActiveOpt = true;
            break;
        case OPTION_SINGLETON:
            singletonOpt = true;
            break;
        }
    }

    if (objc - i < 2 || objc - i > 3) {
	Tcl_WrongNumArgs(
            interp,
            2,
            objv,
            "?-singleton? class constructCommand ?destroyCommand?");
	return TCL_ERROR;
    }

    const Class *pClass = Extension::findClassByCmdName(interp, objv[i]);
    if (pClass == 0) {
        char *className = Tcl_GetStringFromObj(objv[i], 0);
        Tcl_AppendResult(interp, "unknown class ", className, NULL);
        return TCL_ERROR;
    }

    TclObject constructor(objv[i + 1]);

    TclObject destructor;
    if (objc - i == 3) {
        destructor = objv[i + 2];
    }

    ComObjectFactory *pFactory;
    if (singletonOpt) {
        pFactory = new SingletonObjectFactory(
            pClass->interfaces(),
            interp,
            constructor,
            destructor,
            registerActiveOpt);
    } else {
        pFactory = new ComObjectFactory(
            pClass->interfaces(),
            interp,
            constructor,
            destructor,
            registerActiveOpt);
    }
    ComModule::instance().registerFactory(pClass->clsid(), pFactory);
    return TCL_OK;
}

// Find interface description from imported interface name.
// On error, put a message in the Tcl interpreter result and return 0.

static const Interface *
findInterface (Tcl_Interp *interp, Tcl_Obj *pName)
{
    const Interface *pInterface = Extension::findInterfaceByCmdName(interp, pName);
    if (pInterface == 0) {
        char *nameStr = Tcl_GetStringFromObj(pName, 0);
        Tcl_AppendResult(
            interp, "unknown interface name: ", nameStr, NULL);
    }
    return pInterface;
}

static Tcl_Obj *
newUnknownPointer (IUnknown *pUnknown)
{
    Tcl_Obj *pObj = Tcl_NewObj();
    Tcl_InvalidateStringRep(pObj);
    pObj->typePtr = &Extension::unknownPointerType;
    pObj->internalRep.otherValuePtr = pUnknown;
    return pObj;
}

// This Tcl command creates a COM object.

static int
objectCreateCmd (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    bool registerActiveOpt = false;

    int i = 2;
    for (; i < objc; ++i) {
        static char *options[] = {
	    "-registeractive", NULL
        };
        enum OptionEnum {
            OPTION_REGISTERACTIVE
        };

        int index;
        if (Tcl_GetIndexFromObj(NULL, objv[i], options, "option", 0, &index)
         != TCL_OK) {
            break;
        }

        switch (index) {
        case OPTION_REGISTERACTIVE:
            registerActiveOpt = true;
            break;
        }
    }

    if (objc - i < 2 || objc - i > 3) {
	Tcl_WrongNumArgs(
            interp,
            2,
            objv,
            "?-registeractive? class servant ?destroyCommand?");
	return TCL_ERROR;
    }

    TclObject servant(objv[i + 1]);

    TclObject destructor;
    if (objc - i == 3) {
        destructor = objv[i + 2];
    }

    try {
        ComObject *pComObject;

        const Class *pClass = Extension::findClassByCmdName(interp, objv[i]);
        if (pClass != 0) {
            pComObject = ComObject::newInstance(
                pClass->interfaces(),
                interp,
                servant,
                destructor);

            if (registerActiveOpt) {
                pComObject->registerActiveObject(pClass->clsid());
            }
        } else {
            // Check if the argument is a list of imported interface names.
            int interfaceCount;
            Tcl_Obj **interfaceObj;
            int result = Tcl_ListObjGetElements(
                interp, objv[i], &interfaceCount, &interfaceObj);
            if (result != TCL_OK) {
                return TCL_ERROR;
            }
            if (interfaceCount < 1) {
                Tcl_AppendResult(
                    interp, "must specify at least one interface name", NULL);
                return TCL_ERROR;
            }

            Class::Interfaces interfaces;
            for (int i = 0; i < interfaceCount; ++i) {
                const Interface *pInterface =
                    findInterface(interp, interfaceObj[i]);
                if (pInterface == 0) {
                    return TCL_ERROR;
                }
                interfaces.push_back(pInterface);
            }

            pComObject = ComObject::newInstance(
                interfaces,
                interp,
                servant,
                destructor);
        }

        Tcl_SetObjResult(interp, newUnknownPointer(pComObject->unknown()));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}

// This Tcl command creates a null IUnknown pointer.

static int
objectNullCmd (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 2) {
	Tcl_WrongNumArgs(
            interp,
            2,
            objv,
            NULL);
	return TCL_ERROR;
    }

    Tcl_SetObjResult(interp, newUnknownPointer(0));
    return TCL_OK;
}

// This Tcl command provides operations for creating COM objects.

int
Extension::objectCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "subcommand ?arg ...?");
	return TCL_ERROR;
    }

    Extension *pExtension =
        static_cast<Extension *>(clientData);
    pExtension->initializeCom();

    static char *options[] = {
	"create", "null", "registerfactory", NULL
    };
    enum SubCommandEnum {
        CREATE, OBJECT_NULL, REGISTER_FACTORY
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "subcommand", 0,
     &index) != TCL_OK) {
        return TCL_ERROR;
    }

    switch (index) {
    case CREATE:
        return objectCreateCmd(interp, objc, objv);
    case OBJECT_NULL:
        return objectNullCmd(interp, objc, objv);
    case REGISTER_FACTORY:
        return objectRegisterFactoryCmd(interp, objc, objv);
    }
    return TCL_ERROR;
}
