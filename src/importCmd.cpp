// $Id: importCmd.cpp,v 1.26 2002/05/31 04:03:06 cthuang Exp $
#pragma warning(disable: 4786)
#include "Extension.h"
#include <sstream>
#include "Reference.h"
#include "TypeLib.h"
#include "TclObject.h"

// interface currently being parsed
static Interface *s_pCurrentInterface;

// Parse method parameters from list.

static int
listObjToParameters (Tcl_Interp *interp, Tcl_Obj *pParameters, Method &method)
{
    int paramCount;
    if (Tcl_ListObjLength(interp, pParameters, &paramCount) != TCL_OK) {
        return TCL_ERROR;
    }

    for (int i = 0; i < paramCount; ++i) {
        Tcl_Obj *pParameter;
        if (Tcl_ListObjIndex(interp, pParameters, i, &pParameter)
         != TCL_OK) {
            return TCL_ERROR;
        }
        
        int paramObjc;
        Tcl_Obj **paramObjv;
        if (Tcl_ListObjGetElements(interp, pParameter, &paramObjc, &paramObjv)
         != TCL_OK) {
            return TCL_ERROR;
        }
        Parameter parameter(
            Tcl_GetStringFromObj(paramObjv[0], 0),
            Tcl_GetStringFromObj(paramObjv[1], 0),
            Tcl_GetStringFromObj(paramObjv[2], 0));
        method.addParameter(parameter);
    }

    return TCL_OK;
}

// This Tcl command defines a method.

int
Extension::methodCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 5) {
	Tcl_WrongNumArgs(interp, 1, objv, "dispid returnType name parameters");
	return TCL_ERROR;
    }
    int dispid;
    if (Tcl_GetIntFromObj(interp, objv[1], &dispid) != TCL_OK) {
        return TCL_ERROR;
    }
    char *returnType = Tcl_GetStringFromObj(objv[2], 0);
    char *name = Tcl_GetStringFromObj(objv[3], 0);

    Method method(dispid, returnType, name);
    if (listObjToParameters(interp, objv[4], method) != TCL_OK) {
        return TCL_ERROR;
    }
    s_pCurrentInterface->addMethod(method);

    return TCL_OK;
}

// This Tcl command defines a property.

int
Extension::propertyCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 5 || objc > 6) {
	Tcl_WrongNumArgs(
            interp,
            1,
            objv,
            "dispid modes type name ?parameters?");
	return TCL_ERROR;
    }
    int dispid;
    if (Tcl_GetIntFromObj(interp, objv[1], &dispid) != TCL_OK) {
        return TCL_ERROR;
    }
    char *modes = Tcl_GetStringFromObj(objv[2], 0);
    char *type = Tcl_GetStringFromObj(objv[3], 0);
    char *name = Tcl_GetStringFromObj(objv[4], 0);

    Property property(dispid, modes, type, name);
    if (objc == 6) {
        if (listObjToParameters(interp, objv[5], property) != TCL_OK) {
            return TCL_ERROR;
        }
    }
    s_pCurrentInterface->addProperty(property);

    return TCL_OK;
}

// This Tcl command queries an interface pointer for a specific interface
// and returns a new interface pointer handle.

static int
interfaceObjCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc != 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "object");
	return TCL_ERROR;
    }

    Reference *pSrcRef = Extension::referenceHandles.find(interp, objv[1]);
    if (pSrcRef == 0) {
        return TCL_ERROR;
    }

    const Interface *pInterface =
        reinterpret_cast<const Interface *>(clientData);
    try {
        Tcl_SetObjResult(
            interp,
            Extension::referenceHandles.newObj(
                interp,
                Reference::newReference(pSrcRef->unknown(), pInterface)));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}

// This Tcl command defines an interface.

int
Extension::interfaceCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 3 || objc > 4) {
	Tcl_WrongNumArgs(interp, 1, objv, "name iid ?body?");
	return TCL_ERROR;
    }
    char *name = Tcl_GetStringFromObj(objv[1], 0);
    char *iidStr = Tcl_GetStringFromObj(objv[2], 0);

    IID iid;
    if (UuidFromString(reinterpret_cast<unsigned char *>(iidStr), &iid)
     != RPC_S_OK) {
	Tcl_AppendResult(interp, "cannot convert to IID: ", iidStr, NULL);
        return TCL_ERROR;
    }

    Interface *pInterface;
    if (objc == 4) {
        pInterface =
            InterfaceManager::instance().newInterface(iid, name);

        s_pCurrentInterface = pInterface;

        int completionCode =
#if TCL_MINOR_VERSION >= 1
            Tcl_EvalObjEx(interp, objv[3], TCL_EVAL_GLOBAL);
#else
            Tcl_GlobalEvalObj(interp, objv[3]);
#endif

        if (completionCode != TCL_OK) {
            return TCL_ERROR;
        }
    } else {
        pInterface = const_cast<Interface *>(
            InterfaceManager::instance().find(iid));
        if (pInterface == 0) {
	    Tcl_AppendResult(interp, "unknown IID ", iidStr, NULL);
            return TCL_ERROR;
        }
    }

    Tcl_CreateObjCommand(
        interp,
        name,
        interfaceObjCmd,
	reinterpret_cast<ClientData>(pInterface),
        (Tcl_CmdDeleteProc *)0);
    return TCL_OK;
}

// This Tcl command creates an instance of a COM class and returns an
// interface pointer handle.

static int
classObjCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc > 4) {
	Tcl_WrongNumArgs(
            interp,
            1,
            objv,
            "?-inproc? ?-local? ?-remote? ?-withevents servant? ?hostName?");
	return TCL_ERROR;
    }

    DWORD clsCtx = CLSCTX_SERVER;
    bool withEvents = false;
    TclObject servant;

    int i = 1;
    for (; i < objc; ++i) {
        static char *options[] = {
	    "-inproc", "-local", "-remote", "-withevents", NULL
        };
        enum OptionEnum {
            OPTION_INPROC, OPTION_LOCAL, OPTION_REMOTE, OPTION_WITHEVENTS
        };

        int index;
        if (Tcl_GetIndexFromObj(NULL, objv[i], options, "option", 0, &index)
         != TCL_OK) {
            break;
        }

        switch (index) {
        case OPTION_INPROC:
            clsCtx = CLSCTX_INPROC_SERVER;
            break;
        case OPTION_LOCAL:
            clsCtx = CLSCTX_LOCAL_SERVER;
            break;
        case OPTION_REMOTE:
            clsCtx = CLSCTX_REMOTE_SERVER;
            break;
        case OPTION_WITHEVENTS:
            if (i + 1 < objc) {
                withEvents = true;
                servant = objv[++i];
            }
            break;
        }
    }

    char *hostName = (i < objc) ? Tcl_GetStringFromObj(objv[i], 0) : 0;
    if (clsCtx == CLSCTX_REMOTE_SERVER && hostName == 0) {
        Tcl_AppendResult(
            interp, "hostname required with -remote option", NULL);
        return TCL_ERROR;
    }

    Class *pClass = reinterpret_cast<Class *>(clientData);
    try {
        Reference *pRef = Reference::createInstance(
            pClass->clsid(),
            pClass->defaultInterface(),
            clsCtx,
            hostName);

        if (withEvents) {
            if (pClass->sourceInterface() == 0) {
        	Tcl_AppendResult(
                    interp, "no default event source", NULL);
                return TCL_ERROR;
            }
            pRef->advise(interp, *(pClass->sourceInterface()), servant);
        }

        Tcl_SetObjResult(
            interp,
            Extension::referenceHandles.newObj(interp, pRef));
    }
    catch (_com_error &e) {
        return Extension::setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}

static void
classCmdDeleteProc (ClientData clientData)
{
    delete reinterpret_cast<Class *>(clientData);
}

// This Tcl command defines a COM class.

int
Extension::classCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 4 || objc > 5) {
	Tcl_WrongNumArgs(interp, 1, objv, "name clsid defaultIID ?sourceIID?");
	return TCL_ERROR;
    }
    char *name = Tcl_GetStringFromObj(objv[1], 0);
    char *clsidStr = Tcl_GetStringFromObj(objv[2], 0);
    char *defaultStr = Tcl_GetStringFromObj(objv[3], 0);
    char *sourceStr = (objc == 5) ? Tcl_GetStringFromObj(objv[4], 0) : 0;

    // Convert CLSID.
    CLSID clsid;
    if (UuidFromString(reinterpret_cast<unsigned char *>(clsidStr), &clsid)
     != RPC_S_OK) {
	Tcl_AppendResult(interp, "cannot convert to CLSID: ", clsidStr, NULL);
        return TCL_ERROR;
    }

    // Convert default IID.
    IID iid;

    if (UuidFromString(reinterpret_cast<unsigned char *>(defaultStr), &iid)
     != RPC_S_OK) {
	Tcl_AppendResult(interp, "cannot convert to IID: ", defaultStr, NULL);
        return TCL_ERROR;
    }

    const Interface *pDefaultInterface = InterfaceManager::instance().find(iid);
    if (pDefaultInterface == 0) {
	Tcl_AppendResult(interp, "unknown interface ", defaultStr, NULL);
        return TCL_ERROR;
    }

    // Convert source IID.
    const Interface *pSourceInterface;
    if (sourceStr != 0) {
        if (UuidFromString(reinterpret_cast<unsigned char *>(sourceStr), &iid)
         != RPC_S_OK) {
	    Tcl_AppendResult(interp, "cannot convert to IID: ", sourceStr, NULL);
            return TCL_ERROR;
        }

        pSourceInterface = InterfaceManager::instance().find(iid);
        if (pSourceInterface == 0) {
	    Tcl_AppendResult(interp, "unknown interface ", sourceStr, NULL);
            return TCL_ERROR;
        }
    } else {
        pSourceInterface = 0;
    }

    Tcl_CreateObjCommand(
        interp,
        name,
        classObjCmd,
	new Class(name, clsid, pDefaultInterface, pSourceInterface),
        classCmdDeleteProc);
    return TCL_OK;
}

const Class *
Extension::findClassByCmdName (Tcl_Interp *interp, Tcl_Obj *pName)
{
    char *nameStr = Tcl_GetStringFromObj(pName, 0);

    Tcl_CmdInfo cmdInfo;
    if (Tcl_GetCommandInfo(interp, nameStr, &cmdInfo) == 0) {
        return 0;
    }
    
    if (cmdInfo.objProc == classObjCmd) {
        return static_cast<Class *>(cmdInfo.objClientData);
    }
    return 0;
}

const Interface *
Extension::findInterfaceByCmdName (Tcl_Interp *interp, Tcl_Obj *pNameObj)
{
    char *pName = Tcl_GetStringFromObj(pNameObj, 0);

    // Check if it's the name of an interface.
    Tcl_CmdInfo cmdInfo;
    if (Tcl_GetCommandInfo(interp, pName, &cmdInfo) == 0) {
        return 0;
    }

    if (cmdInfo.objProc == interfaceObjCmd) {
        return reinterpret_cast<const Interface *>(cmdInfo.objClientData);
    }

    return 0;
}

// This Tcl command reads interface and class information from a type library
// and creates Tcl commands for accessing that information.

int
Extension::importCmd (
    ClientData clientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "typeLibrary ?namespace?");
	return TCL_ERROR;
    }

    Extension *pExtension =
        static_cast<Extension *>(clientData);
    pExtension->initializeCom();

    char *typeLibName = Tcl_GetStringFromObj(objv[1], 0);

    try {
        TypeLib *pTypeLib = TypeLib::load(typeLibName);

        // Create the new Tcl commands in a namespace named after the type
        // library, unless a specific namespace was given.
        std::string namesp;
        if (objc > 2) {
	    namesp = Tcl_GetStringFromObj(objv[2], 0);
	} else {
	    namesp = pTypeLib->name();
	}
        std::string fullyQualifiedNamespace = "::" + namesp;

        std::ostringstream script;
        script << "namespace eval " << fullyQualifiedNamespace << " {"
            << std::endl;

        // Export interface commands.
        const TypeLib::Interfaces &interfaces = pTypeLib->interfaces();
        TypeLib::Interfaces::const_iterator pInterface;
        for (pInterface = interfaces.begin(); pInterface != interfaces.end();
         ++pInterface) {
            script << "namespace export " << (*pInterface)->name() << std::endl;
        }

        // Export class commands.
        const TypeLib::Classes &classes = pTypeLib->classes();
        TypeLib::Classes::const_iterator pClass;
        for (pClass = classes.begin(); pClass != classes.end(); ++pClass) {
            script << "namespace export " << pClass->name() << std::endl;
        }

        // Generate IID and CLSID constants.
        script << "array set __uuidof {" << std::endl;

        for (pInterface = interfaces.begin(); pInterface != interfaces.end();
         ++pInterface) {
            script << (*pInterface)->name() << ' ' << (*pInterface)->iidString()
                << std::endl;
        }

        for (pClass = classes.begin(); pClass != classes.end(); ++pClass) {
            script << pClass->name() << ' ' << pClass->clsidString()
                << std::endl;
        }

        script << '}' << std::endl;     // end of array set

        // Generate enumerations.
        const TypeLib::Enums &enums = pTypeLib->enums();
        for (TypeLib::Enums::const_iterator pEnum = enums.begin();
         pEnum != enums.end(); ++pEnum) {
            script << "array set " << pEnum->name() << " {" << std::endl;

            for (Enum::const_iterator p = pEnum->begin(); p != pEnum->end();
             ++p) {
                script << p->first << ' ' << p->second << std::endl;
            }

            script << '}' << std::endl;     // end of array set
        }

        script << '}' << std::endl;     // end of namespace

#if TCL_MINOR_VERSION >= 1
        Tcl_EvalEx(
            interp,
            const_cast<char *>(script.str().c_str()),
            -1,
            TCL_EVAL_GLOBAL);
#else
        Tcl_Eval(interp, const_cast<char *>(script.str().c_str()));
#endif

        // Create interface commands.
        for (pInterface = interfaces.begin(); pInterface != interfaces.end();
         ++pInterface) {
            std::string fullyQualifiedName =
                fullyQualifiedNamespace + "::" + (*pInterface)->name();

            Tcl_CreateObjCommand(
                interp,
                const_cast<char *>(fullyQualifiedName.c_str()),
                interfaceObjCmd,
	        const_cast<Interface *>(*pInterface),
                0);
        }

        // Create class commands.
        for (pClass = classes.begin(); pClass != classes.end(); ++pClass) {
            std::string fullyQualifiedName =
                fullyQualifiedNamespace + "::" + pClass->name();

            Tcl_CreateObjCommand(
                interp,
                const_cast<char *>(fullyQualifiedName.c_str()),
                classObjCmd,
	        new Class(*pClass),
                classCmdDeleteProc);
        }

        // Return the library name.
        Tcl_AppendResult(interp, pTypeLib->name().c_str(), NULL);

        delete pTypeLib;
    }
    catch (_com_error &e) {
        return setComErrorResult(interp, e, __FILE__, __LINE__);
    }

    return TCL_OK;
}
