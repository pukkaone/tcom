// $Id$
#include "Extension.h"
#include "TclObject.h"
#include "Reference.h"

static int interfaceObjCmd(ClientData, Tcl_Interp *, int, Tcl_Obj *CONST []);
HandleSupport<InterfaceHolder> Extension::interfaceHolderHandles(interfaceObjCmd);

// Convert type description to a Tcl list representation.

static TclObject
typeToListObj (const Type &type)
{
    TclObject list(Tcl_NewListObj(0, 0));

    list.lappend(
        Tcl_NewStringObj(const_cast<char *>(type.toString().c_str()), -1));

    for (unsigned i = 0; i < type.pointerCount(); ++i) {
        list.lappend(Tcl_NewStringObj("*", -1));
    }

    return list;
}

// Convert parameter description to a Tcl list representation.

static TclObject
parameterToListObj (const Parameter &parameter)
{
    TclObject list(Tcl_NewListObj(0, 0));

    // Put parameter passing modes.
    TclObject modes(Tcl_NewListObj(0, 0));

    if (parameter.flags() & PARAMFLAG_FIN) {
        modes.lappend(Tcl_NewStringObj("in", -1));
    }
    if (parameter.flags() & PARAMFLAG_FOUT) {
        modes.lappend(Tcl_NewStringObj("out", -1));
    }
    list.lappend(modes);

    // Put parameter type.
    list.lappend(typeToListObj(parameter.type()));

    // Put parameter name.
    list.lappend(
        Tcl_NewStringObj(const_cast<char *>(parameter.name().c_str()), -1));

    return list;
}

// Convert method description to a Tcl list representation.

static TclObject
methodToListObj (const Method &method)
{
    TclObject list(Tcl_NewListObj(0, 0));

    // Put member id.
    list.lappend(Tcl_NewIntObj(method.memberid()));

    // Put return type.
    list.lappend(typeToListObj(method.type()));

    // Put method name.
    list.lappend(
        Tcl_NewStringObj(const_cast<char *>(method.name().c_str()), -1));

    // Put parameters.
    TclObject parameterList(Tcl_NewListObj(0, 0));

    const Method::Parameters &parameters = method.parameters();
    for (Method::Parameters::const_iterator p = parameters.begin();
     p != parameters.end(); ++p) {
        parameterList.lappend(parameterToListObj(*p));
    }

    list.lappend(parameterList);

    return list;
}

// Implement interface descriptor object command.

static int
interfaceObjCmd (
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
        "iid", "methods", "name", "properties", NULL
    };
    enum MethodEnum {
        IID, METHODS, NAME, PROPERTIES
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "method", 0, &index)
     != TCL_OK) {
	return TCL_ERROR;
    }

    const InterfaceHolder *pHolder =
        reinterpret_cast<const InterfaceHolder *>(clientData);
    const Interface *pInterface = pHolder->interfaceDesc();

    switch (index) {
    case IID:
        // Get IID.
        Tcl_AppendResult(interp, pInterface->iidString().c_str(), NULL);
	return TCL_OK;

    case METHODS:
        // Get method descriptions.
        if (objc != 2) {
            Tcl_WrongNumArgs(interp, 2, objv, NULL);
	    return TCL_ERROR;
        } else {
            TclObject methodList(Tcl_NewListObj(0, 0));

            const Interface::Methods &methods = pInterface->methods();
            for (Interface::Methods::const_iterator p = methods.begin();
             p != methods.end(); ++p) {
                methodList.lappend(methodToListObj(*p));
            }

            Tcl_SetObjResult(interp, methodList);
        }
	return TCL_OK;

    case NAME:
        // Get interface name.
        Tcl_AppendResult(interp, pInterface->name().c_str(), NULL);
	return TCL_OK;

    case PROPERTIES:
        // Get property descriptions.
        // Returns a list where each element is a list consisting of
        // { dispatchID {modes} {type} name {parameters} }
        if (objc != 2) {
	    Tcl_WrongNumArgs(interp, 2, objv, NULL);
            return TCL_ERROR;
        } else {
            TclObject propertyList(Tcl_NewListObj(0, 0));

            const Interface::Properties &properties = pInterface->properties();
            for (Interface::Properties::const_iterator p = properties.begin();
             p != properties.end(); ++p) {
                TclObject property(Tcl_NewListObj(0, 0));

                // Set dispatch ID.
                property.lappend(Tcl_NewIntObj(p->memberid()));

                // Set read/write modes.
                TclObject modes(Tcl_NewListObj(0, 0));

                if (!p->readOnly()) {
                    modes.lappend(Tcl_NewStringObj("in", -1));
                }
                modes.lappend(Tcl_NewStringObj("out", -1));

                property.lappend(modes);

                // Set property type.
                property.lappend(typeToListObj(p->type()));

                // Put property name.
                property.lappend(Tcl_NewStringObj(
                    const_cast<char *>(p->name().c_str()), -1));

                // Put parameters.
                const Property::Parameters &parameters = p->parameters();
                if (parameters.size() > 0) {
                    TclObject parameterList(Tcl_NewListObj(0, 0));

                    for (Property::Parameters::const_iterator q =
                     parameters.begin(); q != parameters.end(); ++q) {
                        parameterList.lappend(parameterToListObj(*q));
                    }

                    property.lappend(parameterList);
                }

                propertyList.lappend(property);
            }

            Tcl_SetObjResult(interp, propertyList);
        }
	return TCL_OK;
    }

    return TCL_ERROR;
}

// This Tcl command returns descriptions of object interfaces.

int
Extension::infoCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 2) {
	Tcl_WrongNumArgs(interp, 1, objv, "subcommand ?arg ...?");
	return TCL_ERROR;
    }

    static char *options[] = {
	"interface", NULL
    };
    enum SubCommandEnum {
        INTERFACE
    };

    int index;
    if (Tcl_GetIndexFromObj(interp, objv[1], options, "subcommand", 0,
     &index) != TCL_OK) {
        return TCL_ERROR;
    }

    switch (index) {
    case INTERFACE:
        // Create interface description object.
        {
            if (objc != 3) {
	        Tcl_WrongNumArgs(interp, 2, objv, "handle");
	        return TCL_ERROR;
            }
            const Interface *pInterface;
            Tcl_Obj *pObj = objv[2];

            Reference *pRef = Extension::referenceHandles.find(interp, pObj);
            if (pRef != 0) {
                pInterface = pRef->interfaceDesc();
            } else {
                pInterface = Extension::findInterfaceByCmdName(interp, pObj);
                if (pInterface == 0) {
                    const Class *pClass =
                        Extension::findClassByCmdName(interp, pObj);
                    if (pClass != 0) {
                        pInterface = pClass->defaultInterface();
                    }
                }
            }

            if (pInterface == 0) {
                Tcl_AppendResult(interp, "cannot get type information", NULL);
                return TCL_ERROR;
            }

            InterfaceHolder *pHolder = new InterfaceHolder(pInterface);
            Tcl_Obj *pHandle = 
                interfaceHolderHandles.newObj(interp, pHolder);
            Tcl_SetObjResult(interp, pHandle);
        }
	return TCL_OK;
    }
    return TCL_ERROR;
}
