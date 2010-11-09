// $Id: Arguments.cpp 20 2005-05-04 16:53:05Z cthuang $
#include "Arguments.h"
#include "Extension.h"
#include "TclObject.h"

Arguments::Arguments ():
    m_args(0)
{
    m_dispParams.rgvarg = NULL;
    m_dispParams.rgdispidNamedArgs = NULL;
    m_dispParams.cArgs = 0;
    m_dispParams.cNamedArgs = 0;
}

Arguments::~Arguments ()
{
    delete[] m_args;
}


TypedArguments::TypedArguments ():
    m_outValues(0)
{ }

TypedArguments::~TypedArguments ()
{
    delete[] m_outValues;
}

int
TypedArguments::initArgument (
    Tcl_Interp *interp,
    Tcl_Obj *pObj,
    int argIndex,
    const Parameter &parameter)
{
    TclObject argument(pObj);
    VARTYPE vt = parameter.type().vartype();

    if (pObj->typePtr == &Extension::naType) {
        // This variant indicates a missing optional argument.
        m_args[argIndex] = vtMissing;

    } else if (parameter.type().pointerCount() > 0) {
        // The argument is passed by reference.

        switch (vt) {
        case VT_INT:
            // IDispatch::Invoke returns DISP_E_TYPEMISMATCH on
            // VT_INT | VT_BYREF parameters.
            vt = VT_I4;
            break;

        case VT_UINT:
            // IDispatch::Invoke returns DISP_E_TYPEMISMATCH on
            // VT_UINT | VT_BYREF parameters.
            vt = VT_UI4;
            break;

        case VT_USERDEFINED:
            // Assume user defined types derive from IUnknown.
            vt = VT_UNKNOWN;
            break;
        }

        if (vt == VT_SAFEARRAY) {
            m_args[argIndex].vt = VT_BYREF | VT_ARRAY |
                parameter.type().elementType().vartype();
        } else {
            m_args[argIndex].vt = VT_BYREF | vt;
        }

        if (vt == VT_VARIANT || vt == VT_DECIMAL) {
            // Set a pointer to out variant.
            m_args[argIndex].byref = &m_outValues[argIndex];
        } else {
            // Set a pointer to variant data value.
            m_args[argIndex].byref = &m_outValues[argIndex].byref;
        }

        if (parameter.flags() & PARAMFLAG_FIN) {
            if (parameter.flags() & PARAMFLAG_FOUT) {
                // Set the value for an in/out parameter.
                Tcl_Obj *pValue = Tcl_ObjGetVar2(
                    interp, pObj, NULL, TCL_LEAVE_ERR_MSG);
                if (pValue == 0) {
                    return TCL_ERROR;
                }

                TclObject value(pValue);

                // If the argument is an interface pointer, increment its
                // reference count because the _variant_t destructor will
                // release it.
                value.toNativeValue(
                    &m_outValues[argIndex], parameter.type(), interp, true);
            } else {
                // If the argument is an interface pointer, increment its
                // reference count because the _variant_t destructor will
                // release it.
                argument.toNativeValue(
                    &m_outValues[argIndex], parameter.type(), interp, true);
            }
        } else {
            if (vt == VT_UNKNOWN) {
                m_outValues[argIndex].vt = vt;
                m_outValues[argIndex].punkVal = NULL;
            } else if (vt == VT_DISPATCH) {
                m_outValues[argIndex].vt = vt;
                m_outValues[argIndex].pdispVal = NULL;
            } else if (vt == VT_SAFEARRAY) {
                VARTYPE elementType = parameter.type().elementType().vartype();
                m_outValues[argIndex].vt = VT_ARRAY | elementType;
                m_outValues[argIndex].parray =
                    SafeArrayCreateVector(elementType, 0, 1);
            } else if (vt != VT_VARIANT) {
                m_outValues[argIndex].ChangeType(vt);
            }
        }

    } else {
        // If the argument is an interface pointer, increment its reference
        // count because the _variant_t destructor will release it.
        argument.toNativeValue(
            &m_args[argIndex], parameter.type(), interp, true);
    }

    return TCL_OK;
}

void
TypedArguments::storeOutValues (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[],
    const Method::Parameters &parameters)
{
    if (objc > 0) {
        int j = objc - 1;
        Method::Parameters::const_iterator p = parameters.begin();
        for (int i = 0; i < objc && p != parameters.end(); ++i, --j, ++p) {
            if (p->flags() & PARAMFLAG_FOUT) {
                TclObject value(&m_outValues[j], p->type(), interp);
                Tcl_ObjSetVar2(
                    interp, objv[i], NULL, value, TCL_LEAVE_ERR_MSG);
            }
        }
    }
}


int
PositionalArguments::initialize (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[],
    const Method &method,
    WORD dispatchFlags)
{
    const Method::Parameters &parameters = method.parameters();

    int paramCount = parameters.size();
    int inputCount = objc;
    if (dispatchFlags == DISPATCH_PROPERTYPUT
     || dispatchFlags == DISPATCH_PROPERTYPUTREF) {
        paramCount = objc;
        --inputCount;
    }

    if (method.vararg() && inputCount > 0) {
        m_args = new NativeValue[inputCount];

        // Convert the arguments actually provided.
        int inputIndex = 0;
        int argIndex = inputCount - 1;
        for (; inputIndex < inputCount; ++inputIndex, --argIndex) {
            TclObject value(objv[inputIndex]);
            value.toNativeValue(
                &m_args[argIndex], Type::variant(), interp, true);
        }

        paramCount = inputCount;

    } else if (paramCount > 0) {
        m_args = new NativeValue[paramCount];
        m_outValues = new NativeValue[paramCount];

        int j = paramCount - 1;
        Method::Parameters::const_iterator p = parameters.begin();

        // Convert the arguments actually provided.
	int i = 0;
        for (; i < inputCount; ++i, --j, ++p) {
            int result = initArgument(interp, objv[i], j, *p);
            if (result != TCL_OK) {
                return result;
            }
        }

        // Fill in missing arguments.
        for (; p != parameters.end(); ++p, --j) {
            m_args[j] = vtMissing;
        }

        // Convert argument for property put operations.
        if (dispatchFlags == DISPATCH_PROPERTYPUT
         || dispatchFlags == DISPATCH_PROPERTYPUTREF) {
            TclObject value = objv[i];
            value.toNativeValue(&m_args[j], method.type(), interp, true);
        }
    }

    m_dispParams.rgvarg = m_args;
    m_dispParams.cArgs = paramCount;

    if (dispatchFlags == DISPATCH_PROPERTYPUT
     || dispatchFlags == DISPATCH_PROPERTYPUTREF) {
        // Property puts have a named argument that represents the value being
        // assigned to the property.
        static DISPID mydispid = DISPID_PROPERTYPUT;
        m_dispParams.rgdispidNamedArgs = &mydispid;
        m_dispParams.cNamedArgs = 1;
    }

    return TCL_OK;
}


NamedArguments::~NamedArguments ()
{
    delete[] m_namedDispids;
}

Method::Parameters::const_iterator
NamedArguments::findParameter (const Method::Parameters &parameters,
                               const std::string &name,
                               DISPID &dispid)
{
    int i = 0;
    Method::Parameters::const_iterator p = parameters.begin();
    for (; p != parameters.end(); ++p, ++i) {
        if (p->name() == name) {
            dispid = i;
            break;
        }
    }
    return p;
}

int
NamedArguments::initialize (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[],
    const Method &method,
    WORD /*dispatchFlags*/)
{
    const Method::Parameters &parameters = method.parameters();

    if (objc % 2 != 0) {
        Tcl_AppendResult(interp, "name value pairs required", NULL);
        return TCL_ERROR;
    }

    int cArgs = objc / 2;
    if (cArgs > 0) {
        m_args = new NativeValue[cArgs];
        m_outValues = new NativeValue[cArgs];
        m_namedDispids = new DISPID[cArgs];

        int j = cArgs - 1;
        for (int i = 0; i < objc; i += 2, --j) {
            char *name = Tcl_GetStringFromObj(objv[i], 0);
            Method::Parameters::const_iterator p = findParameter(
                parameters,
                name,
                m_namedDispids[j]);
            if (p == parameters.end()) {
                Tcl_AppendResult(interp, "unknown parameter ", name, NULL);
                return TCL_ERROR;
            }

            int result = initArgument(interp, objv[i+1], j, *p);
            if (result != TCL_OK) {
                return result;
            }
        }
    }

    m_dispParams.rgvarg = m_args;
    m_dispParams.rgdispidNamedArgs = m_namedDispids;
    m_dispParams.cArgs = cArgs;
    m_dispParams.cNamedArgs = cArgs;

    return TCL_OK;
}


int
UntypedArguments::initialize (
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[],
    WORD dispatchFlags)
{
    if (objc > 0) {
        m_args = new NativeValue[objc];

        int j = objc - 1;
        for (int i = 0; i < objc; ++i, --j) {
            TclObject value(objv[i]);

            // If the argument is an interface pointer, increment its reference
            // count because the _variant_t destructor will release it.
            value.toNativeValue(&m_args[j], Type::variant(), interp, true);
        }
    }

    m_dispParams.rgvarg = m_args;
    m_dispParams.cArgs = objc;

    if (dispatchFlags == DISPATCH_PROPERTYPUT
     || dispatchFlags == DISPATCH_PROPERTYPUTREF) {
        // Property puts have a named argument that represents the value being
        // assigned to the property.
        static DISPID mydispid = DISPID_PROPERTYPUT;
        m_dispParams.rgdispidNamedArgs = &mydispid;
        m_dispParams.cNamedArgs = 1;
    }

    return TCL_OK;
}
