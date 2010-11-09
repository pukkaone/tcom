// $Id$
#ifndef ARGUMENTS_H
#define ARGUMENTS_H

#include "TypeInfo.h"
#include "NativeValue.h"

class Arguments
{
protected:
    DISPPARAMS m_dispParams;

    // argument values
    NativeValue *m_args;

    Arguments();

public:
    virtual ~Arguments();

    // Get arguments in the format required by the Invoke function.
    DISPPARAMS *dispParams () const
    { return const_cast<DISPPARAMS *>(&m_dispParams); }
};

// This abstract class represents the arguments passed in and out of a method
// for the case where we know the types of the parameters.

class TypedArguments: public Arguments
{
protected:
    // used to hold values returned from out parameters
    NativeValue *m_outValues;

    TypedArguments();

    // Initialize a single argument.
    int initArgument(
        Tcl_Interp *interp,
        Tcl_Obj *obj,
        int argIndex,
        const Parameter &parameter);

public:
    virtual ~TypedArguments();

    // Ready arguments for method invocation.
    // Returns a Tcl completion code.
    virtual int initialize(
        Tcl_Interp *interp,
        int objc,
        Tcl_Obj *CONST objv[],
        const Method &method,
        WORD dispatchFlags) = 0;

    // Put the values returned from out parameters into Tcl variables.
    void storeOutValues(
        Tcl_Interp *interp,
        int objc,
        Tcl_Obj *CONST objv[],
        const Method::Parameters &parameters);
};

// This class represents arguments specified by their position in an argument
// list.

class PositionalArguments: public TypedArguments
{
public:
    // Ready arguments for method invocation.
    // Returns a Tcl completion code.
    virtual int initialize(
        Tcl_Interp *interp,
        int objc,
        Tcl_Obj *CONST objv[],
        const Method &method,
        WORD dispatchFlags);
};

// This class represents arguments specified by name.

class NamedArguments: public TypedArguments
{
    // Search the parameter list for the named parameter.
    static Method::Parameters::const_iterator findParameter(
        const Method::Parameters &parameters,
        const std::string &name,
        DISPID &dispid);

    // IDs of named arguments
    DISPID *m_namedDispids;

public:
    NamedArguments ():
        m_namedDispids(0)
    { }

    ~NamedArguments();

    // Ready arguments for method invocation.
    // Returns a Tcl completion code.
    virtual int initialize(
        Tcl_Interp *interp,
        int objc,
        Tcl_Obj *CONST objv[],
        const Method &method,
        WORD dispatchFlags);
};

// This class represents the arguments passed into a method
// for the case where we don't know the types of the parameters.

class UntypedArguments: public Arguments
{
public:
    // Ready arguments for method invocation.
    // Returns a Tcl result code.
    int initialize(
        Tcl_Interp *interp,
        int objc,
        Tcl_Obj *CONST objv[],
        WORD dispatchFlags);
};

#endif 
