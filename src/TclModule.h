// $Id$
#ifndef TCLMODULE_H
#define TCLMODULE_H

#include "ComModule.h"
#include "TclInterp.h"

// This is a COM module used to implement COM objects in Tcl.

class TclModule: public ComModule
{
    TclInterp m_interp;

protected:
    TclModule ()
    { }

public:
    // Register a class factory by executing a Tcl script associated with
    // its CLSID.  It's expected the Tcl script will register a class factory
    // using the "::tcom::object registerfactory" command.
    // Returns a Tcl completion code.
    int registerFactoryByScript(const std::string &clsid);

    // Shut down server.
    void terminate();
};

#endif
