// $Id: TclInterp.h,v 1.8 2002/04/13 03:53:56 cthuang Exp $
#ifndef TCLINTERP_H
#define TCLINTERP_H

#include <string>
#include <tcl.h>

class TclObject;

// This class provides access to a Tcl interpreter loaded from a DLL.

class TclInterp
{
    HINSTANCE m_hmodTcl;
    Tcl_Interp *m_interp;

    // Load and initialize interpreter.
    void doInitialize(const std::string &dllPath);

    // Do not allow others to copy instances of this class.
    TclInterp(const TclInterp &);       // not implemented
    void operator=(const TclInterp &);  // not implemented

public:
    TclInterp();

    // Load Tcl DLL and create interpreter.
    void initialize(const std::string &dllPath);

    // Delete interpreter and unload Tcl DLL.
    void terminate();

    // Evaluate script.
    int eval(const std::string &script);
    int eval(TclObject script, TclObject *pResult=0);

    // Get interpreter result as a string.
    const char *resultString() const
    { return Tcl_GetStringResult(m_interp); }

#if 0
    // Get variable value.
    int getVariable(const char *name, TclObject *pValue) const;

    // Set variable value.
    int setVariable(const char *name, TclObject value);
#endif
};

#endif
