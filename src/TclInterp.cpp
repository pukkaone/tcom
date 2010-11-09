// $Id$
#include <sstream>
#include "RegistryKey.h"
#include "TclObject.h"
#include "TclInterp.h"

TclInterp::TclInterp ():
    m_hmodTcl(NULL),
    m_interp(0)
{ }

void
TclInterp::terminate ()
{
    if (m_interp != 0) {
        Tcl_DeleteInterp(m_interp);
    }

    if (m_hmodTcl != NULL) {
        FreeLibrary(m_hmodTcl);
    }
}

void
TclInterp::initialize (const std::string &dllPath)
{
    if (m_interp == 0) {
        doInitialize(dllPath);
    }
}

// Load the Tcl DLL.  First try to load from the given path.  If that
// fails, look for the Tcl DLL path in the registry and try loading it.

static HINSTANCE
loadTclLibrary (const std::string &firstDllPath, std::string &foundDllPath)
{
    if (!firstDllPath.empty()) {
        _bstr_t path(firstDllPath.c_str());
        HINSTANCE hmod = LoadLibrary(path);
        if (hmod != NULL) {
            foundDllPath = firstDllPath;
            return hmod;
        }
    }

    std::string activeTclKeyName("SOFTWARE\\ActiveState\\ActiveTcl");
    RegistryKey activeTclKey(HKEY_LOCAL_MACHINE, activeTclKeyName);
    std::string currentVersion = activeTclKey.value("CurrentVersion");

    std::istringstream iss(currentVersion);
    int major, minor;
    char dot;
    iss >> major >> dot >> minor;

    std::string versionKeyName(activeTclKeyName);
    versionKeyName += "\\";
    versionKeyName += currentVersion;
    RegistryKey versionKey(HKEY_LOCAL_MACHINE, versionKeyName);

    std::ostringstream oss;
    oss << versionKey.value() << "\\bin\\tcl" << major << minor << ".dll";
    std::string dllPath(oss.str());

    _bstr_t path(dllPath.c_str());
    HINSTANCE hmod = LoadLibrary(path);
    if (hmod != NULL) {
        foundDllPath = dllPath;
    }
    return hmod;
}

void
TclInterp::doInitialize (const std::string &firstDllPath)
{
    // Load Tcl library.
    std::string dllPath;
    m_hmodTcl = loadTclLibrary(firstDllPath, dllPath);
    if (m_hmodTcl == NULL) {
        throw std::runtime_error("LoadLibrary");
    }

    // Get address of Tcl_CreateInterp function.
    typedef Tcl_Interp *(*CreateInterpFunc)();
    CreateInterpFunc createInterp = reinterpret_cast<CreateInterpFunc>(
        GetProcAddress(m_hmodTcl, "Tcl_CreateInterp"));
    if (createInterp == NULL) {
        throw std::runtime_error("GetProcAddress Tcl_CreateInterp");
    }

    // Create Tcl interpreter.
    m_interp = createInterp();
    if (Tcl_InitStubs(m_interp, "8.1", 0) == NULL) {
        throw std::runtime_error("Tcl_InitStubs");
    }

    Tcl_FindExecutable(dllPath.c_str());

    // Find and source Tcl initialization script.
    if (Tcl_Init(m_interp) != TCL_OK) {
        throw std::runtime_error(Tcl_GetStringResult(m_interp));
    }
}

int
TclInterp::eval (const std::string &script)
{
    return Tcl_EvalEx(
        m_interp,
        const_cast<char *>(script.data()),
        script.size(),
        TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);
}

int
TclInterp::eval (TclObject script, TclObject *pResult)
{
    int completionCode = Tcl_EvalObjEx(
        m_interp, script, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);
    if (pResult != 0) {
        *pResult = Tcl_GetObjResult(m_interp);
    }
    return completionCode;
}
