// $Id: HandleSupport.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef HANDLESUPPORT_H
#define HANDLESUPPORT_H

#include <tcl.h>
#include <string>
#include "tcomApi.h"
#include "Singleton.h"

// This class represents an association from a handle to an application object.
// A handle maps to an object of this class.

class TCOM_API InternalRep
{
protected:
    Tcl_Interp *m_interp;
    Tcl_Command m_command;
    ClientData m_clientData;
    Tcl_HashEntry *m_pNameEntry;

    // number of Tcl_Obj instances that are handles to this object
    long m_handleCount;

public:
    InternalRep(
	Tcl_Interp *interp,
	Tcl_ObjCmdProc *pCmdProc,
	ClientData clientData);
    virtual ~InternalRep();

    // Get handle name.
    std::string name() const;

    // Get pointer to the application object.
    ClientData clientData () const
    { return m_clientData; }

    void incrHandleCount();
    long decrHandleCount();
};

// This class extends InternalRep to associate with a specific application
// object class.  The class takes ownership of the passed in application object
// and is responsible for deleting it.

template<class AppType>
class AppInternalRep: public InternalRep
{
public:
    AppInternalRep (
	Tcl_Interp *interp,
	Tcl_ObjCmdProc *pCmdProc,
	AppType *pAppObject):
	    InternalRep(interp, pCmdProc, pAppObject)
    { }

    virtual ~AppInternalRep();
};

template<class AppType>
AppInternalRep<AppType>::~AppInternalRep ()
{
    delete reinterpret_cast<AppType *>(clientData());
}

// Handles are instances of Tcl's cmdName type which this class hijacks in
// order to map handles to application objects.

class TCOM_API CmdNameType
{
    // pointer to Tcl cmdName type
    static Tcl_ObjType *ms_pCmdNameType;

    // saved Tcl cmdName type
    static Tcl_ObjType ms_oldCmdNameType;

    // Tcl type functions
    static void freeInternalRep(Tcl_Obj *pObj);
    static void dupInternalRep(Tcl_Obj *pSrc, Tcl_Obj *pDup);
    static void updateString(Tcl_Obj *pObj);
    static int setFromAny(Tcl_Interp *interp, Tcl_Obj *pObj);

    friend class Singleton<CmdNameType>;
    static Singleton<CmdNameType> ms_singleton;

    CmdNameType();
    ~CmdNameType();

public:
    // Get instance of this class.
    static CmdNameType &instance();

    // Create handle.
    Tcl_Obj *newObj(Tcl_Interp *interp, InternalRep *pRep);
};

// Maps string representation of handle to internal representation.  There's an
// instance of this class associated with each Tcl interpreter that loads the
// extension.

class TCOM_API HandleNameToRepMap
{
    Tcl_Interp *m_interp;

    // handle string representation to internal representation map
    Tcl_HashTable m_handleMap;

    static void deleteInterpProc(ClientData clientData, Tcl_Interp *interp);
    static void exitProc(ClientData clientData);

    ~HandleNameToRepMap();

public:
    HandleNameToRepMap(Tcl_Interp *interp);

    // Get instance associated with the Tcl interpreter.
    static HandleNameToRepMap *instance(Tcl_Interp *interp);
    
    // Insert handle to object mapping.
    Tcl_HashEntry *insert(const char *handleStr, InternalRep *pRep);

    // Get the object represented by the handle.
    InternalRep *find(Tcl_Obj *pHandle) const;

    // Remove handle to object mapping.
    static void erase(Tcl_HashEntry *pNameEntry);
};

// This class provides functions to map handles to objects of a specific
// application class.

template<class AppType>
class HandleSupport
{
    // Tcl command that implements the operations of the object
    Tcl_ObjCmdProc *m_pCmdProc;

public:
    HandleSupport (Tcl_ObjCmdProc *pCmdProc):
        m_pCmdProc(pCmdProc)
    { }

    // Create a handle and associate it with an application object.  Takes
    // ownership of the application object and is responsible for deleting it.
    Tcl_Obj *newObj(Tcl_Interp *interp, AppType *pAppObject);

    // Get the application object represented by the handle.  If the handle
    // is invalid, return 0.
    AppType *find(Tcl_Interp *interp, Tcl_Obj *pHandle) const;
};

template<class AppType>
Tcl_Obj *
HandleSupport<AppType>::newObj (Tcl_Interp *interp, AppType *pAppObject)
{
    AppInternalRep<AppType> *pRep = new AppInternalRep<AppType>(
	interp, m_pCmdProc, pAppObject);
    return CmdNameType::instance().newObj(interp, pRep);
}

template<class AppType>
AppType *
HandleSupport<AppType>::find (Tcl_Interp *interp, Tcl_Obj *pObj) const
{
    InternalRep *pRep = HandleNameToRepMap::instance(interp)->find(pObj);
    if (pRep == 0) {
	return 0;
    }
    return reinterpret_cast<AppType *>(pRep->clientData());
}

#endif
