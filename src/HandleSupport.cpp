// $Id$
#include "HandleSupport.h"
#include <sstream>
#include "ThreadLocalStorage.h"

InternalRep::InternalRep (
    Tcl_Interp *interp,
    Tcl_ObjCmdProc *pCmdProc,
    ClientData objClientData):
        m_interp(interp),
        m_clientData(objClientData),
        m_handleCount(0)
{
    std::string handleName(name());

    m_command = Tcl_CreateObjCommand(
        m_interp,
        const_cast<char *>(handleName.c_str()),
        pCmdProc,
        objClientData,
        0);

    m_pNameEntry = HandleNameToRepMap::instance(interp)->insert(
        handleName.c_str(), this);
}

InternalRep::~InternalRep ()
{
    HandleNameToRepMap::erase(m_pNameEntry);
    if (!Tcl_InterpDeleted(m_interp)) {
        Tcl_DeleteCommandFromToken(m_interp, m_command);
    }
}

std::string
InternalRep::name () const
{
    std::ostringstream oss;
    oss << "::tcom::handle0x" << std::hex << this;
    return oss.str();
}

void
InternalRep::incrHandleCount ()
{
    ++m_handleCount;
}

long
InternalRep::decrHandleCount ()
{
    if (--m_handleCount == 0) {
        delete this;
        return 0;
    }
    return m_handleCount;
}


// This maps Tcl_Obj pointers to an internal representation.

class ObjToRepMap
{
    Tcl_HashTable m_objMap;

    static ThreadLocalStorage<ObjToRepMap> ms_tls;

    static void exitProc(ClientData);

    // not implemented
    ObjToRepMap(const ObjToRepMap &);
    void operator=(const ObjToRepMap &);

    ~ObjToRepMap();

public:
    ObjToRepMap();
    static ObjToRepMap &instance();

    void insert(Tcl_Obj *pObj, InternalRep *pRep);
    InternalRep *find(Tcl_Obj *pObj);
    void erase(Tcl_Obj *pObj);
};

ThreadLocalStorage<ObjToRepMap> ObjToRepMap::ms_tls;

void
ObjToRepMap::exitProc (ClientData clientData)
{
    delete static_cast<ObjToRepMap *>(clientData);
}

ObjToRepMap::ObjToRepMap ()
{
    Tcl_InitHashTable(&m_objMap, TCL_ONE_WORD_KEYS);

#ifdef TCL_THREADS
    Tcl_CreateThreadExitHandler(exitProc, this);
#else
    Tcl_CreateExitHandler(exitProc, 0);
#endif
}

ObjToRepMap::~ObjToRepMap ()
{
    Tcl_DeleteHashTable(&m_objMap);
}

ObjToRepMap &
ObjToRepMap::instance ()
{
    return ms_tls.instance();
}

void
ObjToRepMap::insert (Tcl_Obj *pObj, InternalRep *pRep)
{
    int isNew;
    Tcl_HashEntry *pEntry = Tcl_CreateHashEntry(
        &m_objMap, reinterpret_cast<char *>(pObj), &isNew);
    Tcl_SetHashValue(pEntry, pRep);
}

InternalRep *
ObjToRepMap::find (Tcl_Obj *pObj)
{
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        &m_objMap, reinterpret_cast<char *>(pObj));
    if (pEntry == 0) {
        return 0;
    }
    return static_cast<InternalRep *>(Tcl_GetHashValue(pEntry));
}

void
ObjToRepMap::erase (Tcl_Obj *pObj)
{
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        &m_objMap, reinterpret_cast<char *>(pObj));
    if (pEntry != 0) {
        Tcl_DeleteHashEntry(pEntry);
    }
}


Tcl_ObjType *CmdNameType::ms_pCmdNameType;
Tcl_ObjType CmdNameType::ms_oldCmdNameType;

Singleton<CmdNameType> CmdNameType::ms_singleton;

CmdNameType &
CmdNameType::instance ()
{
    return ms_singleton.instance();
}

CmdNameType::CmdNameType ()
{
    // Hijack Tcl's cmdName type.
    ms_pCmdNameType = Tcl_GetObjType("cmdName");
    ms_oldCmdNameType = *ms_pCmdNameType;
    ms_pCmdNameType->freeIntRepProc = freeInternalRep;
    ms_pCmdNameType->dupIntRepProc = dupInternalRep;
    ms_pCmdNameType->updateStringProc = updateString;
    ms_pCmdNameType->setFromAnyProc = setFromAny;
}

CmdNameType::~CmdNameType ()
{
    // Restore original cmdName type.
    ms_pCmdNameType->freeIntRepProc = ms_oldCmdNameType.freeIntRepProc;
    ms_pCmdNameType->dupIntRepProc = ms_oldCmdNameType.dupIntRepProc;
    ms_pCmdNameType->updateStringProc = ms_oldCmdNameType.updateStringProc;
    ms_pCmdNameType->setFromAnyProc = ms_oldCmdNameType.setFromAnyProc;
}

void
CmdNameType::freeInternalRep (Tcl_Obj *pObj)
{
    if (pObj->refCount == 0) {
        InternalRep *pRep = ObjToRepMap::instance().find(pObj);
        if (pRep != 0) {
            ObjToRepMap::instance().erase(pObj);
            pRep->decrHandleCount();
        }
    }

    ms_oldCmdNameType.freeIntRepProc(pObj);
}

void
CmdNameType::dupInternalRep (Tcl_Obj *pSrc, Tcl_Obj *pDup)
{
    // TODO: An object is duplicated if it is about to be modified.  We cannot
    // allow modifications on a handle.

    ms_oldCmdNameType.dupIntRepProc(pSrc, pDup);
}

// This should never be called because the string representation should already
// be valid.

void
CmdNameType::updateString (Tcl_Obj *pObj)
{
    ms_oldCmdNameType.updateStringProc(pObj);
}

int
CmdNameType::setFromAny (Tcl_Interp *interp, Tcl_Obj *pObj)
{
    // Check if the string represents an existing handle.
    HandleNameToRepMap *pHandleNameToRepMap =
        HandleNameToRepMap::instance(interp);
    if (pHandleNameToRepMap != 0) {
        InternalRep *pRep = pHandleNameToRepMap->find(pObj);
        if (pRep != 0) {
            if (ObjToRepMap::instance().find(pObj) == 0) {
                ObjToRepMap::instance().insert(pObj, pRep);
                pRep->incrHandleCount();
            }
        }
    }

    return ms_oldCmdNameType.setFromAnyProc(interp, pObj);
}

Tcl_Obj *
CmdNameType::newObj (Tcl_Interp *interp, InternalRep *pRep)
{
    Tcl_Obj *pObj = Tcl_NewStringObj(
        const_cast<char *>(pRep->name().c_str()), -1);
    Tcl_ConvertToType(interp, pObj, ms_pCmdNameType);
    return pObj;
}


static char ASSOC_KEY[] = "tcomHandles";

HandleNameToRepMap::HandleNameToRepMap (Tcl_Interp *interp):
    m_interp(interp)
{
    Tcl_InitHashTable(&m_handleMap, TCL_STRING_KEYS);

    Tcl_SetAssocData(interp, ASSOC_KEY, deleteInterpProc, this);
    Tcl_CreateExitHandler(exitProc, this);
}

HandleNameToRepMap::~HandleNameToRepMap ()
{
    // Clean up any left over objects.
    Tcl_HashSearch search;
    Tcl_HashEntry *pEntry = Tcl_FirstHashEntry(&m_handleMap, &search);
    while (pEntry != 0) {
        Tcl_HashEntry *pNext = Tcl_NextHashEntry(&search);
        delete static_cast<InternalRep *>(Tcl_GetHashValue(pEntry));
        pEntry = pNext;
    }

    Tcl_DeleteHashTable(&m_handleMap);
}

void
HandleNameToRepMap::deleteInterpProc (ClientData clientData, Tcl_Interp *)
{
    Tcl_DeleteExitHandler(exitProc, clientData);
    delete static_cast<HandleNameToRepMap *>(clientData);
}

void
HandleNameToRepMap::exitProc (ClientData clientData)
{
    HandleNameToRepMap *pHandleNameToRepMap =
        static_cast<HandleNameToRepMap *>(clientData);
    Tcl_DeleteAssocData(pHandleNameToRepMap->m_interp, ASSOC_KEY);
}

HandleNameToRepMap *
HandleNameToRepMap::instance (Tcl_Interp *interp)
{
    return static_cast<HandleNameToRepMap *>(
        Tcl_GetAssocData(interp, ASSOC_KEY, 0));
}

Tcl_HashEntry *
HandleNameToRepMap::insert (const char *handleStr, InternalRep *pRep)
{
    int isNew;
    Tcl_HashEntry *pEntry = Tcl_CreateHashEntry(
        &m_handleMap, const_cast<char *>(handleStr), &isNew);
    Tcl_SetHashValue(pEntry, static_cast<ClientData>(pRep));
    return pEntry;
}

InternalRep *
HandleNameToRepMap::find (Tcl_Obj *pHandle) const
{
    char *key = Tcl_GetStringFromObj(pHandle, 0);
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        const_cast<Tcl_HashTable *>(&m_handleMap), key);
    if (pEntry == 0) {
        return 0;
    }
    return static_cast<InternalRep *>(Tcl_GetHashValue(pEntry));
}

void
HandleNameToRepMap::erase (Tcl_HashEntry *pNameEntry)
{
    Tcl_DeleteHashEntry(pNameEntry);
}
