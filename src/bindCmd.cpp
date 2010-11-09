// $Id: bindCmd.cpp 13 2005-04-18 12:24:14Z cthuang $
#pragma warning(disable: 4786)
#include "Extension.h"
#include "Reference.h"
#include "TypeLib.h"

// Get the interface description for the specified IID.
// On failure, put a message in the Tcl interpreter result and return 0.

static const Interface *
getInterfaceDesc (Tcl_Interp *interp, REFIID iid)
{
    const Interface *pInterface = InterfaceManager::instance().find(iid);
    if (pInterface == 0) {
        Tcl_AppendResult(
            interp, "no event interface information", NULL);
    }
    return pInterface;
}

// Get the default source interface from the class description provided by
// IProvideClassInfo.
// On failure, return 0.

static const Interface *
findEventInterfaceFromProvideClassInfo (IUnknown *pObject)
{
    HRESULT hr;

    IProvideClassInfoPtr pProvideClassInfo;
    hr = pObject->QueryInterface(
        IID_IProvideClassInfo,
        reinterpret_cast<void **>(&pProvideClassInfo));
    if (FAILED(hr)) {
        return 0;
    }

    ITypeInfoPtr pClassTypeInfo;
    hr = pProvideClassInfo->GetClassInfo(&pClassTypeInfo);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Get default source interface from class description.
    Class aClass(pClassTypeInfo);
    return aClass.sourceInterface();
}

// Get the default source interface from the class description loaded from a
// type library specified by a CLSID.
// On failure, return 0.

static const Interface *
findEventInterfaceFromClsid (Reference *pReference)
{
    const Class *pClass = pReference->classDesc();
    if (pClass == 0) {
        return 0;
    }
    return pClass->sourceInterface();
}

// Get the event interface managed by the first connection point from the
// connection point container.
// On failure, put a message in the Tcl interpreter result and return 0.

static const Interface *
findEventInterfaceFromConnectionPoint (Tcl_Interp *interp, IUnknown *pObject)
{
    HRESULT hr;

    // Get connection point container.
    IConnectionPointContainerPtr pContainer;
    hr = pObject->QueryInterface(
        IID_IConnectionPointContainer,
        reinterpret_cast<void **>(&pContainer));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Get connection point enumerator.
    IEnumConnectionPointsPtr pEnum;
    hr = pContainer->EnumConnectionPoints(&pEnum);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Get first connection point.
    IConnectionPointPtr pConnectionPoint;
    ULONG count;
    hr = pEnum->Next(1, &pConnectionPoint, &count);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
    if (count == 0) {
        Tcl_AppendResult(
            interp, "IEnumConnectionPoints returned no elements", NULL);
        return 0;
    }

    // Get IID of event interface managed by the connection point.
    IID iid;
    hr = pConnectionPoint->GetConnectionInterface(&iid);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    // Get interface description for the IID.
    const Interface *pInterface = InterfaceManager::instance().find(iid);
    if (pInterface != 0) {
        return pInterface;
    }

    // If we don't have the interface description in the cache, try loading
    // it from the type library.
    TypeLib *pTypeLib = TypeLib::loadByIid(iid);
    delete pTypeLib;
    return getInterfaceDesc(interp, iid);
}

// Find the event interface.
// On failure, put a message in the Tcl interpreter result and return 0.

static const Interface *
findEventInterface (
    Tcl_Interp *interp,
    Reference *pReference,
    char *eventIIDStr)
{
    if (eventIIDStr != 0) {
        // The script provided the IID of the event interface.
        IID eventIID;
        if (UuidFromString(reinterpret_cast<unsigned char *>(eventIIDStr),
         &eventIID) != RPC_S_OK) {
            Tcl_AppendResult(
                interp,
                "cannot convert to IID: ",
                eventIIDStr,
                NULL);
            return 0;
        }

        return getInterfaceDesc(interp, eventIID);
    }

    const Interface *pInterface;

    // If the object implements IProvideClassInfo, get the default source
    // interface from the class description.
    pInterface = findEventInterfaceFromProvideClassInfo(pReference->unknown());
    if (pInterface != 0) {
        return pInterface;
    }

    // If we know the CLSID of the object's class, load the type library
    // containing the class description, and get the default source interface
    // from the class description.
    pInterface = findEventInterfaceFromClsid(pReference);
    if (pInterface != 0) {
        return pInterface;
    }

    // Get the event interface of the first connection point in the connection
    // pointer container.
    return findEventInterfaceFromConnectionPoint(interp, pReference->unknown());
}

// This Tcl command binds a Tcl command to an event sink.

int
Extension::bindCmd (
    ClientData,
    Tcl_Interp *interp,
    int objc,
    Tcl_Obj *CONST objv[])
{
    if (objc < 3 || objc > 4) {
	Tcl_WrongNumArgs(interp, 1, objv, "handle sinkCommand ?eventIID?");
	return TCL_ERROR;
    }

    Reference *pReference = referenceHandles.find(interp, objv[1]);
    if (pReference == 0) {
        const char *arg = Tcl_GetStringFromObj(objv[1], 0);
        Tcl_AppendResult(
            interp, "invalid interface pointer handle ", arg, NULL);
        return TCL_ERROR;
    }

    int servantLength;
    Tcl_GetStringFromObj(objv[2], &servantLength);
    if (servantLength == 0) {
        try {
            // Tear down all event connections to the object.
            pReference->unadvise();
        }
        catch (_com_error &e) {
            return setComErrorResult(interp, e, __FILE__, __LINE__);
        }
        return TCL_OK;
    }

    TclObject servant(objv[2]);

    char *eventIIDStr = (objc < 4) ? 0 : Tcl_GetStringFromObj(objv[3], 0);

    try {
        const Interface *pEventInterface = findEventInterface(
            interp, pReference, eventIIDStr);
        if (pEventInterface == 0) {
            return TCL_ERROR;
        }

        pReference->advise(interp, *pEventInterface, servant);
    }
    catch (_com_error &e) {
        return setComErrorResult(interp, e, __FILE__, __LINE__);
    }
    return TCL_OK;
}
