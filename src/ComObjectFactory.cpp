// $Id: ComObjectFactory.cpp,v 1.17 2002/05/31 04:03:06 cthuang Exp $
#pragma warning(disable: 4786)
#include "ComModule.h"
#include "ComObject.h"
#include "ComObjectFactory.h"

ComObjectFactory::ComObjectFactory (const Class::Interfaces &interfaces,
                                    Tcl_Interp *interp,
                                    TclObject constructor,
                                    TclObject destructor,
                                    bool registerActiveObject):
        m_refCount(0),
        m_interfaces(interfaces),
        m_interp(interp),
        m_constructor(constructor),
        m_destructor(destructor),
        m_registerActiveObject(registerActiveObject),
        m_registeredFactory(false)
{ }

ComObjectFactory::~ComObjectFactory ()
{
    if (m_registeredFactory) {
        // TODO: This call can return an error but I don't want to throw an
        // exception from a destructor.
        CoRevokeClassObject(m_classObjectHandle);
    }
}

void
ComObjectFactory::registerFactory (REFCLSID clsid, DWORD regclsFlags)
{
    m_clsid = clsid;

    HRESULT hr = CoRegisterClassObject(
        clsid,
        this,
        CLSCTX_SERVER,
        regclsFlags,
        &m_classObjectHandle);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }
    m_registeredFactory = true;
}

STDMETHODIMP
ComObjectFactory::QueryInterface (REFIID iid, void **ppvObj)
{
    if (IsEqualIID(iid, IID_IClassFactory) || IsEqualIID(iid, IID_IUnknown)) {
        *ppvObj = this;
        AddRef();
        return S_OK;
    }

    *ppvObj = 0;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG)
ComObjectFactory::AddRef ()
{
    InterlockedIncrement(&m_refCount);
    return m_refCount;
}

STDMETHODIMP_(ULONG)
ComObjectFactory::Release ()
{
    InterlockedDecrement(&m_refCount);
    if (m_refCount == 0) {
        delete this;
        return 0;
    }
    return m_refCount;
}

int
ComObjectFactory::eval (TclObject script, TclObject *pResult)
{
    int completionCode =
#if TCL_MINOR_VERSION >= 1
        Tcl_EvalObjEx(m_interp, script, TCL_EVAL_DIRECT | TCL_EVAL_GLOBAL);
#else
        Tcl_GlobalEvalObj(m_interp, script);
#endif

    if (pResult != 0) {
        *pResult = Tcl_GetObjResult(m_interp);
    }
    return completionCode;
}

STDMETHODIMP
ComObjectFactory::CreateInstance (IUnknown *pOuter, REFIID iid, void **ppvObj)
{
    // We don't support aggregation.
    if (pOuter != 0) {
        *ppvObj = 0;
        return CLASS_E_NOAGGREGATION;
    }

    // Execute Tcl script to create a servant.  The script should return the
    // name of a Tcl command which implements the object's operations.
    TclObject servant;
    int completionCode = eval(m_constructor, &servant);
    if (completionCode != TCL_OK) {
        *ppvObj = 0;
        return E_UNEXPECTED;
    }

    // Create a COM object and tie its implementation to the servant.
    ComObject *pComObject = ComObject::newInstance(
        m_interfaces,
        m_interp,
        servant,
        m_destructor);

    if (m_registerActiveObject) {
        pComObject->registerActiveObject(m_clsid);
    }

    return pComObject->unknown()->QueryInterface(iid, ppvObj);
}

STDMETHODIMP
ComObjectFactory::LockServer (BOOL lock)
{
    if (lock) {
        ComModule::instance().lock();
    } else {
        ComModule::instance().unlock();
    }
    return S_OK;
}


SingletonObjectFactory::SingletonObjectFactory (
    const Class::Interfaces &interfaces,
    Tcl_Interp *interp,
    TclObject constructor,
    TclObject destructor,
    bool registerActiveObject):
        ComObjectFactory(
            interfaces,
            interp,
            constructor,
            destructor,
            registerActiveObject),
        m_pInstance(0)
{ }

SingletonObjectFactory::~SingletonObjectFactory ()
{
    if (m_pInstance != 0) {
        m_pInstance->Release();
    }
}

STDMETHODIMP
SingletonObjectFactory::CreateInstance (IUnknown *pOuter,
                                        REFIID iid,
                                        void **ppvObj)
{
    if (m_pInstance == 0) {
	LOCK_MUTEX(m_mutex)
        if (m_pInstance == 0) {
	    HRESULT hr = ComObjectFactory::CreateInstance(
                pOuter,
                iid,
                reinterpret_cast<void **>(&m_pInstance));
            if (FAILED(hr)) {
                return hr;
            }
        }
    }

    return m_pInstance->QueryInterface(iid, ppvObj);
}
