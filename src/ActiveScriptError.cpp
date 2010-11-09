// $Id: ActiveScriptError.cpp,v 1.1 2002/03/30 18:49:53 cthuang Exp $
#include "ActiveScriptError.h"

STDMETHODIMP
ActiveScriptError::QueryInterface (REFIID iid, void **ppvObj)
{
    if (IsEqualIID(iid, IID_IUnknown)
     || IsEqualIID(iid, IID_IActiveScriptError)) {
	*ppvObj = this;
        AddRef();
	return S_OK;
    }

    *ppvObj = 0;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG)
ActiveScriptError::AddRef ()
{
    InterlockedIncrement(&m_refCount);
    return m_refCount;
}

STDMETHODIMP_(ULONG)
ActiveScriptError::Release ()
{
    InterlockedDecrement(&m_refCount);
    if (m_refCount == 0) {
	delete this;
        return 0;
    }
    return m_refCount;
}

STDMETHODIMP
ActiveScriptError::GetExceptionInfo (EXCEPINFO *pExcepInfo)
{
    if (pExcepInfo == 0) {
        return E_POINTER;
    }

    memset(pExcepInfo, 0, sizeof(EXCEPINFO));

    pExcepInfo->scode = m_hresult;
    pExcepInfo->bstrSource = SysAllocString(m_source);
    pExcepInfo->bstrDescription = SysAllocString(m_description);
    return S_OK;
}

STDMETHODIMP
ActiveScriptError::GetSourcePosition ( 
    DWORD *pSourceContext, ULONG *pLineNumber, LONG *pCharacterPosition)
{
    *pSourceContext = 0;
    *pLineNumber = m_lineNumber;
    *pCharacterPosition = m_characterPosition;
    return S_OK;
}

STDMETHODIMP
ActiveScriptError::GetSourceLineText (BSTR *pSourceLineText)
{
    *pSourceLineText = SysAllocString(m_sourceLineText);
    return S_OK;
}
