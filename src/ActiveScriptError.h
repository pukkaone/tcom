// $Id: ActiveScriptError.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef ACTIVESCRIPTERROR_H
#define ACTIVESCRIPTERROR_H

#include <activscp.h>
#include <comdef.h>

// This class implements IActiveScriptError.

class ActiveScriptError: public IActiveScriptError
{
    long m_refCount;
    HRESULT m_hresult;
    _bstr_t m_source;
    _bstr_t m_description;
    ULONG m_lineNumber;
    long m_characterPosition;
    _bstr_t m_sourceLineText;
    
public:
    ActiveScriptError (
            HRESULT hresult,
            const char *source,
            const char *description,
            ULONG lineNumber,
            long characterPosition,
            const char *sourceLineText):
        m_refCount(0),
        m_hresult(hresult),
        m_source(source),
        m_description(description),
        m_lineNumber(lineNumber),
        m_characterPosition(characterPosition),
        m_sourceLineText(sourceLineText)
    { }

    // IUnknown implementation
    STDMETHODIMP QueryInterface(REFIID iid, void **ppvObj);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IActiveScriptError implementation
    STDMETHODIMP GetExceptionInfo(EXCEPINFO *pExcepInfo);
    STDMETHODIMP GetSourcePosition( 
        DWORD *pSourceContext, ULONG *pLineNumber, LONG *pCharacterPosition);
    STDMETHODIMP GetSourceLineText(BSTR *pSourceLineText);
};

#endif
