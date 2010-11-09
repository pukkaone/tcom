// $Id: SupportErrorInfo.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef SUPPORTERRORINFO_H
#define SUPPORTERRORINFO_H

#include <comdef.h>
#include "tcomApi.h"

class TCOM_API ComObject;

// This class implements ISupportErrorInfo.

class SupportErrorInfo: public ISupportErrorInfo
{
    ComObject &m_object;

public:
    SupportErrorInfo (ComObject &object):
        m_object(object)
    { }

    // IUnknown implementation
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // ISupportErrorInfo implementation
    STDMETHODIMP InterfaceSupportsErrorInfo(REFIID riid);
};

#endif
