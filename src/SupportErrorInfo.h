// $Id: SupportErrorInfo.h,v 1.3 2001/07/17 02:24:08 cthuang Exp $
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
