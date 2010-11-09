// $Id: SupportErrorInfo.cpp,v 1.3 2001/07/17 02:24:08 cthuang Exp $
#include "ComObject.h"
#include "SupportErrorInfo.h"

STDMETHODIMP
SupportErrorInfo::QueryInterface (REFIID iid, void **ppv)
{
   return m_object.queryInterface(iid, ppv);
}

STDMETHODIMP_(ULONG)
SupportErrorInfo::AddRef ()
{
   return m_object.addRef();
}

STDMETHODIMP_(ULONG)
SupportErrorInfo::Release ()
{
   return m_object.release();
}

STDMETHODIMP
SupportErrorInfo::InterfaceSupportsErrorInfo (REFIID iid)
{
   return m_object.implemented(iid);
}
