// $Id: NativeValue.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef NATIVEVALUE_H
#define NATIVEVALUE_H

#include <comdef.h>

// This is a value in the native machine format.

class NativeValue: public _variant_t
{
public:
    NativeValue ()
    { }

    ~NativeValue ()
    { fixInvalidVariantType(); }

    NativeValue &operator= (const _variant_t &rhs)
    { _variant_t::operator=(rhs); return *this; }

    // Change the variant type if it is invalid.
    void fixInvalidVariantType();
};

#endif
