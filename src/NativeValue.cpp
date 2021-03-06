// $Id$
#include "NativeValue.h"

void
NativeValue::fixInvalidVariantType ()
{
    if (vt == VT_I8 || vt == VT_UI8) {
        // 64-bit integers are not valid VARIANT types.  Change the VARIANT
        // type to something valid so VariantClear does not return an error.
        vt = VT_EMPTY;
    }
}
