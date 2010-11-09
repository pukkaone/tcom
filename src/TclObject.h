// $Id: TclObject.h,v 1.12 2002/04/12 02:55:28 cthuang Exp $
#ifndef TCLOBJECT_H
#define TCLOBJECT_H

#ifdef WIN32
#include "TypeInfo.h"
#endif
#include <tcl.h>
#include <string>
#include "tcomApi.h"

// This class provides access to Tcl's built-in internal representation types.

class TclTypes
{
    static Tcl_ObjType *ms_pBooleanType;
    static Tcl_ObjType *ms_pDoubleType;
    static Tcl_ObjType *ms_pIntType;
    static Tcl_ObjType *ms_pListType;

public:
    static void initialize();

    static Tcl_ObjType *booleanType ()
    { return ms_pBooleanType; }

    static Tcl_ObjType *doubleType ()
    { return ms_pDoubleType; }

    static Tcl_ObjType *intType ()
    { return ms_pIntType; }

    static Tcl_ObjType *listType ()
    { return ms_pListType; }

#if TCL_MINOR_VERSION >= 1
private:
    static Tcl_ObjType *ms_pByteArrayType;

public:
    static Tcl_ObjType *byteArrayType ()
    { return ms_pByteArrayType; }
#endif
};

// This class wraps a Tcl_Obj pointer to provide copy and value semantics by
// automatically incrementing and decrementing its reference count.

class TCOM_API TclObject
{
    Tcl_Obj *m_pObj;

public:
    TclObject();
    TclObject(const TclObject &rhs);
    TclObject(Tcl_Obj *pObj);
    TclObject(const char *src, int len = -1);
    TclObject(const std::string &s);
    TclObject(bool value);
    TclObject(int value);
    TclObject(long value);
    TclObject(double value);
    TclObject(int objc, Tcl_Obj *CONST objv[]);
    ~TclObject();

    TclObject &operator=(const TclObject &rhs);
    TclObject &operator=(Tcl_Obj *pObj);

    // Get raw object pointer.
    operator Tcl_Obj * () const
    { return const_cast<Tcl_Obj *>(m_pObj); }

    // Get UTF-8 string representation of the object.
    const char *c_str () const
    { return Tcl_GetStringFromObj(const_cast<Tcl_Obj *>(m_pObj), 0); }

#if TCL_MINOR_VERSION >= 2
    // Construct Unicode string value.
    TclObject(const wchar_t *src, int len = -1);

    // Get Unicode string representation of the object.
    const wchar_t *getUnicode() const
    { return reinterpret_cast<const wchar_t *>(
	Tcl_GetUnicode(const_cast<Tcl_Obj *>(m_pObj))); }
#endif

    // Convert object to bool and return value.
    bool getBool() const;

    // Convert object to int and return value.
    int getInt() const;

    // Convert object to long and return value.
    long getLong() const;

    // Convert object to double and return value.
    double getDouble() const;

    // Convert the object to a list if it's not already a list,
    // and then append the element to the end of the list.
    TclObject &lappend(Tcl_Obj *pElement);

#ifdef WIN32
    // Construct Tcl object from VARIANT value.
    TclObject(
        VARIANT *pSrc,          // VARIANT value to convert from
        const Type &type,       // expected type for interface pointers
        Tcl_Interp *interp);

    // Convert Tcl object to VARIANT value.
    void toVariant(
        VARIANT *pDest,         // converted value put here
        const Type &type,       // desired data type
        Tcl_Interp *interp,
        bool addRef=false);     // call AddRef on interface pointer

    // Get BSTR representation.  Caller is responsible for freeing the
    // returned BSTR.
    BSTR getBSTR() const;
#endif
};

#endif
