// $Id: TclObject.cpp 18 2005-05-03 00:40:40Z cthuang $
#include "TclObject.h"
#include <vector>
#ifdef WIN32
#include "Extension.h"
#include "Reference.h"
#endif

Tcl_ObjType *TclTypes::ms_pBooleanType;
Tcl_ObjType *TclTypes::ms_pDoubleType;
Tcl_ObjType *TclTypes::ms_pIntType;
Tcl_ObjType *TclTypes::ms_pListType;
#if TCL_MINOR_VERSION >= 1
Tcl_ObjType *TclTypes::ms_pByteArrayType;
#endif

void
TclTypes::initialize ()
{
    // Don't worry about multiple threads initializing this data because they
    // should all produce the same result anyway.
    ms_pBooleanType = Tcl_GetObjType("boolean");
    ms_pDoubleType = Tcl_GetObjType("double");
    ms_pIntType = Tcl_GetObjType("int");
    ms_pListType = Tcl_GetObjType("list");
#if TCL_MINOR_VERSION >= 1
    ms_pByteArrayType = Tcl_GetObjType("bytearray");
#endif
}


TclObject::TclObject ():
    m_pObj(Tcl_NewObj())
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (const TclObject &rhs):
    m_pObj(rhs.m_pObj)
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (Tcl_Obj *pObj):
    m_pObj(pObj)
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (const char *src, int len):
    m_pObj(Tcl_NewStringObj(const_cast<char *>(src), len))
{ Tcl_IncrRefCount(m_pObj); }

#if TCL_MINOR_VERSION >= 2
TclObject::TclObject (const wchar_t *src, int len):
    m_pObj(Tcl_NewUnicodeObj(
	const_cast<Tcl_UniChar *>(reinterpret_cast<const Tcl_UniChar *>(src)),
	len))
{ Tcl_IncrRefCount(m_pObj); }

static Tcl_Obj *
newUnicodeObj (const Tcl_UniChar *pWide, int length)
{
    if (pWide == 0) {
        return Tcl_NewObj();
    }
    return Tcl_NewUnicodeObj(const_cast<Tcl_UniChar *>(pWide), length);
}
#endif

TclObject::TclObject (const std::string &s):
    m_pObj(Tcl_NewStringObj(const_cast<char *>(s.data()), s.size()))
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (bool value):
    m_pObj(Tcl_NewBooleanObj(static_cast<int>(value)))
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (int value):
    m_pObj(Tcl_NewIntObj(value))
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (long value):
    m_pObj(Tcl_NewLongObj(value))
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (double value):
    m_pObj(Tcl_NewDoubleObj(value))
{ Tcl_IncrRefCount(m_pObj); }

TclObject::TclObject (int objc, Tcl_Obj *CONST objv[]):
    m_pObj(Tcl_NewListObj(objc, objv))
{ Tcl_IncrRefCount(m_pObj); }

TclObject::~TclObject ()
{ Tcl_DecrRefCount(m_pObj); }

TclObject &
TclObject::operator= (const TclObject &rhs)
{
    Tcl_IncrRefCount(rhs.m_pObj);
    Tcl_DecrRefCount(m_pObj);
    m_pObj = rhs.m_pObj;
    return *this;
}

TclObject &
TclObject::operator= (Tcl_Obj *pObj)
{
    Tcl_IncrRefCount(pObj);
    Tcl_DecrRefCount(m_pObj);
    m_pObj = pObj;
    return *this;
}

bool
TclObject::getBool () const
{
    int value;
    Tcl_GetBooleanFromObj(0, m_pObj, &value);
    return value != 0;
}

int
TclObject::getInt () const
{
    int value;
    Tcl_GetIntFromObj(0, m_pObj, &value);
    return value;
}

long
TclObject::getLong () const
{
    long value;
    Tcl_GetLongFromObj(0, m_pObj, &value);
    return value;
}

double
TclObject::getDouble () const
{
    double value;
    Tcl_GetDoubleFromObj(0, m_pObj, &value);
    return value;
}

TclObject &
TclObject::lappend (Tcl_Obj *pElement)
{
    if (Tcl_IsShared(m_pObj)) {
        Tcl_DecrRefCount(m_pObj);
        m_pObj = Tcl_DuplicateObj(m_pObj);
        Tcl_IncrRefCount(m_pObj);
    }
    Tcl_ListObjAppendElement(NULL, m_pObj, pElement);
    // TODO: Should check for error result if conversion to list failed.
    return *this;
}

#ifdef WIN32
// Convert SAFEARRAY to a Tcl value.

static Tcl_Obj *
convertFromSafeArray (
    SAFEARRAY *psa,
    VARTYPE elementType,
    unsigned dim,
    long *pIndices,
    const Type &type,
    Tcl_Interp *interp)
{
    HRESULT hr;

    // Get index range.
    long lowerBound;
    hr = SafeArrayGetLBound(psa, dim, &lowerBound);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    long upperBound;
    hr = SafeArrayGetUBound(psa, dim, &upperBound);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    Tcl_Obj *pResult;
    if (dim < SafeArrayGetDim(psa)) {
        // Create list of list for multi-dimensional array.
        pResult = Tcl_NewListObj(0, 0);
        for (long i = lowerBound; i <= upperBound; ++i) {
            pIndices[dim - 1] = i;
            Tcl_Obj *pElement = convertFromSafeArray(
                psa, elementType, dim + 1, pIndices, type, interp);
            Tcl_ListObjAppendElement(interp, pResult, pElement);
        }
        return pResult;
    }

    if (elementType == VT_UI1 && SafeArrayGetDim(psa) == 1) {
        unsigned char *pData;
        hr = SafeArrayAccessData(psa, reinterpret_cast<void **>(&pData));
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

        long length = upperBound - lowerBound + 1;
        pResult =
#if TCL_MINOR_VERSION >= 1
            // Convert array of bytes to Tcl byte array.
            Tcl_NewByteArrayObj(pData, length);
#else
            // Convert array of bytes to Tcl string.
            Tcl_NewStringObj(reinterpret_cast<char *>(pData), length);
#endif

        hr = SafeArrayUnaccessData(psa);
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

    } else {
        // Create list of Tcl values.
        pResult = Tcl_NewListObj(0, 0);
        for (long i = lowerBound; i <= upperBound; ++i) {
            NativeValue elementVar;

            pIndices[dim - 1] = i;
            if (elementType == VT_VARIANT) {
                hr = SafeArrayGetElement(psa, pIndices, &elementVar);
            } else {
                // I hope the element can be contained in a VARIANT.
                V_VT(&elementVar) = elementType;
                hr = SafeArrayGetElement(psa, pIndices, &elementVar.punkVal);
            }
            if (FAILED(hr)) {
                _com_issue_error(hr);
            }

            TclObject element(&elementVar, type, interp);
            Tcl_ListObjAppendElement(interp, pResult, element);
        }
    }

    return pResult;
}

// Fill SAFEARRAY from Tcl list.

static void
fillSafeArray (
    Tcl_Obj *pList,
    SAFEARRAY *psa,
    unsigned dim,
    long *pIndices,
    Tcl_Interp *interp,
    bool addRef)
{
    HRESULT hr;

    // Get index range.
    long lowerBound;
    hr = SafeArrayGetLBound(psa, dim, &lowerBound);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    long upperBound;
    hr = SafeArrayGetUBound(psa, dim, &upperBound);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    int numElements;
    Tcl_Obj **pElements;
    if (Tcl_ListObjGetElements(interp, pList, &numElements, &pElements)
        != TCL_OK) {
        _com_issue_error(E_INVALIDARG);
    }

    unsigned dim1 = dim - 1;
    if (dim < SafeArrayGetDim(psa)) {
        // Create list of list for multi-dimensional array.
        for (int i = 0; i < numElements; ++i) {
            pIndices[dim1] = i;
            fillSafeArray(pElements[i], psa, dim + 1, pIndices, interp, addRef);
        }

    } else {
        for (int i = 0; i < numElements; ++i) {
            TclObject element(pElements[i]); 
            NativeValue elementVar;
            element.toNativeValue(&elementVar, Type::variant(), interp, addRef);

            pIndices[dim1] = i;
            hr = SafeArrayPutElement(psa, pIndices, &elementVar);
            if (FAILED(hr)) {
                _com_issue_error(hr);
            }
        }
    }
}

static Tcl_Obj *
convertFromUnknown (IUnknown *pUnknown, REFIID iid, Tcl_Interp *interp)
{
    if (pUnknown == 0) {
        return Tcl_NewObj();
    }

    const Interface *pInterface = InterfaceManager::instance().find(iid);
    return Extension::referenceHandles.newObj(
        interp,
        Reference::newReference(pUnknown, pInterface));
}

TclObject::TclObject (VARIANT *pSrc, const Type &type, Tcl_Interp *interp)
{
    if (V_ISARRAY(pSrc)) {
        SAFEARRAY *psa = V_ISBYREF(pSrc) ? *V_ARRAYREF(pSrc) : V_ARRAY(pSrc);
        VARTYPE elementType = V_VT(pSrc) & VT_TYPEMASK;
        unsigned numDimensions = SafeArrayGetDim(psa);
        std::vector<long> indices(numDimensions);
        m_pObj = convertFromSafeArray(
            psa, elementType, 1, &indices[0], type, interp);

    } else if (vtMissing == pSrc) {
        m_pObj = Extension::newNaObj();

    } else {
        switch (V_VT(pSrc)) {
        case VT_BOOL:
            m_pObj = Tcl_NewBooleanObj(V_BOOL(pSrc));
            break;

        case VT_ERROR:
            m_pObj = Tcl_NewLongObj(V_ERROR(pSrc));
            break;

        case VT_I1:
        case VT_UI1:
            m_pObj = Tcl_NewLongObj(V_I1(pSrc));
            break;

        case VT_I2:
        case VT_UI2:
            m_pObj = Tcl_NewLongObj(V_I2(pSrc));
            break;

        case VT_I4:
        case VT_UI4:
        case VT_INT:
        case VT_UINT:
            m_pObj = Tcl_NewLongObj(V_I4(pSrc));
            break;

#ifdef V_I8
        case VT_I8:
        case VT_UI8:
            m_pObj = Tcl_NewWideIntObj(V_I8(pSrc));
            break;
#endif

        case VT_R4:
            m_pObj = Tcl_NewDoubleObj(V_R4(pSrc));
            break;

        case VT_DATE:
        case VT_R8:
            m_pObj = Tcl_NewDoubleObj(V_R8(pSrc));
            break;

        case VT_DISPATCH:
            m_pObj = convertFromUnknown(V_DISPATCH(pSrc), type.iid(), interp);
            break;

        case VT_DISPATCH | VT_BYREF:
            m_pObj = convertFromUnknown(
                (V_DISPATCHREF(pSrc) != 0) ? *V_DISPATCHREF(pSrc) : 0,
                type.iid(),
                interp);
            break;

        case VT_UNKNOWN:
            m_pObj = convertFromUnknown(V_UNKNOWN(pSrc), type.iid(), interp);
            break;

        case VT_UNKNOWN | VT_BYREF:
            m_pObj = convertFromUnknown(
                (V_UNKNOWNREF(pSrc) != 0) ? *V_UNKNOWNREF(pSrc) : 0,
                type.iid(),
                interp);
            break;

        case VT_NULL:
            m_pObj = Extension::newNullObj();
            break;

        case VT_LPSTR:
            m_pObj = Tcl_NewStringObj(V_I1REF(pSrc), -1);
            break;

        case VT_LPWSTR:
            {
#if TCL_MINOR_VERSION >= 2
                // Uses Unicode function introduced in Tcl 8.2.
                m_pObj = newUnicodeObj(V_UI2REF(pSrc), -1);
#else
		const wchar_t *pWide = V_UI2REF(pSrc);
                _bstr_t str(pWide);
                m_pObj = Tcl_NewStringObj(str, -1);
#endif
            }
            break;

        default:
            if (V_VT(pSrc) == VT_USERDEFINED && type.name() == "GUID") {
                Uuid uuid(*static_cast<UUID *>(V_BYREF(pSrc)));
                m_pObj = Tcl_NewStringObj(
                    const_cast<char *>(uuid.toString().c_str()), -1);
            } else {
                if (V_VT(pSrc) == (VT_VARIANT | VT_BYREF)) {
                    pSrc = V_VARIANTREF(pSrc);
                }

                _bstr_t str(pSrc);
#if TCL_MINOR_VERSION >= 2
                // Uses Unicode function introduced in Tcl 8.2.
		wchar_t *pWide = str;
                m_pObj = newUnicodeObj(
                    reinterpret_cast<Tcl_UniChar *>(pWide), str.length());
#else
                m_pObj = Tcl_NewStringObj(str, -1);
#endif
            }
        }
    }

    Tcl_IncrRefCount(m_pObj);
}

TclObject::TclObject (const _bstr_t &src)
{
    if (src.length() > 0) {
#if TCL_MINOR_VERSION >= 2
        // Uses Unicode functions introduced in Tcl 8.2.
        m_pObj = Tcl_NewUnicodeObj(src, -1);
#else
        m_pObj = Tcl_NewStringObj(src, -1);
#endif
    } else {
        m_pObj = Tcl_NewObj();
    }

    Tcl_IncrRefCount(m_pObj);
}

TclObject::TclObject (
    SAFEARRAY *psa, const Type &type, Tcl_Interp *interp)
{
    unsigned numDimensions = SafeArrayGetDim(psa);
    std::vector<long> indices(numDimensions);
    m_pObj = convertFromSafeArray(
        psa, type.elementType().vartype(), 1, &indices[0], type, interp);

    Tcl_IncrRefCount(m_pObj);
}

BSTR
TclObject::getBSTR () const
{
#if TCL_MINOR_VERSION >= 2
    // Uses Unicode function introduced in Tcl 8.2.
    return SysAllocString(getUnicode());
#else
    _bstr_t str(c_str());
    return SysAllocString(str);
#endif
}

#if TCL_MINOR_VERSION >= 1
// Convert Tcl byte array to SAFEARRAY of bytes.

static SAFEARRAY *
newByteSafeArray (Tcl_Obj *pObj)
{
    int length;
    unsigned char *pSrc = Tcl_GetByteArrayFromObj(pObj, &length);

    SAFEARRAY *psa = SafeArrayCreateVector(VT_UI1, 0, length);
    if (psa == 0) {
        _com_issue_error(E_OUTOFMEMORY);
    }

    unsigned char *pDest;
    HRESULT hr;
    hr = SafeArrayAccessData(
        psa, reinterpret_cast<void **>(&pDest));
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    memcpy(pDest, pSrc, length);

    hr = SafeArrayUnaccessData(psa);
    if (FAILED(hr)) {
        _com_issue_error(hr);
    }

    return psa;
}
#endif

SAFEARRAY *
TclObject::getSafeArray (const Type &elementType, Tcl_Interp *interp) const
{
    SAFEARRAY *psa;

    if (elementType.vartype() == VT_UI1) {
        psa = newByteSafeArray(m_pObj);
    } else {
        // Convert Tcl list to SAFEARRAY.
        int numElements;
        Tcl_Obj **pElements;
        if (Tcl_ListObjGetElements(interp, m_pObj, &numElements, &pElements)
          != TCL_OK) {
            _com_issue_error(E_INVALIDARG);
        }

        psa = SafeArrayCreateVector(elementType.vartype(), 0, numElements);
        if (psa == 0) {
            _com_issue_error(E_OUTOFMEMORY);
        }

        void *pData;
        HRESULT hr;
        hr = SafeArrayAccessData(psa, &pData);
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }

        for (int i = 0; i < numElements; ++i) {
            TclObject value(pElements[i]);

            switch (elementType.vartype()) {
            case VT_BOOL:
                static_cast<VARIANT_BOOL *>(pData)[i] =
                    value.getBool() ? VARIANT_TRUE : VARIANT_FALSE;
                break;

            case VT_I2:
            case VT_UI2:
                static_cast<short *>(pData)[i] = value.getLong();
                break;

            case VT_R4:
                static_cast<float *>(pData)[i] =
                    static_cast<float>(value.getDouble());
                break;

            case VT_R8:
                static_cast<double *>(pData)[i] = value.getDouble();
                break;

            case VT_BSTR:
                static_cast<BSTR *>(pData)[i] = value.getBSTR();
                break;

            case VT_VARIANT:
                {
                    VARIANT *pDest = static_cast<VARIANT *>(pData) + i;
                    VariantInit(pDest);
                    value.toVariant(pDest, elementType, interp);
                }
                break;

            default:
                static_cast<int *>(pData)[i] = value.getLong();
            }
        }

        hr = SafeArrayUnaccessData(psa);
        if (FAILED(hr)) {
            _com_issue_error(hr);
        }
    }

    return psa;
}

void
TclObject::toVariant (VARIANT *pDest,
                      const Type &type,
                      Tcl_Interp *interp,
                      bool addRef)
{
    VariantClear(pDest);
    VARTYPE vt = type.vartype();

    Reference *pReference = Extension::referenceHandles.find(interp, m_pObj);
    if (pReference != 0) {
        // Convert interface pointer handle to interface pointer.
        if (addRef) {
            // Must increment reference count of interface pointers returned
            // from methods.
            pReference->unknown()->AddRef();
        }

        IDispatch *pDispatch = pReference->dispatch();
        if (pDispatch != 0) {
            V_VT(pDest) = VT_DISPATCH;
            V_DISPATCH(pDest) = pDispatch;
        } else {
            V_VT(pDest) = VT_UNKNOWN;
            V_UNKNOWN(pDest) = pReference->unknown();
        }

    } else if (m_pObj->typePtr == &Extension::unknownPointerType) {
        // Convert to interface pointer.
        IUnknown *pUnknown = static_cast<IUnknown *>(
            m_pObj->internalRep.otherValuePtr);
        if (addRef && pUnknown != 0) {
            // Must increment reference count of interface pointers returned
            // from methods.
            pUnknown->AddRef();
        }
        V_VT(pDest) = VT_UNKNOWN;
        V_UNKNOWN(pDest) = pUnknown;

    } else if (vt == VT_SAFEARRAY) {

        const Type &elementType = type.elementType();
        V_VT(pDest) = VT_ARRAY | elementType.vartype();
        V_ARRAY(pDest) = getSafeArray(elementType, interp);

    } else if (m_pObj->typePtr == TclTypes::listType()) {
        // Convert Tcl list to array of VARIANT.
        int numElements;
        Tcl_Obj **pElements;
        if (Tcl_ListObjGetElements(interp, m_pObj, &numElements, &pElements)
          != TCL_OK) {
            _com_issue_error(E_INVALIDARG);
        }

        SAFEARRAYBOUND bounds[2];
        bounds[0].cElements = numElements;
        bounds[0].lLbound = 0;

        unsigned numDimensions;

        // Check if the first element of the list is a list.
        if (numElements > 0 && pElements[0]->typePtr == TclTypes::listType()) {
            int colSize;
            Tcl_Obj **pCol;
            if (Tcl_ListObjGetElements(interp, pElements[0], &colSize, &pCol)
             != TCL_OK) {
                _com_issue_error(E_INVALIDARG);
            }

            bounds[1].cElements = colSize;
            bounds[1].lLbound = 0;
            numDimensions = 2;
        } else {
            numDimensions = 1;
        }

        SAFEARRAY *psa = SafeArrayCreate(VT_VARIANT, numDimensions, bounds);
        std::vector<long> indices(numDimensions);
        fillSafeArray(m_pObj, psa, 1, &indices[0], interp, addRef);

        V_VT(pDest) = VT_ARRAY | VT_VARIANT;
        V_ARRAY(pDest) = psa;

#if TCL_MINOR_VERSION >= 1
    } else if (m_pObj->typePtr == TclTypes::byteArrayType()) {
        // Convert Tcl byte array to SAFEARRAY of bytes.

        V_VT(pDest) = VT_ARRAY | VT_UI1;
        V_ARRAY(pDest) = newByteSafeArray(m_pObj);
#endif

    } else if (m_pObj->typePtr == &Extension::naType) {
        // This variant indicates a missing optional argument.
        VariantCopy(pDest, &vtMissing);

    } else if (m_pObj->typePtr == &Extension::nullType) {
        V_VT(pDest) = VT_NULL;

    } else if (m_pObj->typePtr == &Extension::variantType) {
        VariantCopy(
            pDest,
            static_cast<_variant_t *>(m_pObj->internalRep.otherValuePtr));

    } else if (m_pObj->typePtr == TclTypes::intType()) {
        long value;
        if (Tcl_GetLongFromObj(interp, m_pObj, &value) != TCL_OK) {
            value = 0;
        }
        V_VT(pDest) = VT_I4;
        V_I4(pDest) = value;

        if (vt != VT_VARIANT && vt != VT_USERDEFINED) {
            VariantChangeType(pDest, pDest, 0, vt);
        }

    } else if (m_pObj->typePtr == TclTypes::doubleType()) {
        double value;
        if (Tcl_GetDoubleFromObj(interp, m_pObj, &value) != TCL_OK) {
            value = 0.0;
        }
        V_VT(pDest) = VT_R8;
        V_R8(pDest) = value;

        if (vt != VT_VARIANT && vt != VT_USERDEFINED) {
            VariantChangeType(pDest, pDest, 0, vt);
        }

    } else if (m_pObj->typePtr == TclTypes::booleanType()) {
        int value;
        if (Tcl_GetBooleanFromObj(interp, m_pObj, &value) != TCL_OK) {
            value = 0;
        }
        V_VT(pDest) = VT_BOOL;
        V_BOOL(pDest) = (value != 0) ? VARIANT_TRUE : VARIANT_FALSE;

        if (vt != VT_VARIANT && vt != VT_USERDEFINED) {
            VariantChangeType(pDest, pDest, 0, vt);
        }

    } else if (vt == VT_BOOL) {
        V_VT(pDest) = VT_BOOL;
        V_BOOL(pDest) = getBool() ? VARIANT_TRUE : VARIANT_FALSE;

    } else {
        V_VT(pDest) = VT_BSTR;
        V_BSTR(pDest) = getBSTR();

        // If trying to convert from a string to a date,
        // we need to convert to a double (VT_R8) first.
        if (vt == VT_DATE) {
            VariantChangeType(pDest, pDest, 0, VT_R8);
        }

        // Try to convert from a string representation.
        if (vt != VT_VARIANT && vt != VT_USERDEFINED && vt != VT_LPWSTR) {
            VariantChangeType(pDest, pDest, 0, vt);
        }
    }
}


void
TclObject::toNativeValue (NativeValue *pDest,
                          const Type &type,
                          Tcl_Interp *interp,
                          bool addRef)
{
#ifdef V_I8
    VARTYPE vt = type.vartype();
    if (vt == VT_I8 || vt == VT_UI8) {
        pDest->fixInvalidVariantType();
        VariantClear(pDest);
        V_VT(pDest) = vt;
        Tcl_GetWideIntFromObj(interp, m_pObj, &V_I8(pDest));
        return;
    }
#endif

    pDest->fixInvalidVariantType();
    toVariant(pDest, type, interp, addRef);
}

#endif
