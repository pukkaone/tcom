// $Id: TypeLib.h,v 1.21 2002/03/09 16:40:24 cthuang Exp $
#ifndef TYPELIB_H
#define TYPELIB_H

#include <vector>
#include <map>
#include <string>
#include "TypeInfo.h"

// This describes an enumeration with a map where
// the key is an enumerator name and
// the data is the enumerator value.

class Enum: public std::map<std::string, std::string>
{
    std::string m_name;

public:
    Enum(const ITypeInfoPtr &pTypeInfo, TypeAttr &attr);

    // Get name.
    const std::string &name () const
    { return m_name; }
};

// Define smart pointer class named ITypeLibPtr which automatically calls the
// IUnknown methods.
_COM_SMARTPTR_TYPEDEF(ITypeLib, __uuidof(ITypeLib));

// This wrapper class takes ownership of the resource and is responsible for
// releasing it.

class TCOM_API TypeLibAttr
{
    ITypeLibPtr m_pTypeLib;
    TLIBATTR *m_pLibAttr;

    // Do not allow others to copy instances of this class.
    TypeLibAttr(const TypeLibAttr &);       // not implemented
    void operator=(const TypeLibAttr &);    // not implemented

public:
    // TODO: I should probably check for a failure result and throw an
    // exception in that case.
    TypeLibAttr (const ITypeLibPtr &pTypeLib):
        m_pTypeLib(pTypeLib)
    { m_pTypeLib->GetLibAttr(&m_pLibAttr); }

    ~TypeLibAttr ()
    { m_pTypeLib->ReleaseTLibAttr(m_pLibAttr); }

    // Deference pointer.
    TLIBATTR *operator-> () const
    { return m_pLibAttr; }
};

// This class represents a type library.

class TypeLib
{
public:
    typedef std::vector<const Interface *> Interfaces;
    typedef std::vector<Class> Classes;
    typedef std::vector<Enum> Enums;

private:
    ITypeLibPtr m_pTypeLib;
    Interfaces m_interfaces;
    Classes m_classes;
    Enums m_enums;

    TypeLib (const ITypeLibPtr &pTypeLib):
        m_pTypeLib(pTypeLib)
    { readTypeLib(); }

    // Do not allow others to copy instances of this class.
    TypeLib(const TypeLib &);           // not implemented
    void operator=(const TypeLib &);    // not implemented

    // Get information from type library.
    void readTypeLib();

public:
    // Load a type library from the specified file.
    static TypeLib *load(const char *name, bool registerTypeLib=false);

    // Unregister a type library.
    static void unregister(const char *name);

    // Load a type library specified by a LIBID.
    // Return 0 if the required registry entries were not found.
    static TypeLib *loadByLibid(
        const std::string &libid, const std::string &version);

    // Load a type library specified by a CLSID.
    // Return 0 if the required registry entries were not found.
    static TypeLib *loadByClsid(REFCLSID clsid);

    // Load a type library specified by an IID.
    // Return 0 if the required registry entries were not found.
    static TypeLib *loadByIid(REFIID iid);

    // Get string representation of type library ID.
    std::string libidString() const;

    // Get type library version.
    std::string version() const;

    // Get type library name.
    std::string name() const;

    // Get type library documentation string.
    std::string documentation() const;

    // Get interfaces.
    const Interfaces &interfaces () const
    { return m_interfaces; }

    // Get the named interface.
    const Interface *findInterface(const char *name) const;

    // Get classes.
    const Classes &classes () const
    { return m_classes; }

    // Get the named class.
    const Class *findClass(const char *name) const;

    // Find class by CLSID.
    const Class *findClass(REFCLSID clsid) const;

    // Get enumerations.
    const Enums &enums () const
    { return m_enums; }

    // Get the named enumeration.
    const Enum *findEnum(const char *name) const;
};

#endif 
