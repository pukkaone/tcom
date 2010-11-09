// $Id: TypeInfo.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef TYPEINFO_H
#define TYPEINFO_H

#include <vector>
#include <string>
#include <comdef.h>
#include "tcomApi.h"
#include "Uuid.h"
#include "HashTable.h"
#include "Singleton.h"

// Define smart pointer class named ITypeInfoPtr which automatically calls the
// IUnknown methods.
_COM_SMARTPTR_TYPEDEF(ITypeInfo, __uuidof(ITypeInfo));

// Get name of type described by the ITypeInfo object.
std::string getTypeInfoName(ITypeInfo *pTypeInfo, MEMBERID id=MEMBERID_NIL);

// This wrapper class takes ownership of the resource and is responsible for
// releasing it.

class TCOM_API TypeAttr
{
    ITypeInfoPtr m_pTypeInfo;
    TYPEATTR *m_pTypeAttr;

    // Do not allow others to copy instances of this class.
    TypeAttr(const TypeAttr &);         // not implemented
    void operator=(const TypeAttr &);   // not implemented

public:
    // TODO: I should probably check for a failure result and throw an
    // exception in that case.
    TypeAttr (const ITypeInfoPtr &pTypeInfo):
        m_pTypeInfo(pTypeInfo)
    { m_pTypeInfo->GetTypeAttr(&m_pTypeAttr); }

    ~TypeAttr ()
    { m_pTypeInfo->ReleaseTypeAttr(m_pTypeAttr); }

    // Deference pointer.
    TYPEATTR *operator-> () const
    { return m_pTypeAttr; }
};

// This wrapper class takes ownership of the resource and is responsible for
// releasing it.

class TCOM_API FuncDesc
{
    ITypeInfoPtr m_pTypeInfo;
    FUNCDESC *m_pFuncDesc;

    // Do not allow others to copy instances of this class.
    FuncDesc(const FuncDesc &);         // not implemented
    void operator=(const FuncDesc &);   // not implemented

public:
    FuncDesc(const ITypeInfoPtr &pTypeInfo, unsigned index);

    ~FuncDesc ()
    { m_pTypeInfo->ReleaseFuncDesc(m_pFuncDesc); }

    // Dereference pointer.
    FUNCDESC *operator-> () const
    { return m_pFuncDesc; }
};

// This wrapper class takes ownership of the resource and is responsible for
// releasing it.

class TCOM_API VarDesc
{
    ITypeInfoPtr m_pTypeInfo;
    VARDESC *m_pVarDesc;

    // Do not allow others to copy instances of this class.
    VarDesc(const VarDesc &);           // not implemented
    void operator=(const VarDesc &);    // not implemented

public:
    VarDesc(const ITypeInfoPtr &pTypeInfo, unsigned index);

    ~VarDesc ()
    { m_pTypeInfo->ReleaseVarDesc(m_pVarDesc); }

    // Dereference pointer.
    VARDESC *operator-> () const
    { return m_pVarDesc; }
};

// This class describes a type for a function return value or parameter.

class TCOM_API Type
{
    std::string m_name;
    VARTYPE m_vt;
    IID m_iid;
    unsigned m_pointerCount;
    Type *m_pElementType;       // element type for arrays

    // description for VARIANT type
    static Type ms_variant;

    void readTypeDesc(const ITypeInfoPtr &pTypeInfo, TYPEDESC &typeDesc);

public:
    Type(const ITypeInfoPtr &pTypeInfo, TYPEDESC &typeDesc);
    Type(const std::string &str);
    Type(const Type &rhs);

    ~Type();

    Type &operator=(const Type &rhs);

    const std::string &name () const
    { return m_name; }

    VARTYPE vartype () const
    { return m_vt; }

    const IID &iid () const
    { return m_iid; }

    unsigned pointerCount () const
    { return m_pointerCount; }

    const Type &elementType () const
    { return *m_pElementType; }

    // Get string representation.
    std::string toString() const;

    // Get description for VARIANT type.
    static Type &variant()
    { return ms_variant; }
};

// This class describes a function parameter.

class TCOM_API Parameter
{
    std::string m_name;
    Type m_type;
    unsigned m_flags;

public:
    Parameter(const ITypeInfoPtr &pTypeInfo, ELEMDESC &elemDesc, const char *name);

    Parameter(
        const std::string &modes,
        const std::string &type,
        const std::string &name);

    Parameter (unsigned flags, const Type &type, const char *name):
        m_flags(flags),
        m_type(type),
        m_name(name)
    { }

    const std::string &name () const
    { return m_name; }

    const Type &type () const
    { return m_type; }

    unsigned flags () const
    { return m_flags; }
};

// This class describes a method.

class TCOM_API Method
{
public:
    typedef std::vector<Parameter> Parameters;

private:
    std::string m_name;
    Type m_type;
    Parameters m_parameters;
    MEMBERID m_memberid;
    INVOKEKIND m_invokeKind;
    short m_vtblIndex;      // position in virtual function table
    bool m_vararg;          // method accepts variable number of arguments

protected:
    Method(const ITypeInfoPtr &pTypeInfo, DISPID memberid, ELEMDESC &elemDesc);

public:
    Method(const ITypeInfoPtr &pTypeInfo, FuncDesc &funcDesc);

    Method(
        MEMBERID memberid,
        const std::string &type,
        const std::string &name,
        INVOKEKIND invokeKind=INVOKE_FUNC);

    // Get method name.
    const std::string &name () const
    { return m_name; }

    // Get return value type.
    const Type &type () const
    { return m_type; }

    // Set return value type.
    void type (const Type &rhs)
    { m_type = rhs; }

    // Insert parameter.
    void addParameter (const Parameter &parameter)
    { m_parameters.push_back(parameter); }

    // Get parameters.
    const Parameters &parameters () const
    { return m_parameters; }

    // Get member ID.
    MEMBERID memberid () const
    { return m_memberid; }

    // Get indicator whether this is a method or property.
    INVOKEKIND invokeKind () const
    { return m_invokeKind; }

    // Set indicator whether this is a method or property.
    void invokeKind (INVOKEKIND invokeKind)
    { m_invokeKind = invokeKind; }

    // Get position in virtual function table.
    short vtblIndex () const
    { return m_vtblIndex; }

    // Return true if the method accepts variable number of arguments.
    bool vararg () const
    { return m_vararg; }
};

// This class describes a property.

class TCOM_API Property: public Method
{
    WORD m_putDispatchFlag;
    bool m_readOnly;

    // Initialize data members.
    void initialize();

public:
    Property(const ITypeInfoPtr &pTypeInfo, FuncDesc &funcDesc);
    Property(const ITypeInfoPtr &pTypeInfo, VarDesc &varDesc);

    Property(
        MEMBERID memberid,
        const std::string &modes,
        const std::string &type,
        const std::string &name);

    // Merge data from other property into this one.
    void merge(const Property &property);

    // Get dispatch flag for setting property.
    WORD putDispatchFlag () const
    { return m_putDispatchFlag; }

    // Get read only flag.
    bool readOnly () const
    { return m_readOnly; }
};

// This describes an interface.

class TCOM_API Interface
{
    // Make the following a friend so it can construct instances of this class.
    friend class InterfaceManager;

public:
    typedef std::vector<Method> Methods;
    typedef std::vector<Property> Properties;

private:
    ITypeInfoPtr m_pTypeInfo;
    IID m_iid;
    std::string m_name;
    Methods m_methods;
    Properties m_properties;
    bool m_dispatchOnly;
    bool m_dispatchable;

    // Get information on methods and properties.
    void readFunctions(const ITypeInfoPtr &pTypeInfo, TypeAttr &typeAttr);

    // Get information on interface.
    void readTypeInfo(const ITypeInfoPtr &pTypeInfo);

    // Constructors
    Interface (const ITypeInfoPtr &pTypeInfo):
        m_pTypeInfo(pTypeInfo)
    { readTypeInfo(pTypeInfo); }

    Interface (REFIID iid, const char *name):
        m_iid(iid),
        m_name(name)
    { }

public:
    // Get IID.
    const IID &iid () const
    { return m_iid; }

    // Get string representatin of IID.
    std::string iidString() const;

    // Get name.
    const std::string &name () const
    { return m_name; }

    // Get type info description.
    ITypeInfo *typeInfo () const
    { return m_pTypeInfo.GetInterfacePtr(); }

    // Insert method information.
    void addMethod(const Method &method);

    // Get methods.
    const Methods &methods () const
    { return m_methods; }

    // Find the named method.
    const Method *findMethod(const char *name) const;

    // Insert property information.
    void addProperty(const Property &property);

    // Get properties.
    const Properties &properties () const
    { return m_properties; }

    // Find the named property.
    const Property *findProperty(const char *name) const;

    // Return true if this interface can only be invoked through IDispatch.
    bool dispatchOnly () const
    { return m_dispatchOnly; }

    // Return true if this interface derives from IDispatch.
    bool dispatchable () const
    { return m_dispatchable; }
};

// This is a cache of interface descriptions.

class TCOM_API InterfaceManager
{
    // used to synchronize access to hash table
    Mutex m_mutex;

    // IID to interface description map
    typedef HashTable<IID, Interface *> IidToInterfaceDescMap;
    IidToInterfaceDescMap m_hashTable;

    friend class Singleton<InterfaceManager>;
    static Singleton<InterfaceManager> ms_singleton;

    // Do not allow others to create or copy instances of this class.
    InterfaceManager();
    ~InterfaceManager();
    InterfaceManager(const InterfaceManager &);     // not implemented
    void operator=(const InterfaceManager &);       // not implemented

public:
    // Get singleton instance.
    static InterfaceManager &instance();

    // If the interface description already exists, return it,
    // otherwise create a new interface description.
    const Interface *newInterface(REFIID iid, const ITypeInfoPtr &pTypeInfo);

    // If the interface description already exists, return it,
    // otherwise create a new interface description.
    Interface *newInterface(REFIID iid, const char *name);

    // Look for the interface description.
    const Interface *find(REFIID iid) const;
};

// This adapts an interface description so we can create a handle for it.
// We need this because the handle support classes take ownership of the
// application objects passed to them and will delete them, but the
// InterfaceManager maintains the life cycle of interface descriptions.

class TCOM_API InterfaceHolder
{
    const Interface *m_pInterface;

public:
    InterfaceHolder (const Interface *pInterface):
        m_pInterface(pInterface)
    { }

    const Interface *interfaceDesc () const
    { return m_pInterface; }
};

// This describes a class.

class Class
{
public:
    typedef std::vector<const Interface *> Interfaces;

private:
    std::string m_name;
    CLSID m_clsid;
    Interfaces m_interfaces;
    const Interface *m_pDefaultInterface;
    const Interface *m_pSourceInterface;

public:
    Class(const ITypeInfoPtr &pTypeInfo);
    Class(
        const char *name,
        const CLSID &clsid,
        const Interface *pDefaultInterface,
        const Interface *pSourceInterface);

    // Get name.
    const std::string &name () const
    { return m_name; }

    // Get CLSID.
    const CLSID &clsid () const
    { return m_clsid; }

    // Get CLSID as string.
    std::string clsidString () const
    { Uuid uuid(m_clsid); return uuid.toString(); }

    // Get interfaces this class implements.
    const Interfaces &interfaces () const
    { return m_interfaces; }

    // Get default interface.
    const Interface *defaultInterface () const
    { return m_pDefaultInterface; }

    // Get default source interface.
    const Interface *sourceInterface () const
    { return m_pSourceInterface; }
};

#endif 
