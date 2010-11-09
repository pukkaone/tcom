// $Id: HashTable.h,v 1.22 2003/07/17 22:33:31 cthuang Exp $
#ifndef HASHTABLE_H
#define HASHTABLE_H

#include <tcl.h>

// Function object that invokes delete on its argument

struct Delete
{
    template<typename T>
    void operator() (T p) const
    { delete p; }
};

// This is a base class used to implement hash tables.

template<typename D>
class BasicHashTable
{
protected:
    Tcl_HashTable m_hashTable;

public:
    BasicHashTable (int keyType)
    { Tcl_InitHashTable(&m_hashTable, keyType); }

    ~BasicHashTable ()
    { Tcl_DeleteHashTable(&m_hashTable); }

    // Remove all elements.
    void clear();

    // Call function on all data elements.
    template<typename F>
    void forEach (F f)
    {
        Tcl_HashSearch search;
        Tcl_HashEntry *pEntry = Tcl_FirstHashEntry(&m_hashTable, &search);
        while (pEntry != 0) {
            Tcl_HashEntry *pNext = Tcl_NextHashEntry(&search);
            f(reinterpret_cast<D>(Tcl_GetHashValue(pEntry)));
            pEntry = pNext;
        }
    }
};

template<typename D>
void
BasicHashTable<D>::clear ()
{
    Tcl_HashSearch search;
    Tcl_HashEntry *pEntry = Tcl_FirstHashEntry(&m_hashTable, &search);
    while (pEntry != 0) {
        Tcl_HashEntry *pNext = Tcl_NextHashEntry(&search);
        Tcl_DeleteHashEntry(pEntry);
        pEntry = pNext;
    }
}

// This class wraps a Tcl hash table that uses structures as keys.  The mapped
// type is assumed to be a pointer type.

template<typename K, typename D>
class HashTable: public BasicHashTable<D>
{
public:
    typedef K key_type;
    typedef D mapped_type;

    HashTable (): BasicHashTable<D>(sizeof(K) / sizeof(int))
    { }

    // Insert data into table.
    void insert(const K &key, D data);

    // Find data in table.
    D find(const K &key) const;

    // Remove data.
    void erase(const K &key);
};

template<typename K, typename D>
void
HashTable<K,D>::insert (const K &key, D value)
{
    int isNew;
    Tcl_HashEntry *pEntry = Tcl_CreateHashEntry(
        &m_hashTable,
        reinterpret_cast<const char *>(&key),
        &isNew);
    Tcl_SetHashValue(pEntry, reinterpret_cast<ClientData>(value));
}

template<typename K, typename D>
D 
HashTable<K,D>::find (const K &key) const
{
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        const_cast<Tcl_HashTable *>(&m_hashTable),
        reinterpret_cast<const char *>(&key));
    if (pEntry == 0) {
        return 0;
    }
    return reinterpret_cast<D>(Tcl_GetHashValue(pEntry));
}

template<typename K, typename D>
void
HashTable<K,D>::erase (const K &key)
{
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        &m_hashTable,
        reinterpret_cast<const char *>(&key));
    if (pEntry != 0) {
        Tcl_DeleteHashEntry(pEntry);
    }
}

#if 0
// This class wraps a Tcl hash table that uses null-terminated strings as keys.
// The mapped type is assumed to be a pointer type.

template<typename D>
class StringHashTable: public BasicHashTable<D>
{
public:
    typedef const char *key_type;
    typedef D mapped_type;

    StringHashTable (): BasicHashTable<D>(TCL_STRING_KEYS)
    { }

    void insert(const char *key, D value);
    D find(const char *key) const;
    void erase(const char *key);
};

template<typename D>
void
StringHashTable<D>::insert (const char *key, D value)
{
    int isNew;
    Tcl_HashEntry *pEntry = Tcl_CreateHashEntry(
        &m_hashTable,
        const_cast<char *>(key),
        &isNew);
    Tcl_SetHashValue(pEntry, reinterpret_cast<ClientData>(value));
}

template<typename D>
D 
StringHashTable<D>::find (const char *key) const
{
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        const_cast<Tcl_HashTable *>(&m_hashTable),
        const_cast<char *>(key));
    if (pEntry == 0) {
        return 0;
    }
    return reinterpret_cast<D>(Tcl_GetHashValue(pEntry));
}

template<typename D>
void
StringHashTable<D>::erase (const char *key)
{
    Tcl_HashEntry *pEntry = Tcl_FindHashEntry(
        &m_hashTable,
        const_cast<char *>(key));
    if (pEntry != 0) {
        Tcl_DeleteHashEntry(pEntry);
    }
}
#endif

#endif
