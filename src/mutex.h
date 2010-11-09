// $Id: mutex.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef MUTEX_H
#define MUTEX_H

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

// This class is used for mutual-exclusion synchronization.

class CriticalSectionMutex
{
    CRITICAL_SECTION m_cs;

    // Disallow others from copying instances of this class.
    CriticalSectionMutex(const CriticalSectionMutex &); // not implemented
    void operator=(const CriticalSectionMutex &);       // not implemented
    
public:
    CriticalSectionMutex ()
    { InitializeCriticalSection(&m_cs); }

    ~CriticalSectionMutex ()
    { DeleteCriticalSection(&m_cs); }

    void enter ()
    { EnterCriticalSection(&m_cs); }

    void leave ()
    { LeaveCriticalSection(&m_cs); }
};

// This class mirrors the operations of a mutex except the operations do
// nothing.

class FakeMutex
{
public:
    void enter ()
    { }

    void leave ()
    { }
};

#ifdef TCL_THREADS
typedef CriticalSectionMutex Mutex;
#else
typedef FakeMutex Mutex;
#endif

// This class locks a mutex when constructed and unlocks it when destroyed.

class SingleLock
{
    Mutex &m_mutex;

public:
    SingleLock (Mutex &mutex):
        m_mutex(mutex)
    { m_mutex.enter(); }

    ~SingleLock ()
    { m_mutex.leave(); }
};

#ifdef TCL_THREADS
#define LOCK_MUTEX(mutex) \
    SingleLock criticalSectionLock(const_cast<Mutex &>(mutex));
#else
#define LOCK_MUTEX(mutex)
#endif

#endif
