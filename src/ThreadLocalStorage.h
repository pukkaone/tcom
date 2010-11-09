// $Id: ThreadLocalStorage.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef THREADLOCALSTORAGE_H
#define THREADLOCALSTORAGE_H

#include "mutex.h"

// This factory creates an instance of type T for each calling thread.

template<typename T>
class ThreadLocalStorage
{
    // used to synchronize initialization of index
    Mutex m_mutex;

    DWORD m_index;
    bool m_initialized;

    // not implemented
    ThreadLocalStorage(const ThreadLocalStorage &);
    void operator=(const ThreadLocalStorage &);

public:
    ThreadLocalStorage();
    ~ThreadLocalStorage();

    // Get instance specific to the calling thread.
    T &instance() const;
};

template<typename T>
ThreadLocalStorage<T>::ThreadLocalStorage ()
{
    LOCK_MUTEX(m_mutex)

    if (!m_initialized) {
	m_index = TlsAlloc();
        m_initialized = true;
    }
}

template<typename T>
ThreadLocalStorage<T>::~ThreadLocalStorage ()
{
    LOCK_MUTEX(m_mutex)

    if (m_initialized) {
        TlsFree(m_index);
        m_initialized = false;
    }
}

template<typename T>
T &
ThreadLocalStorage<T>::instance () const
{
    T *pValue = static_cast<T *>(TlsGetValue(m_index));
    if (pValue == 0) {
        pValue = new T;
        TlsSetValue(m_index, pValue);
    }
    return *pValue;
}

#endif
