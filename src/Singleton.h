// $Id: Singleton.h 5 2005-02-16 14:57:24Z cthuang $
#ifndef SINGLETON_H
#define SINGLETON_H

#include <tcl.h>
#include "mutex.h"

// This template class provides code to construct and destroy a singleton.

template<class T>
class Singleton
{
    // singleton instance
    static T *ms_pInstance;

    // used to synchronize construction of singleton instance
    static Mutex ms_singletonMutex;

    // Delete the instance when exiting.
    static void exitProc(ClientData clientData);

public:
    // Get instance.
    static T &instance();
};

template<class T>
T *Singleton<T>::ms_pInstance = 0;

template<class T>
Mutex Singleton<T>::ms_singletonMutex;

template<class T>
void
Singleton<T>::exitProc (ClientData clientData)
{
    delete reinterpret_cast<T *>(clientData);
}

template<class T>
T &
Singleton<T>::instance ()
{
    if (ms_pInstance == 0) {
        LOCK_MUTEX(ms_singletonMutex)
        if (ms_pInstance == 0) {
            ms_pInstance = new T;

            // Install an exit handler to destroy the instance when Tcl exits
            // instead of depending on the destruction of a static C++ object
            // because the Tcl library may have been finalized before the
            // destructor is called.
            Tcl_CreateExitHandler(exitProc, ms_pInstance);
        }
    }
    return *ms_pInstance;
}

#endif
