// $Id: exemain.cpp,v 1.12 2002/07/14 18:42:57 cthuang Exp $
#pragma warning(disable: 4786)
#include "TclModule.h"
#include "tclRunTime.h"

// This class implements a COM module for an EXE server.

class ExeModule: public TclModule
{
    DWORD m_threadId;
    HANDLE m_shutdownEvent;

protected:
    virtual DWORD regclsFlags() const;

public:
    // Increment lock count.
    virtual void lock();

    // Decrement lock count.
    virtual long unlock();

    // Wait for the shutdown event to be raised.
    void waitForShutdown();

    // Start thread waiting for shutdown event.
    bool startMonitor(DWORD threadId);
};

DWORD
ExeModule::regclsFlags () const
{
    return ComModule::regclsFlags() | REGCLS_SUSPENDED;
}

void
ExeModule::lock()
{
    CoAddRefServerProcess();
}

long
ExeModule::unlock()
{
    long count = CoReleaseServerProcess();
    if (count == 0) {
        // Notify monitor to exit application.
        SetEvent(m_shutdownEvent);
    }
    return count;
}

void
ExeModule::waitForShutdown()
{
    WaitForSingleObject(m_shutdownEvent, INFINITE);
    CloseHandle(m_shutdownEvent);
    PostThreadMessage(m_threadId, WM_QUIT, 0, 0);
}

// Passed to CreateThread to monitor the shutdown event.

static DWORD WINAPI
monitorProc (void *pv)
{
    ExeModule *pModule = reinterpret_cast<ExeModule *>(pv);
    pModule->waitForShutdown();
    return 0;
}

bool
ExeModule::startMonitor (DWORD threadId)
{
    m_threadId = threadId;

    m_shutdownEvent = CreateEvent(NULL, false, false, NULL);
    if (m_shutdownEvent == NULL) {
        return false;
    }

    DWORD myThreadId;
    HANDLE h = CreateThread(NULL, 0, monitorProc, this, 0, &myThreadId);
    return h != NULL;
}

extern "C" int WINAPI
WinMain (HINSTANCE /*hInstance*/,
         HINSTANCE /*hPrevInstance*/,
         LPTSTR lpCmdLine,
         int /*nShowCmd*/)
{
    ExeModule module;
    module.startMonitor(GetCurrentThreadId());

    // Get CLSID string from command line.
    std::string cmdLine(lpCmdLine);
    std::string::size_type clsidEnd = cmdLine.find_first_of(" \t");
    std::string clsidStr(cmdLine, 0, clsidEnd);

    // Evaluate script to register class.
    int completionCode = module.registerFactoryByScript(clsidStr);
    if (completionCode != TCL_OK) {
        return completionCode;
    }

    CoResumeClassObjects();

    MSG msg;
    while (GetMessage(&msg, 0, 0, 0)) {
        DispatchMessage(&msg);
    }
    
    module.terminate();

    // Wait for any threads to finish.
    Sleep(1000);

    return 0;
}
