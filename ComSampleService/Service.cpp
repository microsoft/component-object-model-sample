//
// Contains the LocalSystem service code.
//_______________________________________________________________________________


#include "pch.h"
#pragma hdrstop
#include "Infrastructure.h"

class CMpcService
{
public:
    CMpcService();
    ~CMpcService();
    HRESULT ServiceHandler(__in DWORD dwControl,
                           __in DWORD dwEventType,
                           __in LPVOID lpEventData,
                           __in LPVOID lpContext);
    HRESULT ServiceMain();

private:
    HRESULT _StartService();
    HRESULT _StopService();
    HRESULT _SetStatus(__in DWORD dwCurrentState,
                       __in DWORD dwCheckPoint,
                       __in DWORD dwWaitHint);
    static DWORD WINAPI s_ServiceHandler(__in DWORD dwControl,
                                         __in DWORD dwEventType,
                                         __in LPVOID lpEventData,
                                         __in LPVOID lpContext);

private:
    CRITICAL_SECTION        _csLock;
    SERVICE_STATUS_HANDLE   _hServiceStatus;
    SERVICE_STATUS          _serviceStatus;         // Current status of the service.
    HANDLE                  _hServiceStoppingEvent; // Event used to signal that the service is stopping.
};

//  Global pointer for the service instance, only created when we are started by SCM
CMpcService* g_pSvc = nullptr;

//  Service main function; creates a new service class instance and calls into its main function.
void WINAPI ServiceMain(__in DWORD dwArgc, __in_ecount(dwArgc) LPWSTR *pwszArgv)
{
    assert(g_pSvc == nullptr);      // The service instance should only be valid inside service main.

    g_pSvc = new CMpcService();
    HRESULT hr = (g_pSvc != nullptr) ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = g_pSvc->ServiceMain();

        delete g_pSvc;
        g_pSvc = nullptr;
    }
}

// Service Control handler function, a wrapper for service instance's own service handler.
// This is called by SCM after it calls ServiceMain(), so the instance should always be initialized.
// Please note that this always runs in different thread than ServiceMain
DWORD WINAPI CMpcService::s_ServiceHandler(__in DWORD dwControl,
                                           __in DWORD dwEventType,
                                           __in LPVOID lpEventData,
                                           __in LPVOID lpContext)
{
    if (g_pSvc != nullptr)
    {
        (void)g_pSvc->ServiceHandler(dwControl,
                                     dwEventType,
                                     lpEventData,
                                     lpContext);
    }

    return NOERROR;
}

CMpcService::CMpcService()
{
    _hServiceStatus = nullptr;
    _hServiceStoppingEvent = nullptr;

    // Initialize the structure.
    _serviceStatus.dwControlsAccepted        = SERVICE_ACCEPT_STOP;
    _serviceStatus.dwServiceType             = SERVICE_WIN32_OWN_PROCESS;
    _serviceStatus.dwServiceSpecificExitCode = 0;
    _serviceStatus.dwCurrentState            = SERVICE_STOPPED;
    _serviceStatus.dwWin32ExitCode           = NO_ERROR;
    _serviceStatus.dwCheckPoint              = 0;
    _serviceStatus.dwWaitHint                = 0;

    InitializeCriticalSection(&_csLock);
}

CMpcService::~CMpcService()
{
    DeleteCriticalSection(&_csLock);
    //  Do not need to close Service Status Handle
}

//
//  Set the service status, a wrapper for SetServiceStatus().
//  We don't need the critical section here since both service start/stop and
//  service handler are protected by the critical seciton
HRESULT CMpcService::_SetStatus(__in DWORD dwCurrentState,
                                __in DWORD dwCheckPoint,
                                __in DWORD dwWaitHint)
{
    _serviceStatus.dwCurrentState = dwCurrentState;
    _serviceStatus.dwCheckPoint   = dwCheckPoint;
    _serviceStatus.dwWaitHint     = dwWaitHint;

    // Send status of the service to the Service Controller.
    return SetServiceStatus(_hServiceStatus, &_serviceStatus) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
}


//
// Service control handler. For this service, we should be always running (cannot be paused).
// So the only ctrl code we should respond is STOP.
HRESULT CMpcService::ServiceHandler(__in DWORD dwControl,
                                    __in DWORD dwEventType,
                                    __in LPVOID lpEventData,
                                    __in LPVOID lpContext)
{
    HRESULT hr = S_OK;
    EnterCriticalSection(&_csLock);
    {
        // Respond to control code
        switch (dwControl)
        {
            case SERVICE_CONTROL_STOP:
                (void)_SetStatus(SERVICE_STOP_PENDING, 0, 20 * 1000);
                (void)SetEvent(_hServiceStoppingEvent); // Set the service stop event.
                break;

            default:
                break;
        }
    }
    LeaveCriticalSection(&_csLock);

    return hr;
}

// Starts the service; called from service main.
HRESULT CMpcService::_StartService()
{
    HRESULT hr = S_OK;

    EnterCriticalSection(&_csLock);
    {
        _hServiceStoppingEvent = CreateEvent(nullptr, TRUE /* manual reset */, FALSE /* Initial state non-signalled*/, nullptr);
        hr = (_hServiceStoppingEvent != nullptr) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
        if (SUCCEEDED(hr))
        {
            _hServiceStatus = RegisterServiceCtrlHandlerEx(SERVICE_NAME, s_ServiceHandler, nullptr);
            hr = (_hServiceStatus != nullptr) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
            if (SUCCEEDED(hr))
            {
                // Set the status to start pending, giving a 30 seconds wait hint
                hr = _SetStatus(SERVICE_START_PENDING, 1, 30 * 1000);
                if (SUCCEEDED(hr))
                {
                    // Start our class factories so that COM will know what classes we can create
                    hr = StartFactories();
                    if (SUCCEEDED(hr))
                    {
                        // Now the service is fully up and running.
                        hr = _SetStatus(SERVICE_RUNNING, 0, 0);
                    }
                }
            }
        }
    }
    LeaveCriticalSection(&_csLock);

    return hr;
}

// Stops the service; called from service main.
HRESULT CMpcService::_StopService()
{
    HRESULT hr = S_OK;

    EnterCriticalSection(&_csLock);
    {
        (void)StopFactories();
        hr = _SetStatus(SERVICE_STOPPED, 0, 0);

        if (_hServiceStoppingEvent != nullptr)
        {
            CloseHandle(_hServiceStoppingEvent);
            _hServiceStoppingEvent = nullptr;
        }
    }
    LeaveCriticalSection(&_csLock);

    return hr;
}

//  Service Main thread, it will start the service, then just wait for the stop event to be signaled.
HRESULT CMpcService::ServiceMain()
{
    // Initialize COM for MTA.
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (SUCCEEDED(hr))
    {
        hr = CoInitializeSecurity(nullptr, 
                                  -1, 
                                  nullptr, 
                                  nullptr, 
                                  RPC_C_AUTHN_LEVEL_CONNECT, 
                                  RPC_C_IMP_LEVEL_IMPERSONATE, 
                                  nullptr, 
                                  EOAC_NONE, 
                                  0);
        if (SUCCEEDED(hr))
        {
            hr = _StartService();
            if (SUCCEEDED(hr))
            {
                // Wait for the service stop event to be set.
                (void)WaitForSingleObject(_hServiceStoppingEvent, INFINITE);

                hr = _StopService();
            }
        }
        
        CoUninitialize();
    }

    return hr;
}