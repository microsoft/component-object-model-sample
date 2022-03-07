//
// Takes care of registering the COM Servers.
//_______________________________________________________________________________

#include "pch.h"
#pragma hdrstop
#include <initguid.h>
#include "ClassFactory.h"
#include "Infrastructure.h"
#include "ComSampleServiceGuids.h"
#include "ComSampleServiceCreateInstances.h"

//  Every class in this exe should also add an entry of the following structure
//  into the global table. The the code in this file will take care of the
//  registration and create instance for it.
struct CLASS_INFO
{
    REFCLSID            rclsid;
    PCWSTR              pszName;
    PFNCREATEINSTANCE   pfnCreateInstance;
    IUnknown*           pUnknown;   // Save the class factory pointer when exe starts
    DWORD               dwRegister; // Save the registration cookie for this class
};

//  This is the list of all the classes that are supported by this COM server.
CLASS_INFO   g_Classes[] = 
{
    {
        CLSID_CComServiceTest,
        L"COM Test Service Object",
        CComServiceTest_CreateInstance,
        nullptr,
        0
    },
};

//  Starts factories for the classes in this exe.
//  This is needed for the exe to tell COM that we are ready to serve client calls.
HRESULT StartFactories()
{
    HRESULT hr = S_OK;
    for (int i = 0; (i < ARRAYSIZE(g_Classes)) && SUCCEEDED(hr); ++i)
    {        
        hr = CClassFactory_CreateInstance(g_Classes[i].pfnCreateInstance,
                                          IID_IUnknown,
                                          (LPVOID*) &g_Classes[i].pUnknown);
        if (SUCCEEDED(hr))
        {
            hr = CoRegisterClassObject(g_Classes[i].rclsid,
                                       g_Classes[i].pUnknown,
                                       CLSCTX_LOCAL_SERVER,
                                       REGCLS_MULTIPLEUSE,
                                       &g_Classes[i].dwRegister);

            if (FAILED(hr))
            {
                g_Classes[i].pUnknown->Release();
                g_Classes[i].pUnknown = nullptr;
            }
        }
    }

    if (FAILED(hr))
    {
        (void)StopFactories();
    }

    return hr; 
}

HRESULT StopFactories()
{
    HRESULT hr = S_OK;
    for (int i = 0; i < ARRAYSIZE(g_Classes); ++i)
    {        
        if (g_Classes[i].dwRegister != 0)
        {
            HRESULT hrTemp = CoRevokeClassObject(g_Classes[i].dwRegister);
            if (FAILED(hrTemp))
            {
                hr = hrTemp;
            }
        }

        if (g_Classes[i].pUnknown != nullptr)
        {
            g_Classes[i].pUnknown->Release();
            g_Classes[i].pUnknown = nullptr;
        }
    }

    return hr; 
}

//  Unregisters all the stuff we have registered.
HRESULT ExeUnregisterServer()
{
    HRESULT hr = S_OK;

    //  Unregister service 
    HRESULT hrTemp = UnregisterService(SERVICE_NAME);
    if (FAILED(hrTemp))
    {
        hr = hrTemp;
        printf("Failed to unregister service, hr = %08X\n", hr);
    }

    //  Unregister AppID 
    hrTemp = UnregisterAppID(SERVICE_APPID);
    if (FAILED(hrTemp))
    {
        hr = hrTemp;
        printf("Failed to unregister AppID, hr = %08X\n", hr);
    }
    
    //  Loop through the global classes table to unregister each class.
    //  Do not bail on failures, continue to unregister whatever we had registered.
    for (int i = 0; i < ARRAYSIZE(g_Classes); ++i)
    {
        hrTemp = UnregisterClass(g_Classes[i].rclsid);
        if (FAILED(hrTemp))
        {
            hr = hrTemp;
            printf("Failed to unregister class %ls, hr = %08X\n", g_Classes[i].pszName, hr);
        }
    }

    return hr;
}

//  Registers different parts that needed for our service, it includes:
//  1. The service information itself
//  2. AppID for connecting the service to a COM app
//  3. Classes that implemented in this exe, and connect them with the AppID
HRESULT ExeRegisterServer()
{
    //  Register our service first
    HRESULT hr = RegisterService(SERVICE_NAME, SERVICE_DISPLAY_NAME);
    if (SUCCEEDED(hr))
    {
        //  Register the AppID
        hr = RegisterAppID(SERVICE_APPID, SERVICE_DISPLAY_NAME, SERVICE_NAME);
        if (SUCCEEDED(hr))
        {
            //  Loop through the global classes table to register each class.
            //  Bail if we fail to register any of the classes.
            for (int i = 0; (i < ARRAYSIZE(g_Classes)) && SUCCEEDED(hr); ++i)
            {
                hr = RegisterClass(g_Classes[i].rclsid, 
                                   g_Classes[i].pszName,
                                   SERVICE_APPID);
                if (FAILED(hr))
                {
                    printf("Failed to register class %ls, hr = %08X\n", g_Classes[i].pszName, hr);
                }
            }
        }
        else
        {
            printf("Failed to register AppID, hr = %08X\n", hr);
        }
    }
    else
    {
        printf("Failed to register service, hr = %08X\n", hr);
    }

    // Unregister everything if any of the registration process fails.
    if (FAILED(hr))
    {
        (void)ExeUnregisterServer();
    }

    return hr;
}

void PrintUsage()
{
    printf("Usage: <This EXE> [/RegisterServer][/UnregisterServer]\n");
}

int __cdecl wmain(int argc, __in wchar_t * argv [])
{
    // Parse the command line, several cases:
    //
    // 1. User runs this exe with RegisterServer/UnregisterServer to register/unregister it
    // 2. COM activate this exe with "Embedding" option, 
    // 3. SCM launches this exe from a service control program such as "net start <ServiceName>"
    // 4. User launches this exe directly. 
    
    HRESULT hr = E_INVALIDARG;
    bool fStartService = false;
    if (argc > 1)
    {        
        wchar_t szTokens[] = L"-/";
        LPWSTR pszNextToken;
        LPWSTR pszToken = wcstok_s(argv[1], szTokens, &pszNextToken);
        if (pszToken != nullptr)
        {
            if (_wcsicmp(pszToken, L"RegisterServer") == 0)
            {
                hr = ExeRegisterServer();
                printf("RegisterServer done, hr = %08X\n", hr);
            }
            else if (_wcsicmp(pszToken, L"UnregisterServer") == 0)
            {
                hr = ExeUnregisterServer();
                printf("UnregisterServer done, hr = %08X\n", hr);
            }
            else if (_wcsicmp(pszToken, L"Embedding") == 0)
            {
                // COM started us, we will start our service.
                hr = S_OK;
                fStartService = true;
            }
            else
            {
                printf("Unknown switch.\n");
                PrintUsage();
            }
        }
        else
        {
            printf("Unknown parameter.\n");
            PrintUsage();
        }
    }    
    else
    {
        // User launches our exe, or SCM launches us. We will start our service.
        hr = S_OK;
        fStartService = true;
    }

    // Ask SCM to start our service.
    if (fStartService)
    {
        SERVICE_TABLE_ENTRY rgServiceTable[2] = {0};   // Only need 1 entry and then the trailing nullptr.
        rgServiceTable[0].lpServiceName = SERVICE_NAME;
        rgServiceTable[0].lpServiceProc = (LPSERVICE_MAIN_FUNCTION)ServiceMain;
        hr = StartServiceCtrlDispatcher(rgServiceTable) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    return (SUCCEEDED(hr) ? 0 : hr);
}
