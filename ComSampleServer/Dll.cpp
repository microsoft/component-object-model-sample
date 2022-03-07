//
// Implements the functions requires to register the COM classes.
//_______________________________________________________________________________

#include "pch.h"
#pragma hdrstop

#include <shlwapi.h>
#include <initguid.h>
#include "classfactory.h"
#include "dll.h"
#include "ComSampleServerGuids.h"
#include "ComSampleServerCreateInstances.h"

const int c_nExpectedGuidStringLength = 39;

// This is our AppId to host the out-of-proc components. We'll be using DllHost.exe
// that comes with Windows as a surrogate to host our out-of-proc components.
// We only need the string form since we just need to put this into registry.
PCWSTR APPID_LOCALSERVER = L"{fe0c701d-bd76-4d25-8f9c-d24a5270630d}";

// Every class in this dll should also add an entry of following structure into the global table.
// The code in this file will take care of registration and create instance for it.
struct CLASS_INFO
{
    REFCLSID            rclsid;
    PCWSTR              pwszName;
    PCWSTR              pwszThreadModel;
    PCWSTR              pwszAppId; // The AppId of the host. If you do not want this object to an out-of-proc component, pass an empty string here (L"").
    PFNCREATEINSTANCE   pfnCreateInstance;    
};

//  This is all the classes that are supported by this dll. 
CLASS_INFO   g_Classes[] = 
{
    {
        CLSID_CComServerTest,
        L"COM Server Test",
        L"Both",
        APPID_LOCALSERVER,
        CComServerTest_CreateInstance
    },
};

//  Global variables that help implement IClassFactory.
LONG      g_cServerLocks;
LONG      g_cRefDll;
HINSTANCE g_hInstance;

STDAPI DllGetClassObject(__in REFCLSID rclsid,
                         __in REFIID riid,
                         __deref_out void **ppv)
{
    HRESULT hr = (ppv != nullptr) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        //  Loop through the global object table for matching CLSID, 
        //  then create the class factory instance corresponding to its create function.
        hr = CLASS_E_CLASSNOTAVAILABLE;
        bool fFound = false;
        for (int i = 0; ((i < ARRAYSIZE(g_Classes)) && (!fFound)); ++i)
        {
            if (rclsid == g_Classes[i].rclsid)
            {
                hr = CClassFactory_CreateInstance(g_Classes[i].pfnCreateInstance, riid, ppv);
                fFound = true;
            }
        }
    }

    return hr;
}

STDAPI DllCanUnloadNow()
{
    return ((0 == g_cServerLocks) && (0 == g_cRefDll)) ? S_OK : S_FALSE;
}

// Structure to hold values for registration.
struct REGSTRUCT
{
    HKEY   hRootKey;
    PCWSTR pwszSubKey;
    PCWSTR pwszValueName;
    PCWSTR pwszData;
    DWORD  dwType;
    bool   fAddModulePath;
};

HRESULT _WriteRegistryEntries(__in REGSTRUCT regEntry, __in PCWSTR pwszSubKeyName)
{
    wchar_t wszSubKey[MAX_PATH] = {};
    wchar_t wszData[MAX_PATH] = {};

    // Create the sub key string.
    HRESULT hr = StringCchPrintfW(wszSubKey, ARRAYSIZE(wszSubKey), regEntry.pwszSubKey, pwszSubKeyName);
    if (SUCCEEDED(hr))
    {
        HKEY hKey;
        DWORD dwDisp;
        LONG lResult = RegCreateKeyExW(regEntry.hRootKey,
                                       wszSubKey, 
                                       0, 
                                       nullptr, 
                                       REG_OPTION_NON_VOLATILE,
                                       KEY_WRITE, 
                                       nullptr, 
                                       &hKey, 
                                       &dwDisp);
        hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
        if (SUCCEEDED(hr))
        {
            // Add the module path and name to the string if needed.
            if (regEntry.fAddModulePath)
            {
                // Get the path and module name.
                wchar_t wszModulePathAndName[MAX_PATH] = {};
                DWORD dwResult = GetModuleFileNameW(g_hInstance, wszModulePathAndName, ARRAYSIZE(wszModulePathAndName));
                hr = ((dwResult < ARRAYSIZE(wszModulePathAndName)) && (dwResult != 0)) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
                if (SUCCEEDED(hr))
                {
                    hr = StringCchPrintfW(wszData, 
                                            ARRAYSIZE(wszData), 
                                            (LPWSTR)regEntry.pwszData, 
                                            wszModulePathAndName);
                }
            }
            else
            {
                hr = StringCchCopyW(wszData, 
                                    ARRAYSIZE(wszData), 
                                    (LPWSTR)regEntry.pwszData);
            }
                        
            if (SUCCEEDED(hr))
            {
                lResult = RegSetValueExW(hKey, 
                                            regEntry.pwszValueName, 
                                            0,
                                            regEntry.dwType, 
                                            (LPBYTE)wszData, 
                                            (lstrlenW(wszData) + 1) * sizeof(wchar_t));
                hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
            }
                                   
            (void)RegCloseKey(hKey);
        }
        else
        {
            hr = SELFREG_E_CLASS;
        }
    }

    return hr;
}

HRESULT _RegisterClass(__in REFCLSID rclsid, __in PCWSTR pwszName, __in PCWSTR pwszThreadingModel, __in PCWSTR pwszAppId)
{
    wchar_t wszClassID[64] = {};
    
    // Convert the IID to a string.
    int nResult = StringFromGUID2(rclsid, wszClassID, ARRAYSIZE(wszClassID));
    HRESULT hr = (nResult == c_nExpectedGuidStringLength) ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {       
        // This will setup and register the basic ClassIDs.
        REGSTRUCT rgRegEntries[] = 
        {
            //hRootKey         pwszSubKey  %s->CLSID            pwszValueName       pwszData %s->Module dwType  fAddModulePath
            {HKEY_CLASSES_ROOT, L"CLSID\\%s",                   nullptr,            pwszName,           REG_SZ, false},   
            {HKEY_CLASSES_ROOT, L"CLSID\\%s\\InprocServer32",   nullptr,            L"%s",              REG_SZ, true},
            {HKEY_CLASSES_ROOT, L"CLSID\\%s\\InprocServer32",   L"ThreadingModel",  pwszThreadingModel, REG_SZ, false},

            // The following registry entries are required only if we are registering the type for local server activation hosted by a surrogate.
            {HKEY_CLASSES_ROOT, L"CLSID\\%s",                   L"AppID",           pwszAppId,          REG_SZ, false},
        };

        for (int i = 0; SUCCEEDED(hr) && (i < ARRAYSIZE(rgRegEntries)); ++i)
        {
            hr = _WriteRegistryEntries(rgRegEntries[i], wszClassID);
        }
    }

    return hr;
}

HRESULT _UnregisterClass(__in REFCLSID rclsid)
{
    HRESULT hr = SELFREG_E_CLASS;
    wchar_t wszClassID[MAX_PATH] = {0};
    int  nResult = StringFromGUID2(rclsid, wszClassID, ARRAYSIZE(wszClassID)); // Convert the IID to a string.
    
    // Delete the type's registry entries.
    if (nResult == c_nExpectedGuidStringLength)
    {
        wchar_t wszSubKey[MAX_PATH] = {0};
        hr = StringCchPrintfW(wszSubKey, ARRAYSIZE(wszSubKey), L"CLSID\\%s", wszClassID);
        if (SUCCEEDED(hr))
        {
            LONG lResult = SHDeleteKeyW(HKEY_CLASSES_ROOT, wszSubKey);
            hr = (ERROR_SUCCESS == lResult) ? S_OK : HRESULT_FROM_WIN32(lResult);
        }
    }
    
    return hr;
}

// Registers the AppId that will enable the objects to run as out of proc components inside a surrogate host.
HRESULT _RegisterAppId(__in PCWSTR pwszAppId)
{
    REGSTRUCT rgRegEntries[] = 
    {
        //hRootKey          pwszSubKey %s->AppID  pwszValueName     pwszData             dwType  fAddModulePath
        {HKEY_CLASSES_ROOT, L"AppID\\%s",         nullptr,          L"Com Sample Host",  REG_SZ, false},   
        {HKEY_CLASSES_ROOT, L"AppID\\%s",         L"DllSurrogate",  L"",                 REG_SZ, false}, // Empty value for "DllSurrogate" implies using DllHost.exe as the host.
        {HKEY_CLASSES_ROOT, L"AppID\\%s",         L"RunAs",         L"Interactive User", REG_SZ, false},
    };

    HRESULT hr = S_OK;
    for (int i = 0; SUCCEEDED(hr) && (i < ARRAYSIZE(rgRegEntries)); ++i)
    {
        hr = _WriteRegistryEntries(rgRegEntries[i], pwszAppId);
    }

    return hr;
}

// Unregisters the AppId that will enable the objects to run as out of proc components inside a surrogate host.
HRESULT _UnregisterAppId(__in PCWSTR pwszAppId)
{
    // Build the key AppID\\{...}
    wchar_t szKey[MAX_PATH] = {};
    HRESULT hr = StringCchPrintf(szKey, ARRAYSIZE(szKey), L"AppID\\%s", pwszAppId);
    if (SUCCEEDED(hr))
    {
        // Delete the key
        LONG lResult = SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
        hr = (ERROR_SUCCESS == lResult) ? S_OK : HRESULT_FROM_WIN32(lResult);
    }

    return hr;
}

// Registers the DLL.
// This involves two steps
//  1. Registering all the classes in the DLL.
//  2. Registering the AppId that will enable the objects to run as out of proc components inside a surrogate host.
STDAPI DllRegisterServer()
{
    HRESULT hr = S_OK;

    //  Loop through the global classes table to register each class.
    //  Bail out when we fail to register any of the classes.
    for (int i = 0; ((i < ARRAYSIZE(g_Classes)) && SUCCEEDED(hr)); ++i)
    {
        hr = _RegisterClass(g_Classes[i].rclsid, 
                            g_Classes[i].pwszName,
                            g_Classes[i].pwszThreadModel,
                            g_Classes[i].pwszAppId);
    }

    // Register the AppId that will enable the objects to run as out of proc components inside a surrogate host.
    if (SUCCEEDED(hr))
    {
        hr = _RegisterAppId(APPID_LOCALSERVER);
    }

    // If anything fails at all, unregister whatever registration we might have done so far.
    if (FAILED(hr))
    {
        (void)DllUnregisterServer();
    }

    return hr;
}

// Unregisters the DLL.
// This involves two steps
//  1. Unregistering all the classes in the DLL.
//  2. unregistering the AppId that will enable the objects to run as out of proc components inside a surrogate host.
STDAPI DllUnregisterServer()
{
    HRESULT hr = S_OK;
    
    //  Loop through the global classes table to unregister each class.
    //  Do not bail on failures, continue to unregister whatever we had registered.    
    for (int i = 0; i < ARRAYSIZE(g_Classes); ++i)
    {
        HRESULT hrUnregisterClass = _UnregisterClass(g_Classes[i].rclsid);
        if (FAILED(hrUnregisterClass))
        {
            hr = hrUnregisterClass;
        }
    }

    // Unregister the AppId.
    HRESULT hrUnregisterAppId = _UnregisterAppId(APPID_LOCALSERVER);
    if (FAILED(hrUnregisterAppId))
    {
        hr = hrUnregisterAppId;
    }

    return hr;
}

// Entry point for the DLL.
BOOL APIENTRY DllMain(__in HINSTANCE hModule, __in DWORD  dwReason, __in_opt LPVOID /* lpReserved */)
{
    if (DLL_PROCESS_ATTACH == dwReason)
    {
        g_hInstance = (HINSTANCE)hModule;
        (void)DisableThreadLibraryCalls(g_hInstance);
    }
    
    return TRUE;
}
