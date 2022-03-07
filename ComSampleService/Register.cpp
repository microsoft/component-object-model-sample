//
// Takes care of registering the LocalSystem service.
//_______________________________________________________________________________

#include "pch.h"
#pragma hdrstop
#include "Infrastructure.h"

//  Registers a class.
// Since we implement the various classes in the service, we don't need InProcServer32 or LocalServer32 subkeys.
HRESULT RegisterClass(__in REFCLSID clsid,
                      __in LPCWSTR pszName,
                      __in LPCWSTR pszAppID)
{
    LPOLESTR pszCLSID;
    HRESULT hr = StringFromCLSID(clsid, &pszCLSID);
    if (SUCCEEDED(hr))
    {
        wchar_t szKey[MAX_PATH] = {};
        hr = StringCchPrintf(szKey, ARRAYSIZE(szKey), L"CLSID\\%s", pszCLSID);
        if (SUCCEEDED(hr))
        {
            // Create CLSID key
            HKEY hKey;
            LONG lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT,
                                          szKey,
                                          0,
                                          nullptr,
                                          REG_OPTION_NON_VOLATILE,
                                          KEY_WRITE,
                                          nullptr,
                                          &hKey,
                                          nullptr);
            hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
            if (SUCCEEDED(hr))
            {
                // Set Name
                DWORD dwSize = (lstrlen(pszName) + 1) * sizeof(wchar_t);
                lResult = RegSetValueEx(hKey,
                                        nullptr,
                                        0,
                                        REG_SZ,
                                        (BYTE *)pszName,
                                        dwSize);
                hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
                if (SUCCEEDED(hr))
                {
                    // Set AppID 
                    dwSize = (lstrlen(pszAppID) + 1) * sizeof(wchar_t);
                    lResult = RegSetValueEx(hKey,
                                            L"AppID",
                                            0, 
                                            REG_SZ,
                                            (BYTE *) pszAppID,
                                            dwSize);
                    hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
                }

                (void)RegCloseKey(hKey);
            }
        }

        CoTaskMemFree(pszCLSID);
    }

    return hr;
}

HRESULT UnregisterClass(__in REFCLSID clsid)
{
    LPOLESTR pszCLSID;
    HRESULT hr = StringFromCLSID(clsid, &pszCLSID);
    if (SUCCEEDED(hr))
    {
        wchar_t szKey[MAX_PATH] = {};
        hr = StringCchPrintf(szKey, ARRAYSIZE(szKey), L"CLSID\\%s", pszCLSID);
        if (SUCCEEDED(hr))
        {
            LSTATUS lResult = SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
            hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
        }

        CoTaskMemFree(pszCLSID);
    }

    return hr;
}

//  Registers AppID for the COM server. Since we are a service, we only need the service name in the LocalService value. 
HRESULT RegisterAppID(__in LPCWSTR pszAppID,
                      __in LPCWSTR pszName,
                      __in LPCWSTR pszServiceName)
{
    wchar_t szKey[MAX_PATH] = {};
    HRESULT hr = StringCchPrintf(szKey, ARRAYSIZE(szKey), L"AppID\\%s", pszAppID);
    if (SUCCEEDED(hr))
    {
        HKEY hKey;
        LONG lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT,
                                      szKey,
                                      0,
                                      nullptr,
                                      REG_OPTION_NON_VOLATILE,
                                      KEY_WRITE,
                                      nullptr,
                                      &hKey,
                                      nullptr);
        hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
        if (SUCCEEDED(hr))
        {
            // Set Name
            DWORD dwSize = (lstrlen(pszName) + 1) * sizeof(wchar_t);
            lResult = RegSetValueEx(hKey,
                                    nullptr,
                                    0,
                                    REG_SZ,
                                    (BYTE *)pszName,
                                    dwSize);
            hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
            if (SUCCEEDED(hr))
            {
                // Set the LocalService value, this will let COM knows that the server is a service
                dwSize = (lstrlen(pszServiceName) +1 ) * sizeof(wchar_t);
                lResult = RegSetValueEx(hKey,
                                        L"LocalService",
                                        0, 
                                        REG_SZ,
                                        (BYTE *) pszServiceName,
                                        dwSize);
                hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
            }

            (void)RegCloseKey(hKey);
        }
    }

    return hr;
}

HRESULT UnregisterAppID(__in LPCWSTR pszAppID)
{
    wchar_t szKey[MAX_PATH] = {};
    HRESULT hr = StringCchPrintf(szKey, ARRAYSIZE(szKey), L"AppID\\%s", pszAppID);
    if (SUCCEEDED(hr))
    {
        LSTATUS lResult = SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
        hr = (lResult == ERROR_SUCCESS) ? S_OK : HRESULT_FROM_WIN32(lResult);
    }

    return hr;
}

// Creates this service by calling SCM APIs.
// Return values:
//      S_OK - if the service was successfully created.
//      S_FALSE - if the service already exists.
//      Standard error codes in case of failure.
HRESULT RegisterService(__in LPCWSTR pszServiceName,
                        __in LPCWSTR pszServiceDisplayName)
{
    SC_HANDLE hSCM = OpenSCManager(nullptr, nullptr, SC_MANAGER_CREATE_SERVICE);
    HRESULT hr = (hSCM != nullptr) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    if (SUCCEEDED(hr))
    {
        // Set service exe path
        wchar_t szPath[MAX_PATH] = {};
        GetModuleFileName(nullptr, szPath, ARRAYSIZE(szPath));
        SC_HANDLE hSvc = CreateService(hSCM, 
                                       pszServiceName,
                                       pszServiceDisplayName,
                                       SC_MANAGER_CREATE_SERVICE, // DWORD dwDesiredAccess,
                                       SERVICE_WIN32_OWN_PROCESS, // DWORD dwServiceType,
                                       SERVICE_DEMAND_START,      // DWORD dwStartType,
                                       SERVICE_ERROR_NORMAL,      // DWORD dwErrorControl,
                                       szPath,                    // LPCTSTR lpBinaryPathName,
                                       nullptr,                   // LPCTSTR lpLoadOrderGroup,
                                       nullptr,                   // LPDWORD lpdwTagId,
                                       nullptr,                   // LPCTSTR lpDependencies,
                                       nullptr,                   // LPCTSTR lpServiceStartName, nullptr means LocalSystem
                                       nullptr);                  // LPCTSTR lpPassword
        hr = (hSvc != nullptr) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
        if (SUCCEEDED(hr))
        {
            (void)CloseServiceHandle(hSvc);
        }
        else if (hr == HRESULT_FROM_WIN32(ERROR_SERVICE_EXISTS))
        {
            hr = S_FALSE;   // Success code to indicate that the service already exists.
        }

        (void)CloseServiceHandle(hSCM);
    }

    return hr;
}

// Deletes this service.
// See MSDN: Deleting a service: http://msdn.microsoft.com/library/ms682571(v=VS.85).aspx
HRESULT UnregisterService(__in LPCWSTR pszServiceName)
{
    SC_HANDLE hSCM = OpenSCManager(nullptr, nullptr, SC_MANAGER_ALL_ACCESS);
    HRESULT hr = (hSCM != nullptr) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    if (SUCCEEDED(hr))
    {
        SC_HANDLE hSvc = OpenService(hSCM, pszServiceName, DELETE);
        hr = (hSvc != nullptr) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
        if (SUCCEEDED(hr))
        {
            hr = DeleteService(hSvc) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
            (void)CloseServiceHandle(hSvc);
        }
        
        (void)CloseServiceHandle(hSCM);
    }

    return hr;
}