#pragma once

#define SERVICE_NAME         L"ComSampleService"
#define SERVICE_DISPLAY_NAME L"COM Sample Service"
#define SERVICE_APPID        L"{2eb6d15c-5239-41cf-82fb-353d20b816b9}"

void WINAPI ServiceMain(__in DWORD dwArgc, __in_ecount(dwArgc) LPWSTR *pszArgv);

HRESULT RegisterClass(__in REFCLSID clsid,
                      __in LPCWSTR pszName,
                      __in LPCWSTR pszAppID);
                      
HRESULT UnregisterClass(__in REFCLSID clsid);

HRESULT RegisterAppID(__in LPCWSTR pszAppID,
                      __in LPCWSTR pszName,
                      __in LPCWSTR pszServiceName);

HRESULT UnregisterAppID(__in LPCWSTR pszAppID);

HRESULT RegisterService(__in LPCWSTR pszServiceName,
                        __in LPCWSTR pszServiceDisplayName);

HRESULT UnregisterService(__in LPCWSTR pszServiceName);

HRESULT StartFactories();

HRESULT StopFactories();