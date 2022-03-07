#include <windows.h>
#include <strsafe.h>
#include <initguid.h>
#include "IComTest_h.h"
#include <ComSampleServerGuids.h>
#include <ComSampleServiceGuids.h>
#include <stdio.h>

HRESULT Execute(_In_ const IID &rclsid, _In_ DWORD dwCoInit, _In_ DWORD dwClsContext)
{
    HRESULT hr = CoInitializeEx(nullptr, dwCoInit);
    if (SUCCEEDED(hr))
    {
        IComTest *pComTest;
        hr = CoCreateInstance(rclsid,
                              nullptr,
                              dwClsContext,
                              IID_PPV_ARGS(&pComTest));
        if (SUCCEEDED(hr))
        {
            LPWSTR pwszWhoAmI;
            hr = pComTest->WhoAmI(&pwszWhoAmI);
            if (SUCCEEDED(hr))
            {
                wprintf(L"%s. Client calling from %s. COM Server running %s.\n", 
                        pwszWhoAmI,
                        (dwCoInit == COINIT_MULTITHREADED) ? L"MTA" : L"STA",
                        (dwClsContext == CLSCTX_INPROC_SERVER) ? L"in-process" : L"out-of-process");
                CoTaskMemFree(pwszWhoAmI);
            }

            pComTest->Release();
        }

        CoUninitialize();
    }

    return hr;
}

int __cdecl main()
{
    Execute(CLSID_CComServerTest,  COINIT_MULTITHREADED,     CLSCTX_INPROC_SERVER);
    Execute(CLSID_CComServerTest,  COINIT_MULTITHREADED,     CLSCTX_LOCAL_SERVER);
    Execute(CLSID_CComServerTest,  COINIT_APARTMENTTHREADED, CLSCTX_INPROC_SERVER);
    Execute(CLSID_CComServerTest,  COINIT_APARTMENTTHREADED, CLSCTX_LOCAL_SERVER);
    Execute(CLSID_CComServiceTest, COINIT_MULTITHREADED,     CLSCTX_LOCAL_SERVER);
    Execute(CLSID_CComServiceTest, COINIT_APARTMENTTHREADED, CLSCTX_LOCAL_SERVER);

    return 0;
}