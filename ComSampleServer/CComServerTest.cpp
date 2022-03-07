//
// A class that implements IComTest.
//_______________________________________________________________________________

#include "pch.h"
#pragma hdrstop

#include <objbase.h>
#include <shlwapi.h>
#include <assert.h>
#include <IComTest_h.h>
#include "dll.h"
#include "ComSampleServerGuids.h"
#include "ComSampleServerCreateInstances.h"

class CComServerTest : public IComTest
{
public:

    // IUnknown
        
    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }
    
    IFACEMETHODIMP_(ULONG) Release()
    {                                  
        assert(_cRef > 0);
        
        ULONG cRef = InterlockedDecrement(&_cRef);

        if (0 == cRef)
        {
            delete this;
        }

        return cRef;
    }

    IFACEMETHODIMP QueryInterface(__in REFIID riid, __out void **ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CComServerTest, IComTest),
            { 0 },
        };

        return QISearch(this, qit, riid, ppv);
#if 0
        // ANOTHER POSSIBLE IMPLEMENTATION FOR QueryInterface METHOD.
        HRESULT hr = (ppv != nullptr) ? S_OK : E_INVALIDARG;
        if (SUCCEEDED(hr))
        {
            *ppv = nullptr;
            hr = E_NOINTERFACE;

            if (__uuidof(IComTest) == riid)
            {
                *ppv = static_cast<IComTest*>(this);
                hr = S_OK;
            }
            else if (__uuidof(IUnknown) == riid)
            {
                *ppv = static_cast<IUnknown*>(this);
                hr = S_OK;
            }

            if (SUCCEEDED(hr))
            {
                reinterpret_cast<IUnknown*>(*ppv)->AddRef();
            }
        }

        return hr;
#endif
    }

    // IComTest

    IFACEMETHODIMP WhoAmI(_Out_ LPWSTR *ppwszWhoAmI)
    {
        HRESULT hr = (ppwszWhoAmI != nullptr) ? S_OK : E_INVALIDARG;
        if (SUCCEEDED(hr))
        {
            // Create a message string the contains the name of the current process.
            wchar_t wszProcessName[MAX_PATH] = {};
            DWORD dwResult = GetModuleFileNameW(nullptr, wszProcessName, ARRAYSIZE(wszProcessName));
            hr = ((dwResult < ARRAYSIZE(wszProcessName)) && (dwResult != 0)) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
            if (SUCCEEDED(hr))
            {
                LPCWSTR pwszMessagePreface = L"Com Test Server running in process ";
                
                // Prepare the message.
                size_t cchMessagePlusTerminatingNul = wcslen(pwszMessagePreface) + wcslen(wszProcessName) + 1 /* For the terminating NUL character */;
                wchar_t *pwszMessage = (wchar_t *)CoTaskMemAlloc(cchMessagePlusTerminatingNul * sizeof(wchar_t));
                hr = (pwszMessage != nullptr) ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    hr = StringCchPrintfW(pwszMessage, cchMessagePlusTerminatingNul, L"%s%s", pwszMessagePreface, wszProcessName);
                    if (SUCCEEDED(hr))
                    {
                        hr = SHStrDupW(pwszMessage, ppwszWhoAmI);
                    }

                    CoTaskMemFree(pwszMessage);
                }
            }
        }
        
        return hr;
    }
    

public:

    CComServerTest() : _cRef(1)
    {
        InterlockedIncrement(&g_cRefDll); // g_cRefDll++;
    }

private:

    LONG _cRef;
    
    ~CComServerTest(void)
    {
        InterlockedDecrement(&g_cRefDll); // g_cRefDll--;
    }    
};
//_____________________________________________________________________________

//  Creation function
HRESULT CComServerTest_CreateInstance(__in REFIID riid, __out void **ppv)
{
    HRESULT hr = E_OUTOFMEMORY;
    CComServerTest *pInstance = new CComServerTest();
    if (pInstance != nullptr)
    {
        hr = pInstance->QueryInterface(riid, ppv);
        pInstance->Release();
    }

    return hr;
}
//_____________________________________________________________________________