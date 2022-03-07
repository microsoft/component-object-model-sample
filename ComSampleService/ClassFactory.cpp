//
//  Shared class factory for helping create instances of the objects.
//
//  COM requires each class has its own class factory to create instances,
//  but lots of classes actually can share the same class factory implementation.
//
//  What it does is simply remember the class's creation function pointer, then
//  when IClassFactory::CreateInstance is called, it calls the creation function.
//_______________________________________________________________________________

#include "pch.h"
#pragma hdrstop

#include "classfactory.h"

class CClassFactory : public IClassFactory
{
public:

    // IUnknown methods:
    
    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }
    
    IFACEMETHODIMP_(ULONG) Release()
    { 
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
            QITABENT(CClassFactory, IClassFactory),
            { 0 },
        };

        return QISearch(this, qit, riid, ppv);
    }

    // IClassFactory methods:
    
    IFACEMETHODIMP LockServer(__in BOOL /* fLock */)
    {
        // Since we are a service, SCM controls the lifetime of our process,
        // Client does not have control over us. 
        return S_OK;
    }

    IFACEMETHODIMP CreateInstance(__in_opt IUnknown *punkOuter, __in REFIID riid, __out void **ppv)
    {
        HRESULT hr = (ppv != nullptr) ? S_OK : E_INVALIDARG;
        if (SUCCEEDED(hr))
        {
            *ppv = nullptr;
            hr = (punkOuter == nullptr) ? S_OK : CLASS_E_NOAGGREGATION;
            if (SUCCEEDED(hr))
            {
                hr = (_pfnCreateInstance != nullptr) ? S_OK : E_POINTER;
                if (SUCCEEDED(hr))
                {
                    hr = _pfnCreateInstance(riid, ppv);
                }
            }
        }

        return hr;
    }

public:

    CClassFactory(__in PFNCREATEINSTANCE pfnCreateInstance)
        : _cRef(1),
          _pfnCreateInstance(pfnCreateInstance)
    {
    }

private:

    LONG                _cRef;
    PFNCREATEINSTANCE   _pfnCreateInstance;

    ~CClassFactory()
    {
    }

};
//_____________________________________________________________________________

//  Our own creation function
HRESULT CClassFactory_CreateInstance(__in PFNCREATEINSTANCE pfnCreateInstance, __in REFIID riid, __out void **ppv)
{
    CClassFactory *pClassFactory = new CClassFactory(pfnCreateInstance);
    HRESULT hr = (pClassFactory != nullptr) ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pClassFactory->QueryInterface(riid, ppv);
        pClassFactory->Release();
    }

    return hr;
}
//_____________________________________________________________________________