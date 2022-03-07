// ClassFactory.h
//
//  Shared class factory for helping create instances of the objects.
//
//  COM requires each class has its own class factory to create instances,
//  but lots of classes actually can share the same class factory implementation.
//
//  What it does is simply remember the class's creation function pointer, then
//  when IClassFactory::CreateInstance is called, it calls the creation function.
//_______________________________________________________________________________


#pragma once

//
//  Every object in this dll should implement one copy of this functions so that it can be called
//  by the class factory instance.
//
//  Suggested name is "ObjName_CreateInstance"
//
typedef HRESULT (*PFNCREATEINSTANCE)(__in REFIID riid, __out void **ppv);

//
//  Our own creation function, we take one additional parameter pfnCreateInstance so that we can
//  call the specific class's creation function later when IClassFactory::CreateInstance is called.
//
HRESULT CClassFactory_CreateInstance(__in PFNCREATEINSTANCE pfnCreateInstance, __in REFIID riid, __out void **ppv);