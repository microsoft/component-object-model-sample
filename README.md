# Component Object Model (COM) Sample
COM is a very powerful technology to componentize software based on object oriented design. Please see [documentation here](https://docs.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal "documentation here") for more details.

One of the major drawbacks of COM is the initial boilerplate required to set up the COM components. [ATL](https://docs.microsoft.com/en-us/cpp/atl/active-template-library-atl-concepts?view=msvc-170 "ATL") is a set of libraries that help build the boilerplate, but comes with its own complexity.

This sample here provides the skeletal code that would do the heavylifting of COM setup and registration so that developers can focus on the business logic alone rather than worrying about the infrastructure. It does **not use ATL**. Instead it uses simple plain C++ code so that developers can understand and debug the underlying skeleton if required.

## Terms used
- A COM Server is an object that provides the business logic.
- A COM Client is the code that access the COM Server through any interface exposed by the COM Server.
- The interaction between the COM Client and COM Server happens via marshaling, and requires a Proxy-Stub DLL.
- Please see [documentation here](https://docs.microsoft.com/en-us/windows/win32/com/com-clients-and-servers "documentation here") for more details.

## This project
Here is the summary about the various directories in this sample.
- **ComSampleProxy**: ProxyStub DLL. To add a new interface, simply add a new IDL file to the project.
- **COMSampleServer**: Support for activating the COM Server in-process as well as out-of-process (in an OS provided [DLL Surrogate](https://docs.microsoft.com/en-us/windows/win32/com/dll-surrogates "DLL Surrogate") namely DllHost.exe).
- **COMSampleService**: Support for activating the COM Server in a LocalSystem Service.
- **ComSampleClient**: A sample COM Client that calls into and tests the COM Servers mentioned above.

## Creating and registering a new COM Component
Creating a new COM component is super easy with this model.
1. **Add your interface**: 
 - Add a new IDL file with your interface to ComSampleProxy project. See sample file /ComSampleProxy/IComTest.idl.
2. **Implement your COM class**:
 - *For in-process activation or out-of-process in surrogate activatation*: Add your class implementation to ComSampleServer project. See example ComSampleServer/CComServerTest.cpp. Now go to ComSampleServer/Dll.cpp and simply add your class entries to "g_Classes".
 - *For out-of-process activation in a LocalSystem service*: Add your class implementation to ComSampleService project. See example ComSampleService/CComServiceTest.cpp. Now go to ComSampleService/Main.cpp and simply add your class entries to "g_Classes".
3. **Register your COM components**:
 - Copy ComSampleProxy.dll onto your target machine. From an elevated prompt, run: ***regsvr32 ComSampleProxy.dll***.
 - *For in-process activation or out-of-process in surrogate activatation*: Copy ComSampleServer.dll onto your target machine. From an elevated prompt, run: ***regsvr32 ComSampleServer.dll***.
 - *For out-of-process activation in a LocalSystem service*: Copy ComSampleService.exe onto your target machine. From an elevated command prompt, run: ***ComSampleService.exe /RegisterServer***.

That is it!

## **Trademarks**
This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.
