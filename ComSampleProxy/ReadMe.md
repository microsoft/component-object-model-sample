# In case you want to create a new Proxy DLL from Scratch:
- Create a new Win32 empty project the builds into a DLL.
- Add the IDL file and define the interface.
- Add a definition file and export the following methods:
    DllCanUnloadNow, DllGetClassObject, GetProxyDllInfo, DllRegisterServer, DllUnregisterServer.
- Compile the project. The compilation would fail, but it would generate *_h.h, *_i.c and *_p.c files. Add them to the project.
- In the project properties -> Configuration properties -> C/C++ -> Command line, add "-DREGISTER_PROXY_DLL ". This will generate 
    DllMain, DllRegisterServer, and DllUnregisterServer functions for automatically registering a proxy DLL. Please read the
    comments at the top of the file named "rpcproxy.h" included in dlldata.c
- In the project properties -> Configuration properties -> Linker -> Input, add the following static libraries as well: rpcns4.lib;rpcrt4.lib.
    A lot of helper methods like __imp__NdrDllGetClassObject@24, __imp__NdrDllCanUnloadNow@4 referenced in
    DllCanUnloadNow, DllRegisterServer etc are defined in these LIB files.
- In the project properties -> Configuration properties -> C/C++ -> Code Generation, change "Runtime Library" type to 
    Multi-threaded (/MT) or Multi-threaded Debug (/MTd). The DLL registration would fail otherwise complaining about missing
    dependencies (which would require redistributing the visual studio binaries like msvcr100.dll along with our binary).
- Due to a bug in Visual Studio 2010, 64-bit compilation gives a "unresolved external symbol ###_ProxyFileInfo" error.
    To work around it, change the following settings for the project (by right clicking and bringing up the property pages)
    "MIDL -> Target Environment" and set it to "Microsoft Windows 64-bit on x64 (/env x64)". Please note that this needs to be done only for 64-bit. 32-bit should work fine with the default settings. 