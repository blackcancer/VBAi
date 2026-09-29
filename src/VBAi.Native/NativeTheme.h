#pragma once
#include <windows.h>
struct NativeThemeStatus {
    DWORD size, abi, active, windows, imports, restoredImports;
    DWORD patterns, images, text, fills, unsupported, failures;
};
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeStart(HWND editor);
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeRegister(HWND window);
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeRefresh();
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeStop();
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeStatus(NativeThemeStatus* status);
