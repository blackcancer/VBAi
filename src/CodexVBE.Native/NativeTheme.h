#pragma once
#include <windows.h>
struct NativeThemeStatus {
    DWORD size, abi, active, windows, imports, restoredImports;
    DWORD patterns, images, text, fills, unsupported, failures;
};
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeThemeStart(HWND editor);
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeThemeRegister(HWND window);
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeThemeRefresh();
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeThemeStop();
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeThemeStatus(NativeThemeStatus* status);
