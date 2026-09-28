#pragma once
#include <windows.h>

// Start and stop on the editor window's owning thread, inside its process.
// The library pins itself before installing hooks; do not unload it.
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeTraceStart(
    HWND editorRoot, const wchar_t* outputJsonl, DWORD maxEvents);
// Observe the module that registered a descendant control's window class.
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeTraceStartForWindow(
    HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents);
// Explicit experimental rendering mode, confined to one MsoCommandBar HWND.
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeToolbarPatternStart(
    HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents);
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeTraceStop();
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeTraceCount();
extern "C" __declspec(dllexport) DWORD __cdecl CodexVbeTraceDropped();
