#pragma once
#include <windows.h>

// Start and stop on the editor window's owning thread, inside its process.
// The library pins itself before installing hooks; do not unload it.
extern "C" __declspec(dllexport) DWORD __cdecl VBAiTraceStart(
    HWND editorRoot, const wchar_t* outputJsonl, DWORD maxEvents);
// Observe the module that registered a descendant control's window class.
extern "C" __declspec(dllexport) DWORD __cdecl VBAiTraceStartForWindow(
    HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents);
// Explicit experimental rendering mode, confined to one MsoCommandBar HWND.
extern "C" __declspec(dllexport) DWORD __cdecl VBAiToolbarPatternStart(
    HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents);
extern "C" __declspec(dllexport) DWORD __cdecl VBAiTraceStop();
extern "C" __declspec(dllexport) DWORD __cdecl VBAiTraceCount();
extern "C" __declspec(dllexport) DWORD __cdecl VBAiTraceDropped();
